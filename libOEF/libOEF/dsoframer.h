/***************************************************************************
 * DSOFRAMER.H
 *
 * Developer Support Office ActiveX Document Framer Control Sample
 *
 *  Copyright ?999-2004; Microsoft Corporation. All rights reserved.
 *  Written by Microsoft Developer Support Office Integration (PSS DSOI)
 * 
 *  This code is provided via KB 311765 as a sample. It is not a formal
 *  product and has not been tested with all containers or servers. Use it
 *  for educational purposes only.
 *
 *  You have a royalty-free right to use, modify, reproduce and distribute
 *  this sample application, and/or any modified version, in any way you
 *  find useful, provided that you agree that Microsoft has no warranty,
 *  obligations or liability for the code or information provided herein.
 *
 *  THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
 *  EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
 *  WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 *
 *  See the EULA.TXT file included in the KB download for full terms of use
 *  and restrictions. You should consult documentation on MSDN Library for
 *  possible updates or changes to behaviors or interfaces used in this sample.
 *
 ***************************************************************************/
#ifndef DS_DSOFRAMER_H 
#define DS_DSOFRAMER_H

////////////////////////////////////////////////////////////////////
// We compile at level 4 and disable some unnecessary warnings...
//
#pragma warning(push, 4) // Compile at level-4 warnings
#pragma warning(disable: 4100) // unreferenced formal parameter (in OLE this is common)
#pragma warning(disable: 4146) // unary minus operator applied to unsigned type, result still unsigned
#pragma warning(disable: 4268) // const static/global data initialized with compiler generated default constructor
#pragma warning(disable: 4310) // cast truncates constant value
#pragma warning(disable: 4786) // identifier was truncated in the debug information

////////////////////////////////////////////////////////////////////
// Compile Options For Modified Behavior...
//
#define DSO_MSDAIPP_USE_DAVONLY          // Default to WebDAV protocol for open by HTTP/HTTPS
#define DSO_WORD12_PERSIST_BUG           // Perform workaround for IPersistFile bug in Word 2007

////////////////////////////////////////////////////////////////////
// Needed include files (both standard and custom)
//
#include <windows.h>
#include <ole2.h>
#include <olectl.h>
#include <oleidl.h>
#include <objsafe.h>

#include "utilities.h"
#include "dsofdocobj.h"


////////////////////////////////////////////////////////////////////
// Global Variables
//
extern HINSTANCE        v_hModule;
//extern CRITICAL_SECTION v_csecThreadSynch;
extern HICON            v_icoOffDocIcon;
extern ULONG            v_cLocks;
//extern BOOL             v_fUnicodeAPI;
extern BOOL             v_fWindows2KPlus;

////////////////////////////////////////////////////////////////////
// Custom Errors - we support a very limited set of custom error messages
//
#define DSO_E_ERR_BASE              0x80041100
#define DSO_E_UNKNOWN               0x80041101   // "An unknown problem has occurred."
#define DSO_E_INVALIDPROGID         0x80041102   // "The ProgID/Template could not be found or is not associated with a COM server."
#define DSO_E_INVALIDSERVER         0x80041103   // "The associated COM server does not support ActiveX Document embedding."
#define DSO_E_COMMANDNOTSUPPORTED   0x80041104   // "The command is not supported by the document server."
#define DSO_E_DOCUMENTREADONLY      0x80041105   // "Unable to perform action because document was opened in read-only mode."
#define DSO_E_REQUIRESMSDAIPP       0x80041106   // "The Microsoft Internet Publishing Provider is not installed, so the URL document cannot be open for write access."
#define DSO_E_DOCUMENTNOTOPEN       0x80041107   // "No document is open to perform the operation requested."
#define DSO_E_INMODALSTATE          0x80041108   // "Cannot access document when in modal condition."
#define DSO_E_NOTBEENSAVED          0x80041109   // "Cannot Save file without a file path."
#define DSO_E_FRAMEHOOKFAILED       0x8004110A   // "Unable to set frame hook for the parent window."
#define DSO_E_ERR_MAX               0x8004110B

////////////////////////////////////////////////////////////////////
// Custom OLE Command IDs - we use for special tasks
//
#define OLECMDID_GETDATAFORMAT      0x7001  // 28673
#define OLECMDID_SETDATAFORMAT      0x7002  // 28674
#define OLECMDID_LOCKSERVER         0x7003  // 28675
#define OLECMDID_RESETFRAMEHOOK     0x7009  // 28681
#define OLECMDID_NOTIFYACTIVE       0x700A  // 28682

////////////////////////////////////////////////////////////////////
// Custom Window Messages (only apply to CDsoFramerControl window proc)
//
#define DSO_WM_ASYNCH_OLECOMMAND         (WM_USER + 300)
#define DSO_WM_ASYNCH_STATECHANGE        (WM_USER + 301)

#define DSO_WM_HOOK_NOTIFY_COMPACTIVE    (WM_USER + 400)
#define DSO_WM_HOOK_NOTIFY_APPACTIVATE   (WM_USER + 401)
#define DSO_WM_HOOK_NOTIFY_FOCUSCHANGE   (WM_USER + 402)
#define DSO_WM_HOOK_NOTIFY_SYNCPAINT     (WM_USER + 403)
#define DSO_WM_HOOK_NOTIFY_PALETTECHANGE (WM_USER + 404)

// State Flags for DSO_WM_ASYNCH_STATECHANGE:
#define DSO_STATE_MODAL            1
#define DSO_STATE_ACTIVATION       2
#define DSO_STATE_INTERACTIVE      3
#define DSO_STATE_RETURNFROMMODAL  4


////////////////////////////////////////////////////////////////////
// Menu Bar Items
//
#define DSO_MAX_MENUITEMS         16
#define DSO_MAX_MENUNAME_LENGTH   32

#ifndef DT_HIDEPREFIX
#define DT_HIDEPREFIX             0x00100000
#define DT_PREFIXONLY             0x00200000
#endif

#define SYNCPAINT_TIMER_ID         4


////////////////////////////////////////////////////////////////////
// CDsoFramerControl -- Main Control (OCX) Object 
//
//  The CDsoFramerControl control is standard OLE control designed around 
//  the OCX94 specification. Because we plan on doing custom integration to 
//  act as both OLE object and OLE host, it does not use frameworks like ATL 
//  or MFC which would only complicate the nature of the sample.
//
//  The control inherits from its automation interface, but uses nested 
//  classes for all OLE interfaces. This is not a requirement but does help
//  to clearly seperate the tasks done by each interface and makes finding 
//  ref count problems easier to spot since each interface carries its own
//  counter and will assert (in debug) if interface is over or under released.
//  
//  The control is basically a stage for the ActiveDocument embedding, and 
//  handles any external (user) commands. The task of actually acting as
//  a DocObject host is done in the site object CDsoDocObject, which this 
//  class creates and uses for the embedding.
//
class CDsoFramerControl : public IUnknown
{
public:
	CDsoFramerControl();
    ~CDsoFramerControl(void);

	// IUnknown Implementation
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) { return S_OK; }
	STDMETHODIMP_(ULONG) AddRef(void) { return 0; }
	STDMETHODIMP_(ULONG) Release(void) { return 0; }


 // _FramerControl Implementation
    STDMETHODIMP Activate();
	HRESULT Open(LPWSTR pwszDocument, BOOL fOpenReadOnly, LPWSTR pwszAltProgId, HWND hwndParent, RECT dstRect);

	void    OnResize(RECT dstRect);


    STDMETHODIMP Save(VARIANT SaveAsDocument, VARIANT OverwriteExisting);
	HRESULT Close();

 // IDsoDocObjectSite Implementation (for DocObject Callbacks to control)
    BEGIN_INTERFACE_PART(DsoDocObjectSite, IDsoDocObjectSite)
        STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);
        STDMETHODIMP GetWindow(HWND* phWnd);
        STDMETHODIMP GetBorder(LPRECT prcBorder);
        STDMETHODIMP GetHostName(LPWSTR *ppwszHostName);
        STDMETHODIMP SysMenuCommand(UINT uiCharCode);
        STDMETHODIMP SetStatusText(LPCOLESTR pszText);
    END_INTERFACE_PART(DsoDocObjectSite)

    STDMETHODIMP           InitializeNewInstance();

    STDMETHODIMP           InPlaceActivate(LONG lVerb);
    STDMETHODIMP           UIActivate(BOOL fForceUIActive);
    void    SetInPlaceVisible(BOOL fShow);
    void    UpdateModalState(BOOL fModeless, BOOL fNotifyIPObject);
    void    UpdateInteractiveState(BOOL fActive);
    void    EnableDropFile(BOOL fEnable);

    void    OnDraw(DWORD dvAspect, HDC hdcDraw, LPRECT prcBounds, LPRECT prcWBounds, HDC hicTargetDev, BOOL fOptimize);
    void    OnDestroyWindow();
    
    void    OnTimer(UINT id);

    void    OnForegroundCompChange(BOOL fCompActive);
    void    OnAppActivationChange(BOOL fActive, DWORD dwThreadID);
    void    OnComponentActivationChange(BOOL fActivate);
    void    OnCtrlFocusChange(BOOL fCtlGotFocus, HWND hFocusWnd);
    void    OnUIFocusChange(BOOL fUIActive);


    void    RaiseActivationEvent(BOOL fActive);

    STDMETHODIMP           ProvideErrorInfo(HRESULT hres);
    STDMETHODIMP           RaiseAutomationEvent(DISPID did, ULONG cargs, VARIANT *pvtargs);

    STDMETHODIMP           SetTempServerLock(BOOL fLock);
	STDMETHODIMP           ResetFrameHook(HWND hwndFrameWindow);


 // The control window proceedure is handled through static class method.
    static LRESULT ControlWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

 // Force redaw of all child windows...
	BOOL InvalidateAllChildWindows(HWND hwnd);
	static BOOL InvalidateAllChildWindowsCallback(HWND hwnd, LPARAM lParam);

 // The variables for the control are kept private but accessible to the
 // nested classes for each interface.
private:

    ULONG                   m_cRef;				   // Reference count
    IUnknown               *m_pOuterUnknown;       // Outer IUnknown (points to m_xInternalUnknown if not agg)
    ITypeInfo              *m_ptiDispType;         // ITypeInfo Pointer (IDispatch Impl)
    EXCEPINFO              *m_pDispExcep;          // EXCEPINFO Pointer (IDispatch Impl)

    HWND                    m_hwnd;                // our window
    HWND                    m_hwndParent;          // immediate parent window
    SIZEL                   m_Size;                // the size of this control  
    RECT                    m_rcLocation;          // where we at

    IOleClientSite         *m_pClientSite;         // active client site of host containter
    IOleControlSite        *m_pControlSite;        // control site
    IOleInPlaceSite        *m_pInPlaceSite;        // inplace site
    IOleInPlaceFrame       *m_pInPlaceFrame;       // inplace frame
    IOleInPlaceUIWindow    *m_pInPlaceUIWindow;    // inplace ui window

    IAdviseSink            *m_pViewAdviseSink;     // advise sink for view (only 1 allowed)
    IOleAdviseHolder       *m_pOleAdviseHolder;    // OLE advise holder (for oleobject sinks)
    IDataAdviseHolder      *m_pDataAdviseHolder;   // OLE data advise holder (for dataobject sink)
    IDispatch              *m_dispEvents;          // event sink (we only support 1 at a time)
    IStorage               *m_pOleStorage;         // IStorage for OLE hosts.

    CDsoDocObject          *m_pDocObjFrame;        // The Embedding Class
    CDsoDocObject          *m_pServerLock;         // Optional Server Lock for out-of-proc DocObject

    BSTR                    m_bstrCustomCaption;   // A custom caption (if provided)
    HMENU                   m_hmenuFilePopup;      // The File menu popup
    WORD                    m_wFileMenuFlags;      // Bitflags of enabled file menu items.
    WORD                    m_wSelMenuItem;        // Which item (if any) is selected
    WORD                    m_cMenuItems;          // Count of items on menu bar
    RECT                    m_rgrcMenuItems[DSO_MAX_MENUITEMS]; // Menu bar items
    CHAR                    m_rgchMenuAccel[DSO_MAX_MENUITEMS]; // Menu bar accelerators
    LPWSTR                  m_pwszHostName;        // Custom name for SetHostNames

    class CDsoFrameHookManager*  m_pHookManager;   // Frame Window Hook Manager Class
	LONG                    m_lHookPolicy;         // Policy on how to use frame hook for this host.
	LONG                    m_lActivationPolicy;   // Policy on activation behavior for comp focus
	HBITMAP                 m_hbmDeactive;         // Bitmap used for IPDeactiveOnXXX policies
    UINT                    m_uiSyncPaint;         // Sync paint counter for draw issues with UIDeactivateOnXXX

    unsigned int        m_fDirty:1;                // does the control need to be resaved?
    unsigned int        m_fInPlaceActive:1;        // are we in place active or not?
    unsigned int        m_fInPlaceVisible:1;       // we are in place visible or not?
    unsigned int        m_fUIActive:1;             // are we UI active or not.
    unsigned int        m_fHasFocus:1;             // do we have current focus.
    unsigned int        m_fViewAdvisePrimeFirst: 1;// for IViewobject2::setadvise
    unsigned int        m_fViewAdviseOnlyOnce: 1;  // for IViewobject2::setadvise
    unsigned int        m_fUsingWindowRgn:1;       // for SetObjectRects and clipping
    unsigned int        m_fFreezeEvents:1;         // should events be frozen?
    unsigned int        m_fDesignMode:1;           // are we in design mode?
    unsigned int        m_fModeFlagValid:1;        // has mode changed since last check?
    unsigned int        m_fModalState:1;           // are we modal?
    unsigned int        m_fObjectMenu:1;           // are we over obj menu item?
    unsigned int        m_fConCntDone:1;           // for enum connectpts
    unsigned int        m_fAppActive:1;            // is the app active?
    unsigned int        m_fComponentActive:1;      // is the component active?
    unsigned int        m_fInDocumentLoad:1;       // set when loading file
    unsigned int        m_fNoInteractive:1;        // set when we don't allow interaction with docobj
    unsigned int        m_fSyncPaintTimer:1;       // is syncpaint timer running?
    unsigned int        m_fInControlActivate:1;    // is currently in activation call?
    unsigned int        m_fInFocusChange:1;        // are we in a focus change?
    unsigned int        m_fActivateOnStatus:1;     // we need to activate on change of status 
    unsigned int        m_fDisableMenuAccel:1;     // using menu accelerators
    unsigned int        m_fBkgrdPaintTimer:1;      // using menu accelerators

};


////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook -- Frame Window Hook Class
//
//  Used by the control to allow for proper host notification of focus 
//  and activation events occurring at top-level window frame. Because 
//  this DocObject host is an OCX, we don't own these notifications and
//  have to "steal" them from our parent using a subclass.
//
//  IMPORTANT: Since the parent frame may exist on a separate thread, this
//  class does nothing but the hook. The code to notify the active component
//  is in a separate global class that is shared by all threads.
//
class CDsoFrameWindowHook
{
public:
	CDsoFrameWindowHook(){ODS("CDsoFrameWindowHook created\n");m_cHookCount=0;m_hwndTopLevelHost=NULL;m_pfnOrigWndProc=NULL;m_fHostUnicodeWindow=FALSE;}
	~CDsoFrameWindowHook(){ODS("CDsoFrameWindowHook deleted\n");}

	static STDMETHODIMP_(CDsoFrameWindowHook*) AttachToFrameWindow(HWND hwndParent);
	STDMETHODIMP Detach();

	static STDMETHODIMP_(CDsoFrameWindowHook*) GetHookFromWindow(HWND hwnd);
	inline STDMETHODIMP_(void) AddRef(){InterlockedIncrement((LONG*)&m_cHookCount);}

    static STDMETHODIMP_(LRESULT) 
		HostWindowProcHook(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

protected:
	DWORD                   m_cHookCount;
	HWND                    m_hwndTopLevelHost;    // Top-level host window (hooked)
    WNDPROC                 m_pfnOrigWndProc;
	BOOL                    m_fHostUnicodeWindow;
};

// THE MAX NUMBER OF DSOFRAMER CONTROLS PER PROCESS
#define DSOF_MAX_CONTROLS   10

////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager -- Hook Manager Class
//
//  Used to keep track of which control is active and forward notifications
//  to it using window messages (to cross thread boundaries).  
//
class CDsoFrameHookManager
{
public:
	CDsoFrameHookManager(){ODS("CDsoFrameHookManager created\n"); m_fAppActive=TRUE; m_idxActive=DSOF_MAX_CONTROLS; m_cComponents=0;}
	~CDsoFrameHookManager(){ODS("CDsoFrameHookManager deleted\n");}

	static STDMETHODIMP_(CDsoFrameHookManager*)
		RegisterFramerControl(HWND hwndParent, HWND hwndControl);

	STDMETHODIMP AddComponent(HWND hwndParent, HWND hwndControl);
	STDMETHODIMP DetachComponent(HWND hwndControl);
	STDMETHODIMP SetActiveComponent(HWND hwndControl);
	STDMETHODIMP OnComponentNotify(DWORD msg, WPARAM wParam, LPARAM lParam);

	inline STDMETHODIMP_(HWND)
		GetActiveComponentWindow(){return m_pComponents[m_idxActive].hwndControl;}

	inline STDMETHODIMP_(CDsoFrameWindowHook*)
		GetActiveComponentFrame(){return m_pComponents[m_idxActive].phookFrame;}

	STDMETHODIMP_(BOOL) SendNotifyMessage(HWND hwnd, DWORD msg, WPARAM wParam, LPARAM lParam);

protected:
	BOOL                    m_fAppActive;
	DWORD                   m_idxActive;
	DWORD                   m_cComponents;
    struct FHOOK_COMPONENTS
	{
		HWND hwndControl;
		CDsoFrameWindowHook *phookFrame;
	}                       m_pComponents[DSOF_MAX_CONTROLS];
};

#endif //DS_DSOFRAMER_H