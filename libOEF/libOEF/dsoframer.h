/***************************************************************************
 * DSOFRAMER.H
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
extern ULONG            v_cLocks;

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
class CDsoFramerControl
{
public:
	CDsoFramerControl();
    ~CDsoFramerControl(void);

	HRESULT Open(LPWSTR pwszDocument, BOOL fOpenReadOnly, LPWSTR pwszAltProgId, HWND hwndParent, RECT dstRect);

	void    OnResize(RECT dstRect);


	//HRESULT Save(VARIANT SaveAsDocument, VARIANT OverwriteExisting);
	HRESULT Close();

	HWND GetActiveWindow();

	void ReobtainActiveFrame();

	HRESULT GetActiveDocument(IDispatch** ppdisp);

 // The variables for the control are kept private but accessible to the
 // nested classes for each interface.
private:
    CDsoDocObject          *m_pDocObjFrame;        // The Embedding Class
};

#endif //DS_DSOFRAMER_H