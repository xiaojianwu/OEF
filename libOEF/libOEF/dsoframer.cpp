/***************************************************************************
 * DSOFCONTROL.CPP
 *
 * CDsoFramerControl: The Base Control
 ***************************************************************************/

#include "dsoframer.h"



HINSTANCE        v_hModule = NULL;   // DLL module handle
HANDLE           v_hPrivateHeap = NULL;   // Private Memory Heap
ULONG            v_cLocks = 0;      // Count of server locks
HICON            v_icoOffDocIcon = NULL;   // Small office icon (for caption bar)
//BOOL             v_fUnicodeAPI = FALSE;  // Flag to determine if we should us Unicode API
BOOL             v_fWindows2KPlus = FALSE;
//CRITICAL_SECTION v_csecThreadSynch;

CDsoFramerControl::CDsoFramerControl()
{
	m_pDocObjFrame = nullptr;
	m_hwnd = NULL;
	m_pwszHostName = nullptr;

	m_pHookManager = nullptr;
}

CDsoFramerControl::~CDsoFramerControl(void)
{
	ODS("CDsoFramerControl::~CDsoFramerControl\n");

	SAFE_RELEASE_INTERFACE(m_ptiDispType);
	SAFE_RELEASE_INTERFACE(m_pClientSite);
	SAFE_RELEASE_INTERFACE(m_pControlSite);
	SAFE_RELEASE_INTERFACE(m_pInPlaceSite);
	SAFE_RELEASE_INTERFACE(m_pInPlaceFrame);
	SAFE_RELEASE_INTERFACE(m_pInPlaceUIWindow);
	SAFE_RELEASE_INTERFACE(m_pViewAdviseSink);
	SAFE_RELEASE_INTERFACE(m_pOleAdviseHolder);
	SAFE_RELEASE_INTERFACE(m_pDataAdviseHolder);
	SAFE_RELEASE_INTERFACE(m_dispEvents);
	SAFE_RELEASE_INTERFACE(m_pOleStorage);
	SAFE_FREESTRING(m_pwszHostName);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Open
//
//  Creates a document object based on a file or URL. This simulates an
//  "open", but preserves the correct OLE embedding activation required
//  by ActiveX Documents. Opening directly from a file is not recommended.
//  We keep a lock on the original file (unless opened read-only) so the
//  user cannot tell we don't have the file "open".
//
//  The alternate ProgID allows us to open a file that is not associated 
//  with an DocObject server (like *.asp) with the server specified. Also
//  the username/password are for web access (if Document is a URL).
//
HRESULT CDsoFramerControl::Open(LPWSTR pwszDocument, BOOL fOpenReadOnly, LPWSTR pwszAltProgId, HWND hwndParent, RECT dstRect)
{
	m_hwnd = hwndParent;
	m_rcLocation = dstRect;
	HRESULT   hr;
	//LPWSTR    pwszDocument = LPWSTR_FROM_VARIANT(Document);
	//LPWSTR    pwszAltProgId = LPWSTR_FROM_VARIANT(ProgId);
	//BOOL      fOpenReadOnly = BOOL_FROM_VARIANT(ReadOnly, FALSE);
	CLSID     clsidAlt = GUID_NULL;
	HCURSOR	  hCur;
	IUnknown* punk = NULL;

	BIND_OPTS bopts = { sizeof(BIND_OPTS), BIND_MAYBOTHERUSER, 0, 10000 };

	TRACE1("CDsoFramerControl::Open(%S)\n", pwszDocument);


	// If the user passed the ProgId, find the alternative CLSID for server...
	if ((pwszAltProgId) && FAILED(CLSIDFromProgID(pwszAltProgId, &clsidAlt)))
	{
		return E_INVALIDARG;
	}

	// OK. If here, all the parameters look good and it is time to try and open
	// the document object. Start by closing any existing document first...
	if (m_pDocObjFrame && FAILED(hr = Close()))
	{
		return hr;
	}


	// Make sure we are the active component for this process...
	if (FAILED(hr = Activate()))
	{
		return hr;
	}

	// Let's make a doc frame for ourselves...
	if (!(m_pDocObjFrame = CDsoDocObject::CreateInstance((IDsoDocObjectSite*)&m_xDsoDocObjectSite)))
	{
		return E_OUTOFMEMORY;
	}


	// If we had delayed the frame hook, we should set it up now...
	if (!(m_pHookManager))
	{
		m_pHookManager = CDsoFrameHookManager::RegisterFramerControl(m_hwndParent, m_hwnd);
		if (!m_pHookManager)
		{
			return DSO_E_FRAMEHOOKFAILED;
		}
	}

	// Start a wait operation to notify user...
	hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));
	m_fInDocumentLoad = TRUE;

	// Setup the bind options based on read-only flag....
	bopts.grfMode = (STGM_TRANSACTED | STGM_SHARE_DENY_WRITE | (fOpenReadOnly ? STGM_READ : STGM_READWRITE));

	SEH_TRY

		// Normally user gives a string that is path to file...
		if (pwszDocument)
		{
			hr = m_pDocObjFrame->CreateFromFile(pwszDocument, clsidAlt, &bopts);
		}
		else if (punk)
		{
			// If we have an object to load from, try loading it direct...
			hr = m_pDocObjFrame->CreateFromRunningObject(punk, NULL, &bopts);
		}
		else
		{
			hr = E_UNEXPECTED; // Unhandled load type??
		}

	// If successful, we can activate the object...
	if (SUCCEEDED(hr))
	{
		hr = m_pDocObjFrame->IPActivateView();
	}

	SEH_EXCEPT(hr)

		// Force a close if an error occurred...
		if (FAILED(hr))
		{
			m_fFreezeEvents = TRUE;
			Close();
			m_fFreezeEvents = FALSE;
		}
		else
		{
			// Ensure we are active control...
			Activate();
		}

	m_fInDocumentLoad = FALSE;
	SetCursor(hCur);
	return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Close
//
//  Closes the current object.
//
HRESULT CDsoFramerControl::Close()
{
	ODS("CDsoFramerControl::Close\n");

	CDsoDocObject* pdframe = m_pDocObjFrame;
	if (pdframe)
	{
		// If not canceled, clear the member variable then call close on doc frame...
		m_pDocObjFrame = NULL;

		SEH_TRY
			pdframe->Close();
		SEH_EXCEPT_NULL
			delete pdframe;
	}

	return S_OK;
}



////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::Activate
//
//  Activate the current embedded document (i.e, forward focus).
//
STDMETHODIMP CDsoFramerControl::Activate()
{
	HRESULT hr = S_OK;
	ODS("CDsoFramerControl::Activate\n");
	
	return hr;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoFramerControl::XDsoDocObjectSite -- IDsoDocObjectSite Implementation
//
//  STDMETHODIMP GetWindow(HWND* phWnd);
//  STDMETHODIMP GetBorder(LPRECT prcBorder);
//  STDMETHODIMP SetStatusText(LPCOLESTR pszText);
//  STDMETHODIMP GetHostName(LPWSTR *ppwszHostName);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoFramerControl, DsoDocObjectSite)

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	HRESULT hr = E_NOINTERFACE;
	ODS("CDsoFramerControl::XDsoDocObjectSite::QueryService\n");

	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
	CHECK_NULL_RETURN(ppv, E_POINTER);

	// We will return control IDispatch for cross container automation...
	if (guidService == SID_SContainerDispatch)
	{
		ODS(" -- Container Dispatch Given --\n");
		hr = pThis->QueryInterface(riid, ppv);
	}
	// TODO: Support IHlinkFrame services for cross hyperlinking inside control??

	// If we don't handle the call, and our host has a provider, pass the call along to see
	// if the top-level host offers the service requested (like hyperlinking to new web page)...
	if (FAILED(hr) && (pThis->m_pClientSite))
	{
		IServiceProvider *pprov = NULL;
		if (SUCCEEDED(pThis->m_pClientSite->QueryInterface(IID_IServiceProvider, (void**)&pprov)) && (pprov))
		{
			if (SUCCEEDED(hr = pprov->QueryService(guidService, riid, ppv)))
			{
				ODS(" -- Service provided by top-level host --\n");
			}
			pprov->Release();
		}
	}

	return hr;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::GetWindow(HWND* phWnd)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
	ODS("CDsoFramerControl::XDsoDocObjectSite::GetWindow\n");
	if (phWnd) *phWnd = pThis->m_hwnd;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::GetBorder(LPRECT prcBorder)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
	CopyRect(prcBorder, &(pThis->m_rcLocation));

	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::GetHostName(LPWSTR *ppwszHostName)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);
	ODS("CDsoFramerControl::XDsoDocObjectSite::GetHostName\n");
	if (ppwszHostName) *ppwszHostName = DsoCopyString(pThis->m_pwszHostName);
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::SysMenuCommand(UINT uiCharCode)
{
	return S_FALSE;
}

STDMETHODIMP CDsoFramerControl::XDsoDocObjectSite::SetStatusText(LPCOLESTR pszText)
{
	METHOD_PROLOGUE(CDsoFramerControl, DsoDocObjectSite);

	// If m_fActivateOnStatus flag is set, see if we are not UI active. If not, try to activate...
	if ((pThis->m_fActivateOnStatus) && (pThis->m_fComponentActive) && (pThis->m_fAppActive) &&
		!(pThis->m_fUIActive) && !(pThis->m_fModalState))
	{
		ODS(" -- ForceUIActiveFromSetStatusText --\n");
		pThis->m_fActivateOnStatus = FALSE;
		pThis->Activate();
	}

	return S_OK;
}



////////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager
//
CDsoFrameHookManager* vgGlobalManager = NULL;

////////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager::RegisterFramerControl
//
//  Creates a frame hook and registers it with the active manager. If 
//  this is the first time called, the global manager will be created
//  now.  The manager can handle up to DSOF_MAX_CONTROLS controls (10).
//
STDMETHODIMP_(CDsoFrameHookManager*) CDsoFrameHookManager::RegisterFramerControl(HWND hwndParent, HWND hwndControl)
{
	CDsoFrameHookManager* pManager = NULL;
	ODS("CDsoFrameHookManager::RegisterFramerControl\n");
    CHECK_NULL_RETURN(hwndControl, NULL);

	//EnterCriticalSection(&v_csecThreadSynch);
	pManager = vgGlobalManager;
	if (pManager == NULL)
		pManager = (vgGlobalManager = new CDsoFrameHookManager());
	//LeaveCriticalSection(&v_csecThreadSynch);

	if (pManager)
	{
        if (hwndParent == NULL)
            hwndParent = hwndControl;

		if (FAILED(pManager->AddComponent(hwndParent, hwndControl)))
			return NULL;
	}

	return pManager;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameHookManager::AddComponent
//
//  Register the component and create the actual frame hook (if needed).
//
STDMETHODIMP CDsoFrameHookManager::AddComponent(HWND hwndParent, HWND hwndControl)
{
	DWORD dwHostProcessID;
	CDsoFrameWindowHook* phook;
	HRESULT hr = E_OUTOFMEMORY;
	HWND hwndNext, hwndHost = hwndParent;
	ODS("CDsoFrameHookManager::AddComponent\n");

 // Find the top-level window this control is sited on...
	while (hwndNext = GetParent(hwndHost))
		hwndHost = hwndNext;

 // We have to get valid window that is in-thread for a subclass to work...
	if (!IsWindow(hwndHost) || 
		(GetWindowThreadProcessId(hwndHost, &dwHostProcessID), 
			dwHostProcessID != GetCurrentProcessId()))
		return E_ACCESSDENIED;

 // Next if we have room in the array, add the component to the list and
 // hook its parent window so we get messages for it.
	//EnterCriticalSection(&v_csecThreadSynch);
	if ((vgGlobalManager) && (m_cComponents < DSOF_MAX_CONTROLS))
	{
	 // Hook the host window for this control...
		phook = CDsoFrameWindowHook::AttachToFrameWindow(hwndHost);
		if (phook) 
		{// Add the component to the list if hooked...
			m_pComponents[m_cComponents].hwndControl = hwndControl;
			m_pComponents[m_cComponents].phookFrame  = phook;
			m_cComponents++; hr = S_OK; 
		}
	}
	//LeaveCriticalSection(&v_csecThreadSynch);

	if (SUCCEEDED(hr)) // Make sure item is the new active component
		hr = SetActiveComponent(hwndControl);

	return hr;
}

STDMETHODIMP CDsoFrameHookManager::DetachComponent(HWND hwndControl)
{
	DWORD dwIndex;
	CHECK_NULL_RETURN(m_cComponents, E_FAIL);

 // Find this control in the list...
	for (dwIndex = 0; dwIndex < m_cComponents; dwIndex++)
		if (m_pComponents[dwIndex].hwndControl == hwndControl) break;

	if (dwIndex > m_cComponents)
		return E_INVALIDARG;

 // If we found the index, remove the item and shift remaining
 // items down the list. Change active index as needed...
	//EnterCriticalSection(&v_csecThreadSynch);

	m_pComponents[dwIndex].hwndControl = NULL;
	m_pComponents[dwIndex].phookFrame->Detach();

 // Compact the list...
	while (++dwIndex < m_cComponents)
	{
		m_pComponents[dwIndex-1].hwndControl = m_pComponents[dwIndex].hwndControl;
		m_pComponents[dwIndex-1].phookFrame = m_pComponents[dwIndex].phookFrame;
		m_pComponents[dwIndex].hwndControl = NULL;
	}

 // Decrement the count. If count still exists, forward activation
 // to next component in the list...
	if (--m_cComponents)
	{
		while (m_idxActive >= m_cComponents)
			--m_idxActive;

		HWND hwndActive = GetActiveComponentWindow();
		if (hwndActive) 
			SendNotifyMessage(hwndActive, DSO_WM_HOOK_NOTIFY_COMPACTIVE, (WPARAM)TRUE, (LPARAM)(m_fAppActive));
	}
	else
	{
		vgGlobalManager = NULL;
	 // If this is the last control, we can remove the manger!
		delete this;
	}

	//LeaveCriticalSection(&v_csecThreadSynch);

	return S_OK;
}

STDMETHODIMP CDsoFrameHookManager::SetActiveComponent(HWND hwndControl)
{
	ODS(" -- CDsoFrameHookManager::SetActiveComponent -- \n");
	CHECK_NULL_RETURN(m_cComponents, E_FAIL);
	DWORD dwIndex;

 // Find the index of the control...
	for (dwIndex = 0; dwIndex < m_cComponents; dwIndex++)
		if (m_pComponents[dwIndex].hwndControl == hwndControl) break;

	if (dwIndex > m_cComponents)
		return E_INVALIDARG;

 // If it is not already the active item, notify old component it is
 // losing activation, and notify new component that it is gaining...
	//EnterCriticalSection(&v_csecThreadSynch);
	if (dwIndex != m_idxActive)
	{
		HWND hwndActive = GetActiveComponentWindow();
		if (hwndActive) 
			SendNotifyMessage(hwndActive, DSO_WM_HOOK_NOTIFY_COMPACTIVE, (WPARAM)FALSE, (LPARAM)0);

		m_idxActive = dwIndex;

		hwndActive = GetActiveComponentWindow();
		if (hwndActive)
			SendNotifyMessage(hwndActive, DSO_WM_HOOK_NOTIFY_COMPACTIVE, (WPARAM)TRUE, (LPARAM)(m_fAppActive));
	}
	//LeaveCriticalSection(&v_csecThreadSynch);

	return S_OK;
}

STDMETHODIMP CDsoFrameHookManager::OnComponentNotify(DWORD msg, WPARAM wParam, LPARAM lParam)
{
	TRACE3("CDsoFrameHookManager::OnComponentNotify(%d, %d, %d)\n", msg, wParam, lParam);
	HWND hwndActiveComp = GetActiveComponentWindow();
	
 // The manager needs to keep track of AppActive state...
	if (msg == DSO_WM_HOOK_NOTIFY_APPACTIVATE)
	{
		//EnterCriticalSection(&v_csecThreadSynch);
		m_fAppActive = (BOOL)wParam;
		//LeaveCriticalSection(&v_csecThreadSynch);
	}

 // Send the message to the active component
    SendNotifyMessage(hwndActiveComp, msg, wParam, lParam);
	return S_OK;
}

STDMETHODIMP_(BOOL) CDsoFrameHookManager::SendNotifyMessage(HWND hwnd, DWORD msg, WPARAM wParam, LPARAM lParam)
{
    BOOL fResult;

 // Check if the caller is on the same thread as the active component. If so we can
 // forward the message directly for processing. If not, we have to post and wait for
 // that thread to process message at a later time...
	if (GetWindowThreadProcessId(hwnd, NULL) == GetCurrentThreadId())
	{ // Notify now...
		fResult = (BOOL)SendMessage(hwnd, msg, wParam, lParam);
	}
	else
	{ // Notify later...
		fResult = PostMessage(hwnd, msg, wParam, lParam);
	}

    return fResult;
}


////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook
//
////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::AttachToFrameWindow
//
//  Creates the frame hook for the host window.
//
STDMETHODIMP_(CDsoFrameWindowHook*) CDsoFrameWindowHook::AttachToFrameWindow(HWND hwndParent)
{
	CDsoFrameWindowHook* phook = NULL;
	BOOL fHookSuccess = FALSE;

 // We have to get valid window ...
	if (!IsWindow(hwndParent))
		return NULL;

	//EnterCriticalSection(&v_csecThreadSynch);
	phook = CDsoFrameWindowHook::GetHookFromWindow(hwndParent);
	if (phook)
	{ // Already have this window hooked, so just addref the count.
		phook->AddRef();
	}
	else
	{ // Need to create a new hook object for this window...
		if ((phook = new CDsoFrameWindowHook()))
		{
			phook->m_hwndTopLevelHost = hwndParent;
			phook->m_fHostUnicodeWindow = true; //  (v_fUnicodeAPI && IsWindowUnicode(hwndParent));

		 // Setup the subclass using Unicode or Ansi methods depending on window 
		 // create state. Note if either property assoc or window subclass fail...
			if (phook->m_fHostUnicodeWindow)
			{
				fHookSuccess = SetPropW(hwndParent, L"DSOFramerWndHook", (HANDLE)phook);
				if (fHookSuccess)
				{
					phook->m_pfnOrigWndProc = (WNDPROC)SetWindowLongW(hwndParent, GWL_WNDPROC, (LONG)(WNDPROC)HostWindowProcHook);
					fHookSuccess = (phook->m_pfnOrigWndProc != NULL);
				}
			}
			else
			{
				fHookSuccess = SetPropA(hwndParent, "DSOFramerWndHook", (HANDLE)phook);
				if (fHookSuccess)
				{
					phook->m_pfnOrigWndProc = (WNDPROC)SetWindowLongA(hwndParent, GWL_WNDPROC, (LONG)(WNDPROC)HostWindowProcHook);
					fHookSuccess = (phook->m_pfnOrigWndProc != NULL);
				}
			}

		 // If we failed the subclass, kill this object and return error...
			if (!fHookSuccess)
			{
				delete phook;
				phook = NULL;
			}
			else phook->m_cHookCount++;
		}
	}
	//LeaveCriticalSection(&v_csecThreadSynch);

	return phook;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::Detach
//
//  Removes the hook from the frame window if this is the last control
//  to reference the hook.
//
STDMETHODIMP CDsoFrameWindowHook::Detach()
{
 // Only need to call this if item is attached.
	//EnterCriticalSection(&v_csecThreadSynch);
	if (m_cHookCount)
	{
		if (--m_cHookCount == 0)
		{ // We can remove the hook!
			if (m_fHostUnicodeWindow)
			{
				SetWindowLongW(m_hwndTopLevelHost, GWL_WNDPROC, (LONG)(m_pfnOrigWndProc));
				RemovePropW(m_hwndTopLevelHost, L"DSOFramerWndHook");
			}
			else
			{
				SetWindowLongA(m_hwndTopLevelHost, GWL_WNDPROC, (LONG)(m_pfnOrigWndProc));
				RemovePropA(m_hwndTopLevelHost, "DSOFramerWndHook");
			}
			m_hwndTopLevelHost = NULL;

		// we can kill this object now...
			delete this;
		}
	}
	//LeaveCriticalSection(&v_csecThreadSynch);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::GetHookFromWindow
//
//  Returns the hook control associated with this object.
//
STDMETHODIMP_(CDsoFrameWindowHook*) CDsoFrameWindowHook::GetHookFromWindow(HWND hwnd)
{
	CDsoFrameWindowHook* phook = (CDsoFrameWindowHook*)GetPropW(hwnd, L"DSOFramerWndHook");
	if (phook == NULL) phook = (CDsoFrameWindowHook*)GetPropA(hwnd, "DSOFramerWndHook");
	return phook;
}

////////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook::HostWindowProcHook
//
//  The window proc for the subclassed host window.
//
STDMETHODIMP_(LRESULT) CDsoFrameWindowHook::HostWindowProcHook(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	CDsoFrameWindowHook* phook = CDsoFrameWindowHook::GetHookFromWindow(hwnd);
	if (phook) // We better get our prop for this window!
	{
	 // We handle the important ActiveDocument frame-level events like task 
	 // window activation/deactivation, palette change notfication, and key/mouse
	 // focus notification on host app focus...
		switch (msg)
		{
		case WM_ACTIVATEAPP:
			if (vgGlobalManager)
				vgGlobalManager->OnComponentNotify(DSO_WM_HOOK_NOTIFY_APPACTIVATE, wParam, lParam);
			break;

		case WM_PALETTECHANGED:
			if (vgGlobalManager)
				vgGlobalManager->OnComponentNotify(DSO_WM_HOOK_NOTIFY_PALETTECHANGE, wParam, lParam);
			break;

        case WM_SYNCPAINT:
			if (vgGlobalManager)
				vgGlobalManager->OnComponentNotify(DSO_WM_HOOK_NOTIFY_SYNCPAINT, wParam, lParam);
			break;

		}

	 // After processing, allow calls to fall through to host app. We need to call
	 // the appropriate handler for Unicode or non-Unicode windows...
		if (phook->m_fHostUnicodeWindow)
			return CallWindowProcW(phook->m_pfnOrigWndProc, hwnd, msg, wParam, lParam);
		else
			return CallWindowProcA(phook->m_pfnOrigWndProc, hwnd, msg, wParam, lParam);
	}

 // Should not be reached, but just in case call default proc...
	return (IsWindowUnicode(hwnd) ? 
			DefWindowProcW(hwnd, msg, wParam, lParam) : 
			DefWindowProcA(hwnd, msg, wParam, lParam));
}