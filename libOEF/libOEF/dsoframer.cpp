/***************************************************************************
 * DSOFCONTROL.CPP
 *
 * CDsoFramerControl: The Base Control
 ***************************************************************************/

#define INITGUID // Init the GUIDS for the control...

#include "dsoframer.h"



HINSTANCE        v_hModule = NULL;   // DLL module handle
HANDLE           v_hPrivateHeap = NULL;   // Private Memory Heap
ULONG            v_cLocks = 0;      // Count of server locks
HICON            v_icoOffDocIcon = NULL;   // Small office icon (for caption bar)
BOOL             v_fUnicodeAPI = FALSE;  // Flag to determine if we should us Unicode API
BOOL             v_fWindows2KPlus = FALSE;
CRITICAL_SECTION v_csecThreadSynch;

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

	EnterCriticalSection(&v_csecThreadSynch);
	pManager = vgGlobalManager;
	if (pManager == NULL)
		pManager = (vgGlobalManager = new CDsoFrameHookManager());
	LeaveCriticalSection(&v_csecThreadSynch);

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
	EnterCriticalSection(&v_csecThreadSynch);
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
	LeaveCriticalSection(&v_csecThreadSynch);

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
	EnterCriticalSection(&v_csecThreadSynch);

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

	LeaveCriticalSection(&v_csecThreadSynch);

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
	EnterCriticalSection(&v_csecThreadSynch);
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
	LeaveCriticalSection(&v_csecThreadSynch);

	return S_OK;
}

STDMETHODIMP CDsoFrameHookManager::OnComponentNotify(DWORD msg, WPARAM wParam, LPARAM lParam)
{
	TRACE3("CDsoFrameHookManager::OnComponentNotify(%d, %d, %d)\n", msg, wParam, lParam);
	HWND hwndActiveComp = GetActiveComponentWindow();
	
 // The manager needs to keep track of AppActive state...
	if (msg == DSO_WM_HOOK_NOTIFY_APPACTIVATE)
	{
		EnterCriticalSection(&v_csecThreadSynch);
		m_fAppActive = (BOOL)wParam;
		LeaveCriticalSection(&v_csecThreadSynch);
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

	EnterCriticalSection(&v_csecThreadSynch);
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
			phook->m_fHostUnicodeWindow = (v_fUnicodeAPI && IsWindowUnicode(hwndParent));

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
	LeaveCriticalSection(&v_csecThreadSynch);

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
	EnterCriticalSection(&v_csecThreadSynch);
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
	LeaveCriticalSection(&v_csecThreadSynch);
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