/***************************************************************************
 * DSOFDOCOBJ.CPP
 *
 * CDsoDocObject: ActiveX Document Single Instance Frame/Site Object
 *
 ***************************************************************************/
#include "dsoframer.h"

 ////////////////////////////////////////////////////////////////////////
 // CDsoDocObject - The DocObject Site Class
 //
 //  This class wraps the functionality for DocObject hosting. Right now
 //  we are setup for one active site at a time, but this could be changed
 //  to allow multiple sites (although only one could be UI active at any
 //  given time).
 //
CDsoDocObject::CDsoDocObject()
{
	ODS("CDsoDocObject::CDsoDocObject\n");
	m_cRef = 1;
	m_fDisplayTools = TRUE;

	m_pstgroot = nullptr;
	m_pwszHostName = nullptr;
	m_pdocv = nullptr;
	m_pstmview = nullptr;
	m_pipactive = nullptr;
	m_pipobj = nullptr;
	m_pole = nullptr;

	m_pstgSourceFile = nullptr;
	m_pmkSourceFile = nullptr;
	m_pbctxSourceFile = nullptr;

	m_pstgfile = nullptr;
	m_pcmdCtl = nullptr;

	m_pcmdt = nullptr;

	m_hwndIPObject = NULL;
}

CDsoDocObject::~CDsoDocObject(void)
{
	ODS("CDsoDocObject::~CDsoDocObject\n");
	if (m_pole)	Close();
	if (m_hwnd) DestroyWindow(m_hwnd);

	SAFE_FREESTRING(m_pwszHostName);

	SAFE_RELEASE_INTERFACE(m_pstgroot);
	SAFE_RELEASE_INTERFACE(m_pcmdCtl);
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateNewDocObject
//
//  Static Creation Function.
//
CDsoDocObject* CDsoDocObject::CreateInstance(HWND hwndCtl, RECT rect)
{
	ODS("CDsoDocObject::CreateInstance()\n");
	CDsoDocObject* pnew = new CDsoDocObject();
	if ((pnew) && FAILED(pnew->InitializeNewInstance(hwndCtl, rect)))
	{
		pnew->Release();
		pnew = NULL;
	}
	return pnew;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::InitializeNewInstance
//
//  Sets up new docobject class. We must a control site to attach this
//  window to. It will call back to host for menu and IOleCommandTarget.
//
STDMETHODIMP CDsoDocObject::InitializeNewInstance(HWND hwndCtl, RECT rect)
{
	HRESULT hr = E_UNEXPECTED;
	WNDCLASS wndclass;

	ODS("CDsoDocObject::InitializeNewInstance()\n");

	// As an AxDoc site, we need a valid parent window...
	if ((!hwndCtl) || (!IsWindow(hwndCtl)))
		return hr;

	// Create a temp storage for this docobj site (if one already exists, bomb out)...
	if ((m_pstgroot) || FAILED(hr = StgCreateDocfile(NULL, STGM_TRANSACTED | STGM_READWRITE |
		STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_DELETEONRELEASE, 0, &m_pstgroot)))
		return hr;

	// If our site window class has not been registered before, we should register it...

	// This is protected by a critical section just for fun. The fact we had to single
	// instance the OCX because of the host hook makes having multiple instances conflict here
	// very unlikely. However, that could change sometime, so better to be safe than sorry.
	   //EnterCriticalSection(&v_csecThreadSynch);

	if (GetClassInfo(v_hModule, L"DSOFramerDocWnd", &wndclass) == 0)
	{
		memset(&wndclass, 0, sizeof(WNDCLASS));
		wndclass.style = CS_VREDRAW | CS_HREDRAW | CS_DBLCLKS;
		wndclass.lpfnWndProc = (WNDPROC)CDsoDocObject::FrameWindowProc;
		wndclass.hInstance = v_hModule;
		wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
		wndclass.lpszClassName = L"DSOFramerDocWnd";
		if (RegisterClass(&wndclass) == 0)
			hr = E_WIN32_LASTERROR;
	}

	if (FAILED(hr)) return hr;

	CopyRect(&m_rcViewRect, &(rect));


	if (FAILED(hr)) return hr;

	if (m_rcViewRect.top > m_rcViewRect.bottom)
	{
		m_rcViewRect.top = 0; m_rcViewRect.bottom = 0;
	}
	if (m_rcViewRect.left > m_rcViewRect.right)
	{
		m_rcViewRect.left = 0; m_rcViewRect.right = 0;
	}

	// Create our site window at the give location (we are child of the control window)...
	m_hwnd = CreateWindowEx(0, L"DSOFramerDocWnd", NULL, WS_CHILD,
		m_rcViewRect.left, m_rcViewRect.top,
		(m_rcViewRect.right - m_rcViewRect.left),
		(m_rcViewRect.bottom - m_rcViewRect.top),
		hwndCtl, NULL, v_hModule, NULL);

	if (!m_hwnd) return E_OUTOFMEMORY;
	SetWindowLong(m_hwnd, GWL_USERDATA, (LONG)this);

	// Save the control info for this object...
	m_hwndCtl = hwndCtl;

	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateDocObject (From CLSID)
//
//  This does embedding of new object, or loads a copy from storage.
//  To activate and show the object, you must call IPActivateView().
//
STDMETHODIMP CDsoDocObject::CreateDocObject(REFCLSID rclsid)
{
	HRESULT             hr;
	CLSID               clsid;
	DWORD               dwMiscStatus = 0;
	IOleObject*         pole = NULL;
	IPersistStorage*    pipstg = NULL;
	BOOL                fInitNew = (m_pstgfile == NULL);

	ODS("CDsoDocObject::CreateDocObject(CLSID)\n");
	ASSERT(!(m_pole));

	// Don't load if an object has already been loaded...
	if (m_pole) return E_UNEXPECTED;

	// It is possible that someone picked an older ProgId/CLSID that
	// will AutoConvert on CoCreate, so fix up the storage with the
	// new CLSID info. We we actually call CoCreate on the new CLSID...
	if (fInitNew && SUCCEEDED(OleGetAutoConvert(rclsid, &clsid)))
	{
		OleDoAutoConvert(m_pstgfile, &clsid);
	}
	else clsid = rclsid;

	// First, check the server to make sure it is AxDoc server...
	if (FAILED(hr = ValidateDocObjectServer(rclsid)))
		return hr;

	// If we haven't loaded a storage, create a new one and remember to
	// call InitNew (instead of Load) later on...
	if (fInitNew && FAILED(hr = CreateObjectStorage(rclsid)))
		return hr;

	// We are ready to create an instance. Call CoCreate to make an
	// inproc handler and ask for IOleObject (all docobjs must support this)...
	if (FAILED(hr = InstantiateDocObjectServer(clsid, &pole)))
		return hr;

	// Do a quick check to see if server wants us to set client site before the load..
	hr = pole->GetMiscStatus(DVASPECT_CONTENT, &dwMiscStatus);
	if (dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST)
		hr = pole->SetClientSite((IOleClientSite*)&m_xOleClientSite);

	// Load up the bloody thing...
	if (SUCCEEDED(hr = pole->QueryInterface(IID_IPersistStorage, (void**)&pipstg)))
	{
		// Remember to InitNew if this is a new storage...			
		hr = ((fInitNew) ? pipstg->InitNew(m_pstgfile) : pipstg->Load(m_pstgfile));
		pipstg->Release();
	}

	// Assuming all the above worked we should have an OLE Embeddable
	// object and should finish the initialization (set object running)...
	if (SUCCEEDED(hr))
	{
		// Save the IOleObject* and do a disconnect on quit...
		SAFE_SET_INTERFACE(m_pole, pole);
		m_fDisconnectOnQuit = TRUE;

		// Keep server CLSID for this object
		m_clsidObject = clsid;

		// Ensure server is running and locked...
		EnsureOleServerRunning(TRUE);

		// If we didn't do so already, set our client site...
		if (!(dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST))
			hr = m_pole->SetClientSite((IOleClientSite*)&m_xOleClientSite);

		// Set the host names for OLE embedding...
		m_pole->SetHostNames(m_pwszHostName, m_pwszHostName);

		// Ask object to save (if dirty)...
		if (IsDirty()) SaveObjectStorage();
	}
	else
	{// Be sure they disconnect from our site if we failed load...
		pole->SetClientSite(NULL);
	}

	// This will free the OLE server if anything above failed...
	SAFE_RELEASE_INTERFACE(pole);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateDocObject (From IStorage*)
//
//  This does embedding of new object, or loads a copy from storage.
//  To activate and show the object, you must call IPActivateView().
//
STDMETHODIMP CDsoDocObject::CreateDocObject(IStorage *pstg)
{
	HRESULT hr;
	CLSID clsid;

	ODS("CDsoDocObject::CreateDocObject(IStorage*)\n");
	CHECK_NULL_RETURN(pstg, E_POINTER);

	// Read the clsid from the storage..
	if (FAILED(hr = ReadClassStg(pstg, &clsid)))
		return hr;

	// Validate the server is AxDoc server...
	if (FAILED(hr = ValidateDocObjectServer(clsid)))
		return hr;

	// Create a new storage for this CLSID...
	if (FAILED(hr = CreateObjectStorage(clsid)))
		return hr;

	// Copy data into the new storage and commit the change...
	hr = pstg->CopyTo(0, NULL, NULL, m_pstgfile);
	if (SUCCEEDED(hr)) hr = m_pstgfile->Commit(STGC_OVERWRITE);

	// Then do normal create on existing storage made from copy...
	return CreateDocObject(clsid);
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateFromFile
//
//  Loads document into the framer control. The bind options determine if
//  file should be loaded read-only or read-write. An alternate CLSID can 
//  be provided if the document type is not normally associated with a 
//  a DocObject server in OLE.
//
//  To activate and show the object, you must call IPActivateView().
//
//  The code will attempt to open the file in one of three ways:
//
//   1.) Using IMoniker, we will attempt to open and load existing file
//       via a passed moniker, similar to how IE loads files.
//
//   2.) Using IPersistFile, we will attempt to open via direct file path.
//
//   3.) Finally, if above two options don't work, we will try to open the 
//       storage, make a copy and load using IPersistStorage. This last way follows
//       the (old) Binder behavior, and should *by spec* always work, but is a bit 
//       clunky since we open a copy, not the real file.
//
STDMETHODIMP CDsoDocObject::CreateFromFile(LPWSTR pwszFile, REFCLSID rclsid, LPBIND_OPTS pbndopts)
{
	HRESULT			hr;
	CLSID           clsid;
	CLSID           clsidConv;
	IOleObject      *pole = NULL;
	IBindCtx		*pbctx = NULL;
	IMoniker		*pmkfile = NULL;
	IStorage        *pstg = NULL;
	BOOL fLoadFromAltCLSID = (rclsid != GUID_NULL);

	// Sanity check of parameters...
	if (!(pwszFile) || ((*pwszFile) == L'\0') || (pbndopts == NULL))
		return E_INVALIDARG;

	TRACE2("CDsoDocObject::CreateFromFile(%S, %x)\n", pwszFile, pbndopts->grfMode);

	// First. we'll try to find the associated CLSID for the given file,
	// and then set it to the alternate if not found. If we don't have a
	// CLSID by the end of this, because user didn't specify alternate
	// and GetClassFile failed, then we error out...
	if (FAILED(GetClassFile(pwszFile, &clsid)) && !(fLoadFromAltCLSID))
	{
		return DSO_E_INVALIDSERVER;
	}


	// We should try to load from alternate CLSID if provided one...
	if (fLoadFromAltCLSID) clsid = rclsid;

	// We should also handle auto-convert to start "newest" server...
	if (SUCCEEDED(OleGetAutoConvert(clsid, &clsidConv)))
		clsid = clsidConv;

	// Validate that we have a DocObject server...
	if ((clsid == GUID_NULL) || FAILED(ValidateDocObjectServer(clsid)))
	{
		return DSO_E_INVALIDSERVER;
	}

	// First, we try to bind by moniker (same as IE). We'll need a bind context 
	// and a file moniker for the orginal source...
	if (SUCCEEDED(hr = CreateBindCtx(0, &pbctx)))
	{
		if (SUCCEEDED(hr = pbctx->SetBindOptions(pbndopts)) &&
			SUCCEEDED(hr = CreateFileMoniker(pwszFile, &pmkfile)))
		{

			// Bind to the object moniker refers to...
			hr = pmkfile->BindToObject(pbctx, NULL, IID_IOleObject, (void**)&pole);

			// If that failed, try to bind direct to file in new server...
			if (FAILED(hr))
			{
				IPersistFile    *pipfile = NULL;

				if (SUCCEEDED(hr = InstantiateDocObjectServer(clsid, &pole)) &&
					SUCCEEDED(hr = pole->QueryInterface(IID_IPersistFile, (void**)&pipfile)))
				{
					hr = pipfile->Load(pwszFile, pbndopts->grfMode);
					pipfile->Release();
				}
			}


			// If either solution worked, setup the rest of the bind info...
			if (SUCCEEDED(hr))
			{
				// Save the IOleObject* and do a disconnect on quit...
				SAFE_SET_INTERFACE(m_pole, pole);
				m_fDisconnectOnQuit = TRUE;

				// Keep server CLSID for this object
				m_clsidObject = clsid;

				// Keep the moniker and bind ctx...
				SAFE_SET_INTERFACE(m_pmkSourceFile, pmkfile);
				SAFE_SET_INTERFACE(m_pbctxSourceFile, pbctx);

				// Set out client site...
				m_pole->SetClientSite((IOleClientSite*)&m_xOleClientSite);

				// We don't normally set host name for moniker bind, but MSWORD object
				// requires it to IP activate it instead of link & show external...
				if (IsWordObject())
					m_pole->SetHostNames(m_pwszHostName, m_pwszHostName);

			}
			else pole->SetClientSite(NULL);

			// This will release the moniker if above failed...
			pmkfile->Release();
		}
		// This will release the bind ctx if above failed...
		pbctx->Release();
	}

	// If binding by moniker failed, try the old fashion way of bind to OLE storage...
	if (FAILED(hr))
	{
		// Try to open file as OLE storage (native OLE DocFile)...
		if (SUCCEEDED(hr = StgOpenStorage(pwszFile, NULL, pbndopts->grfMode, NULL, 0, &pstg)))
		{
			// Create our substorage, and copy the data over to it...
			if (SUCCEEDED(hr = CreateObjectStorage(clsid)) &&
				SUCCEEDED(hr = pstg->CopyTo(0, NULL, NULL, m_pstgfile)))
			{
				m_pstgfile->Commit(STGC_OVERWRITE);

				// Then create the object from the storage copy...
				if (SUCCEEDED(hr = CreateDocObject(clsid)))
				{
					SAFE_SET_INTERFACE(m_pstgSourceFile, pstg);
				}
			}
			// This will release the storage if above failed...
			pstg->Release();
		}
	}

	// If all went well, we should save file name and whether it opened read-only...
	if (SUCCEEDED(hr))
	{
		m_pwszSourceFile = DsoCopyString(pwszFile);
		m_fOpenReadOnly = ((pbndopts->grfMode & STGM_WRITE) == 0) &&
			((pbndopts->grfMode & STGM_READWRITE) == 0);
	}

	// This will free the OLE server if anything above failed...
	SAFE_RELEASE_INTERFACE(pole);
	return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::IPActivateView
//
//  Activates the object for inplace viewing.
//
STDMETHODIMP CDsoDocObject::IPActivateView()
{
	HRESULT hr = E_UNEXPECTED;
	ODS("CDsoDocObject::IPActivateView()\n");
	ASSERT(m_pole);

	// Make sure the site window is made visible...
	if (!IsWindowVisible(m_hwnd))
		ShowWindow(m_hwnd, SW_SHOW);

	// If we have an IOleDocument pointer, we can use the Show method for
	// inplace activation...
	if (m_pdocv)
	{
		hr = m_pdocv->Show(TRUE);
	}
	else if (m_pole)
	{
		// Try creating a document view and loading it...
		hr = m_xOleDocumentSite.ActivateMe(NULL);
		if (FAILED(hr)) // If that fails, go the old OLE route...
		{
			RECT rcView; GetClientRect(m_hwnd, &rcView);

			// First call IOleObject::DoVerb with OLEIVERB_INPLACEACTIVATE
			// (or OLEIVERB_SHOW), and our view rect and IOleClientSite pointer...
			hr = m_pole->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL,
				(IOleClientSite*)&m_xOleClientSite, (UINT)-1, m_hwnd, &rcView);

			// If the server doesn't recognize IP verb, try OLEIVERB_SHOW instead...
			if (hr == OLEOBJ_E_INVALIDVERB)
				hr = m_pole->DoVerb(OLEIVERB_SHOW, NULL,
				(IOleClientSite*)&m_xOleClientSite, (UINT)-1, m_hwnd, &rcView);

			// There is an issue with Visio 2002 rejecting DoVerb when it is called while
			// Visio is still loading up the main app window. OLE servers normally return RPC 
			// retry later error, which will spin us in the msg-filter, but Visio just rejects
			// the call all together, which pops us out of the filter early. This causes the 
			// embed attempt to fail. So to workaround this, we will check for the condition, 
			// sleep a bit, and then try our own semi-msg-filter loop...
			if ((hr == RPC_E_CALL_REJECTED) && IsVisioObject())
			{
				DWORD dwLoopCnt = 0;
				do
				{
					Sleep((200 * ++dwLoopCnt));
					hr = m_pole->DoVerb(OLEIVERB_SHOW, NULL,
						(IOleClientSite*)&m_xOleClientSite, (UINT)-1, m_hwnd, &rcView);
				} while ((hr == RPC_E_CALL_REJECTED) && (dwLoopCnt < 4));
			}
		}
	}

	// Go ahead and UI activate now...
	if (SUCCEEDED(hr))
		hr = UIActivateView();

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::IPDeactivateView
//
//  Deactivates the object ready for close.
//
STDMETHODIMP CDsoDocObject::IPDeactivateView()
{
	HRESULT hr = S_OK;
	ODS("CDsoDocObject::IPDeactivateView()\n");

	// If we still have a UI active object, tell it to UI deactivate...
	if (m_pipactive)
		UIDeactivateView();

	// Next hide the active object...
	if (m_pdocv)
		m_pdocv->Show(FALSE);

	// Notify object our intention to IP deactivate...
	if (m_pipobj)
		m_pipobj->InPlaceDeactivate();

	// Close the object down and release pointers...
	if (m_pdocv)
	{
		hr = m_pdocv->CloseView(0);
		m_pdocv->SetInPlaceSite(NULL);
	}

	SAFE_RELEASE_INTERFACE(m_pcmdt);
	SAFE_RELEASE_INTERFACE(m_pdocv);
	SAFE_RELEASE_INTERFACE(m_pipobj);

	// Hide the site window...
	ShowWindow(m_hwnd, SW_HIDE);

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::UIActivateView
//
//  UI Activates the object (which should bring up toolbars and caret).
//
STDMETHODIMP CDsoDocObject::UIActivateView()
{
	HRESULT hr = S_FALSE;
	ODS("CDsoDocObject::UIActivateView()\n");

	if (m_pdocv) // Go UI active...
	{
		hr = m_pdocv->UIActivate(TRUE);
	}
	else if (m_pole) // We should never get here, but just in case pdocv is NULL, signal UI active the old way...
	{
		RECT rcView; GetClientRect(m_hwnd, &rcView);
		m_pole->DoVerb(OLEIVERB_UIACTIVATE, NULL, (IOleClientSite*)&m_xOleClientSite, (UINT)-1, m_hwnd, &rcView);
	}

	// Forward focus to the IP object...
	if (SUCCEEDED(hr) && (m_hwndIPObject))
		SetFocus(m_hwndIPObject);

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::UIDeactivateView
//
//  UI Deactivates the object (which should hide toolbars and caret).
//
STDMETHODIMP CDsoDocObject::UIDeactivateView()
{
	HRESULT hr = S_FALSE;
	ODS("CDsoDocObject::UIDeactivateView()\n");
	if (m_pdocv)
		hr = m_pdocv->UIActivate(FALSE);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::IsDirty
//
//  Determines if the file is dirty (by asking IPersistXXX::IsDirty).
//
BOOL CDsoDocObject::IsDirty()
{
	BOOL fDirty = TRUE; // Assume we are dirty unless object says we are not
	IPersistStorage *pprststg;
	IPersistFile *pprst;

	// Can't be dirty without object
	CHECK_NULL_RETURN(m_pole, FALSE);

	// Ask object its dirty state...
	if ((m_pmkSourceFile) &&
		SUCCEEDED(m_pole->QueryInterface(IID_IPersistFile, (void**)&pprst)))
	{
		fDirty = ((pprst->IsDirty() == S_FALSE) ? FALSE : TRUE);
		pprst->Release();
	}
	else if (SUCCEEDED(m_pole->QueryInterface(IID_IPersistStorage, (void**)&pprststg)))
	{
		fDirty = ((pprststg->IsDirty() == S_FALSE) ? FALSE : TRUE);
		pprststg->Release();
	}

	return fDirty;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::DoOleCommand
//
//  Calls IOleCommandTarget::Exec on the active object to do a specific
//  command (like Print, SaveCopy, Zoom, etc.). 
//
STDMETHODIMP CDsoDocObject::DoOleCommand(DWORD dwOleCmdId, DWORD dwOptions, VARIANT* vInParam, VARIANT* vInOutParam)
{
	HRESULT hr;
	OLECMD cmd = { dwOleCmdId, 0 };
	TRACE2("CDsoDocObject::DoOleCommand(cmd=%d, Opts=%d\n", dwOleCmdId, dwOptions);

	// Can't issue OLECOMMANDs when in print preview mode (object calls us)...
	   //if (InPrintPreview()) return E_ACCESSDENIED;

	// The server must support IOleCommandTarget, the CmdID being requested, and
	// the command should be enabled. If this is the case, do the command...
	if ((m_pcmdt) && SUCCEEDED(m_pcmdt->QueryStatus(NULL, 1, &cmd, NULL)) &&
		((cmd.cmdf & OLECMDF_SUPPORTED) && (cmd.cmdf & OLECMDF_ENABLED)))
	{
		TRACE1("QueryStatus say supported = 0x%X\n", cmd.cmdf);

		// Do the command asked by caller on default command group...
		hr = m_pcmdt->Exec(NULL, cmd.cmdID, dwOptions, vInParam, vInOutParam);
		TRACE1("DocObj_IOleCommandTarget::Exec() = 0x%X\n", hr);

		if ((dwOptions == OLECMDEXECOPT_PROMPTUSER))
		{
			// Handle bug issue for PPT when printing using prompt...
			if ((hr == E_INVALIDARG) && (cmd.cmdID == OLECMDID_PRINT) && IsPPTObject())
			{
				ODS("Retry command for PPT\n");
				hr = m_pcmdt->Exec(NULL, cmd.cmdID, OLECMDEXECOPT_DODEFAULT, vInParam, vInOutParam);
				TRACE1("DocObj_IOleCommandTarget::Exec() = 0x%X\n", hr);
			}

			// If user canceled an Office dialog, that's OK...
			if (hr == 0x80040103)
				hr = S_FALSE;
		}
	}
	else
	{
		TRACE1("Command Not supportted (%d)\n", cmd.cmdf);
		hr = DSO_E_COMMANDNOTSUPPORTED;
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::Close
//
//  Close down the object and disconnect us from any handlers/proxies.
//
STDMETHODIMP CDsoDocObject::Close()
{
	HRESULT hr;

	ODS("CDsoDocObject::Close\n");

	// Go ahead an IP deactivate the object...
	hr = IPDeactivateView();

	SAFE_RELEASE_INTERFACE(m_pcmdt);
	SAFE_RELEASE_INTERFACE(m_pdocv);
	SAFE_RELEASE_INTERFACE(m_pipactive);
	SAFE_RELEASE_INTERFACE(m_pipobj);

	// Release the OLE object and cleanup...
	if (m_pole)
	{
		// Free running lock if we set it...
		FreeRunningLock();

		// Tell server to release our client site...
		m_pole->SetClientSite(NULL);

		// Finally, close the object and release the pointer...
		hr = m_pole->Close(OLECLOSE_NOSAVE);

		if (m_fDisconnectOnQuit)
			CoDisconnectObject((IUnknown*)m_pole, 0);

		SAFE_RELEASE_INTERFACE(m_pole);
	}

	SAFE_RELEASE_INTERFACE(m_pstgSourceFile);
	SAFE_RELEASE_INTERFACE(m_pmkSourceFile);
	SAFE_RELEASE_INTERFACE(m_pbctxSourceFile);

	SAFE_FREESTRING(m_pwszSourceFile);

	if (m_fDisconnectOnQuit)
	{
		CoDisconnectObject((IUnknown*)this, 0);
		m_fDisconnectOnQuit = FALSE;
	}

	SAFE_RELEASE_INTERFACE(m_pstmview);
	SAFE_RELEASE_INTERFACE(m_pstgfile);

	return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject Notification Functions - The OCX should call these to
//  let the doc site update the object as needed.
//  

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifySizeChange
//
//  Resets the size of the site window and tells UI active object to 
//  resize as well. If we are UI active, we'll call ResizeBorder to 
//  re-negotiate toolspace (allow toolbars to shrink and grow), otherwise
//  we'll just set the IP active view rect (minus any toolspace, which
//  should be none since object is not UI active!). 
//
void CDsoDocObject::OnNotifySizeChange(LPRECT prc)
{
	RECT rc;

	SetRect(&rc, 0, 0, (prc->right - prc->left), (prc->bottom - prc->top));
	if (rc.right < 0) rc.right = 0;
	if (rc.top < 0) rc.top = 0;

	// First, resize our frame site window tot he new size (don't change focus)...
	if (m_hwnd)
	{
		m_rcViewRect = *prc;

		SetWindowPos(m_hwnd, NULL, m_rcViewRect.left, m_rcViewRect.top,
			rc.right, rc.bottom, SWP_NOACTIVATE | SWP_NOZORDER);

		UpdateWindow(m_hwnd);
	}

	// If we have an active object (i.e., Document is still UI active) we should
	// tell it of the resize so it can re-negotiate border space...
	if ((m_fObjectUIActive) && (m_pipactive))
	{
		m_pipactive->ResizeBorder(&rc, (IOleInPlaceUIWindow*)&m_xOleInPlaceFrame, TRUE);
	}
	else if ((m_fObjectIPActive) && (m_pdocv))
	{
		rc.left += m_bwToolSpace.left;   rc.right -= m_bwToolSpace.right;
		rc.top += m_bwToolSpace.top;    rc.bottom -= m_bwToolSpace.bottom;
		m_pdocv->SetRect(&rc);
	}

	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifyAppActivate
//
//  Notify doc object when the top-level frame window goes active and 
//  deactive so it can handle window focs and paiting correctly. Failure
//  to not forward this notification leads to bad behavior.
// 
void CDsoDocObject::OnNotifyAppActivate(BOOL fActive, DWORD dwThreadID)
{
	// This is critical for DocObject servers, so forward these messages
	// when the object is UI active...
	if (m_pipactive)
	{
		// We should always tell obj server when our frame activates, but
		// don't tell it to go deactive if the thread gaining focus is 
		// the server's since our frame may have lost focus because of
		// a top-level modeless dialog (ex., the RefEdit dialog of Excel)...
		if (!m_fObjectInModalCondition)
			m_pipactive->OnFrameWindowActivate(fActive);
	}

	m_fAppWindowActive = fActive;
}

void CDsoDocObject::OnNotifyControlFocus(BOOL fGotFocus)
{
	HWND hwnd;

	// For UI Active DocObject server, we will tell them we gain/lose control focus...
	if ((m_pipactive) && (!m_fObjectInModalCondition))
	{
		// TODO: Normally we would notify object of loss of control focus (such as user
		// moving focus from framer control to another text box or something on the same
		// form), but this can cause PPT to drop its toolbar and XL to drop its formula
		// bar, so we are skipping this call. We really should determine if it is needed
		// for another host (like Visio or something?) but that is to do...
		//
		// m_pipactive->OnDocWindowActivate(fGotFocus);

		// We should forward the focus to the active window if window with the focus
		// is not already one that is parented to us...
		if ((fGotFocus) && !IsWindowChild(m_hwnd, GetFocus()) &&
			SUCCEEDED(m_pipactive->GetWindow(&hwnd)))
			SetFocus(hwnd);
	}
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifyPaletteChanged
//
//  Give the object first chance at realizing a palette. Important on
//  256 color machines, but not so critical these days when everyone is
//  running full 32-bit True Color graphic cards.
// 
void CDsoDocObject::OnNotifyPaletteChanged(HWND hwndPalChg)
{
	if ((m_fObjectUIActive) && (m_hwndUIActiveObj))
		SendMessage(m_hwndUIActiveObj, WM_PALETTECHANGED, (WPARAM)hwndPalChg, 0L);
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifyChangeToolState
//
//  This should be called to get object to show/hide toolbars as needed.
//
void CDsoDocObject::OnNotifyChangeToolState(BOOL fShowTools)
{
	// If we want to show/hide toolbars, we can do the following...
	if (fShowTools != (BOOL)m_fDisplayTools)
	{
		OLECMD cmd;
		cmd.cmdID = OLECMDID_HIDETOOLBARS;

		m_fDisplayTools = fShowTools;

		// Use IOleCommandTarget(OLECMDID_HIDETOOLBARS) to toggle on/off. We have
		// to check that server supports it and if its state matches our own so
		// when toggle, we do the correct thing by the user...
		if ((m_pcmdt) && SUCCEEDED(m_pcmdt->QueryStatus(NULL, 1, &cmd, NULL)) &&
			((cmd.cmdf & OLECMDF_SUPPORTED) || (cmd.cmdf & OLECMDF_ENABLED)))
		{
			if (((m_fDisplayTools) && ((cmd.cmdf & OLECMDF_LATCHED) == OLECMDF_LATCHED)) ||
				(!(m_fDisplayTools) && !((cmd.cmdf & OLECMDF_LATCHED) == OLECMDF_LATCHED)))
			{
				m_pcmdt->Exec(NULL, OLECMDID_HIDETOOLBARS, OLECMDEXECOPT_PROMPTUSER, NULL, NULL);
			}
			else if (m_fDisplayTools && (IsWordObject() || IsPPTObject()))
			{
				// HACK: Word and PowerPoint 2007 do not report the correct latched state for the ribbon,
				// so if we are trying to show the tools after hiding them, they will say they are already
				// visible when the ribbon is not. This is an apparant trick they are using to force the 
				// ribbon visible in IE embed cases, but it is causing us problems. 
				m_pcmdt->Exec(NULL, OLECMDID_HIDETOOLBARS, OLECMDEXECOPT_PROMPTUSER, NULL, NULL);
			}

			// There can be focus issues when turning them off, so make sure
			// the object is on top of the z-order...
			if ((!m_fDisplayTools) && (m_hwndIPObject))
				BringWindowToTop(m_hwndIPObject);
		}
		else if (m_pdocv)
		{
			// If we have a DocObj server, but no IOleCommandTarget, do things the hard
			// way and resize. When server attempts to resize window it will have to
			// re-negotiate BorderSpace and we fail there, so server "should" not
			// display its tools (at least that is the idea!<g>)...
			RECT rc; GetClientRect(m_hwnd, &rc);
			MapWindowPoints(m_hwnd, m_hwndCtl, (LPPOINT)&rc, 2);
			OnNotifySizeChange(&rc);
		}
	}
	return;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject Protected Functions -- Helpers
//

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::InstantiateDocObjectServer (protected)
//
//  Startup OLE Document Object server and get IOleObject pointer.
//
STDMETHODIMP CDsoDocObject::InstantiateDocObjectServer(REFCLSID rclsid, IOleObject **ppole)
{
	HRESULT hr;
	IUnknown *punk = NULL;

	ODS("CDsoDocObject::InstantiateDocObjectServer()\n");

	// We perform custom create in order of local server then inproc server...
	if (SUCCEEDED(hr = CoCreateInstance(rclsid, NULL, CLSCTX_LOCAL_SERVER, IID_IUnknown, (void**)&punk)) ||
		SUCCEEDED(hr = CoCreateInstance(rclsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&punk)))
	{
		// Ask for IOleObject interface...
		if (SUCCEEDED(hr = punk->QueryInterface(IID_IOleObject, (void**)ppole)))
		{
			// TODO: Add strong connection to remote server (IExternalConnection??)...
		}
		punk->Release();
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateObjectStorage (protected)
//
//  Makes the internal IStorage to host the object and assigns the CLSID.
//
STDMETHODIMP CDsoDocObject::CreateObjectStorage(REFCLSID rclsid)
{
	HRESULT hr;
	LPWSTR pwszName;
	DWORD dwid;
	WCHAR szbuf[256];

	if ((!m_pstgroot)) return E_UNEXPECTED;

	// Next, create a new object storage (with unique name) in our
	// temp root storage "file" (this keeps an OLE integrity some servers
	// need to function correctly instead of IP activating from file directly).

	// We make a fake object storage name...
	dwid = ((rclsid.Data1) | GetTickCount());
	wsprintf(szbuf, L"OLEDocument%X", dwid);

	pwszName = szbuf;

	// Create the sub-storage...
	hr = m_pstgroot->CreateStorage(pwszName,
		STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &m_pstgfile);


	if (FAILED(hr)) return hr;

	// We'll also create a stream for OLE view settings (non-critical)...

	m_pstgroot->CreateStream(pwszName,
		STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &m_pstmview);


	// Finally, write out the CLSID for the new substorage...
	hr = WriteClassStg(m_pstgfile, rclsid);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveObjectStorage (protected)
//
//  Saves the object back to the internal IStorage. Returns S_FALSE if
//  there is no file storage for this document (it depends on how file
//  was loaded). In most cases we should have one, and this copies data
//  into the internal storage for save.
//
STDMETHODIMP CDsoDocObject::SaveObjectStorage()
{
	HRESULT hr = S_FALSE;
	IPersistStorage *pipstg = NULL;

	// Got to have object to save state...
	if (!m_pole) return E_UNEXPECTED;

	// If we have file storage, ask for IPersist and Save (commit changes)...
	if ((m_pstgfile) &&
		SUCCEEDED(hr = m_pole->QueryInterface(IID_IPersistStorage, (void**)&pipstg)))
	{
		if (SUCCEEDED(hr = pipstg->Save(m_pstgfile, TRUE)))
			hr = pipstg->SaveCompleted(NULL);

		hr = m_pstgfile->Commit(STGC_DEFAULT);
		pipstg->Release();
	}

	// Go ahead and save the view state if view still active (non-critical)...
	if ((m_pdocv) && (m_pstmview))
	{
		m_pdocv->SaveViewState(m_pstmview);
		m_pstmview->Commit(STGC_DEFAULT);
	}

	return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveDocToMoniker (protected)
//
//  Saves document to location spcified by the moniker passed. The server
//  needs to support either IPersistMoniker or IPersistFile.  This is used
//  primarily when document was opened by moniker instead of storage.
//
STDMETHODIMP CDsoDocObject::SaveDocToMoniker(IMoniker *pmk, IBindCtx *pbc, BOOL fKeepLock)
{
	HRESULT hr = E_FAIL;
	IPersistMoniker *prstmk;
	LPOLESTR pwszFullName = NULL;

	CHECK_NULL_RETURN(m_pole, E_UNEXPECTED);
	CHECK_NULL_RETURN(pmk, E_POINTER);

	// Get IPersistMoniker interface and ask it to save context if it existing moniker...
	if (SUCCEEDED(hr = m_pole->QueryInterface(IID_IPersistMoniker, (void**)&prstmk)))
	{
		hr = prstmk->Save(pmk, pbc, fKeepLock);
		prstmk->Release();
	}

	// If that failed to work, switch to IPersistFile and use full path to the new file...
	if (FAILED(hr) && SUCCEEDED(hr = pmk->GetDisplayName(pbc, NULL, &pwszFullName)))
	{
		hr = SaveDocToFile(pwszFullName, fKeepLock);
		CoTaskMemFree(pwszFullName);
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveDocToFile (protected)
//
//  Saves document to location spcified by the path passed. The server
//  needs to support either IPersistFile. 
//
STDMETHODIMP CDsoDocObject::SaveDocToFile(LPWSTR pwszFullName, BOOL fKeepLock)
{
	HRESULT hr = E_FAIL;
	IPersistFile *pipfile = NULL;

	BOOL fSameAsOpen = ((m_pwszSourceFile) &&
		DsoCompareStringsEx(pwszFullName, lstrlenW(pwszFullName),
			m_pwszSourceFile, lstrlenW(m_pwszSourceFile)) == CSTR_EQUAL);

	if (SUCCEEDED(hr = m_pole->QueryInterface(IID_IPersistFile, (void**)&pipfile)))
	{
#ifdef DSO_WORD12_PERSIST_BUG
		LPOLESTR pwszWordCurFile = NULL;
		// HACK: Word 2007 RTM has a bug in its IPersistFile::Save method which will cause it to save
		// to the old file location and not the new one, so if the document being saved is a Word
		// object and we are saving to a new file, then got to run this hack to copy the saved bits
		// to the correct location...
		if (IsWordObject())
		{
			hr = pipfile->GetCurFile(&pwszWordCurFile);
			if ((fSameAsOpen) && (pwszWordCurFile))
			{
				// Check again if Word thinks the file is the same as we do...
				fSameAsOpen = (DsoCompareStringsEx(pwszFullName, lstrlenW(pwszFullName),
					pwszWordCurFile, lstrlenW(pwszWordCurFile)) == CSTR_EQUAL);
				if (fSameAsOpen)
				{ // If it is the same file after all, we don't need to do the hack...
					CoTaskMemFree(pwszWordCurFile);
					pwszWordCurFile = NULL;
				}
			}
		}
#endif

		// Do the save using file path or NULL if we are saving to current file...
		hr = pipfile->Save((fSameAsOpen ? NULL : pwszFullName), fKeepLock);

#ifdef DSO_WORD12_PERSIST_BUG
		// HACK: If we have the pwszWordCurFile, we'll assume Word 12 RTM might have saved this, and we
		// need to check if Word saved to the right path or not. So check the paths, and if the new 
		// file is not there, copy the file Word saved to the new location...
		if (pwszWordCurFile)
		{
			if (SUCCEEDED(hr) && !FFileExists(pwszFullName) && FFileExists(pwszWordCurFile))
			{
				if (!FPerformShellOp(FO_COPY, pwszWordCurFile, pwszFullName))
					hr = E_ACCESSDENIED;
			}
			CoTaskMemFree(pwszWordCurFile);
		}
		// Hopefully, Word should have this fixed by Office 2007 SP1 and we can remove this hack!
#endif
		pipfile->Release();
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::ValidateDocObjectServer (protected)
//
//  Quick validation check to see if CLSID is for DocObject server.
//
//  Officially, the only way to determine if a server supports ActiveX
//  Document embedding is to IP activate it and ask for IOleDocument, 
//  but that means going through the IP process just to fail if IOleDoc
//  is not supported. Therefore, we are going to rely on the server's 
//  honesty in setting its reg keys to include the "DocObject" sub key 
//  under their CLSID.
//
//  This is 99% accurate. For those servers that fail, too bad charlie!
//
STDMETHODIMP CDsoDocObject::ValidateDocObjectServer(REFCLSID rclsid)
{
	HRESULT hr = DSO_E_INVALIDSERVER;
	WCHAR  szKeyCheck[256];
	LPWSTR pszClsid;
	HKEY  hkey;

	// We don't handle MSHTML even though it is DocObject server. If you plan
	// to view web pages in browser-like context, best to use WebBrowser control...
	if (rclsid == CLSID_MSHTML_DOCUMENT)
		return hr;

	// Convert the CLSID to a string and check for DocObject sub key...
	if (pszClsid = DsoCLSIDtoLPWSTR(rclsid))
	{
		wsprintf(szKeyCheck, L"CLSID\\%s\\DocObject", pszClsid);

		if (RegOpenKeyEx(HKEY_CLASSES_ROOT, szKeyCheck, 0, KEY_READ, &hkey) == ERROR_SUCCESS)
		{
			hr = S_OK;
			RegCloseKey(hkey);
		}

		DsoMemFree(pszClsid);
	}
	else hr = E_OUTOFMEMORY;

	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::EnsureOleServerRunning (protected)
//
//  Verifies the DocObject server is running and is optionally locked
//  while the embedding is taking place.
//
STDMETHODIMP CDsoDocObject::EnsureOleServerRunning(BOOL fLockRunning)
{
	HRESULT hr = S_FALSE;
	IBindCtx *pbc;
	IRunnableObject *pro;
	IOleContainer *pocnt;
	BIND_OPTS bnd = { sizeof(BIND_OPTS), BIND_MAYBOTHERUSER, (STGM_READWRITE | STGM_SHARE_EXCLUSIVE), 10000 };

	TRACE1("CDsoDocObject::EnsureOleServerRunning(%d)\n", (DWORD)fLockRunning);
	if (m_pole == NULL) return E_FAIL;

	// If we are already locked, don't need to do this again...
	if (m_fLockedServerRunning)
		return hr;

	// Create a bind ctx...
	if (FAILED(CreateBindCtx(0, &pbc)))
		return E_UNEXPECTED;

	// Setup default bind options for the run operation...
	pbc->SetBindOptions(&bnd);

	// Get IRunnableObject and set server to run as OLE object. We check the
	// running state first since this is proper OLE, but note that this is not
	// returned from out-of-proc server, but the in-proc handler. Also note, we
	// specify a timeout in case the object never starts, and "check" the object
	// runs without IMessageFilter errors...
	if (SUCCEEDED(m_pole->QueryInterface(IID_IRunnableObject, (void**)&pro)))
	{

		// If the object is not currently running, let's run it...
		if (!(pro->IsRunning()))
			hr = pro->Run(pbc);

		// Set the object server as a contained object (i.e., OLE object)...
		pro->SetContainedObject(TRUE);

		// Lock running if desired...
		if (fLockRunning)
			m_fLockedServerRunning = SUCCEEDED(pro->LockRunning(TRUE, TRUE));

		pro->Release();
	}
	else if (SUCCEEDED(m_pole->QueryInterface(IID_IOleContainer, (void**)&pocnt)))
	{
		if (fLockRunning)
			m_fLockedServerRunning = SUCCEEDED(pocnt->LockContainer(TRUE));

		pocnt->Release();
	}

	pbc->Release();
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::FreeRunningLock (protected)
//
//  Free any previous lock made by EnsureOleServerRunning.
//
void CDsoDocObject::FreeRunningLock()
{
	IRunnableObject *pro;
	IOleContainer *pocnt;

	ODS("CDsoDocObject::FreeRunningLock(%d)\n");
	ASSERT(m_pole);

	// Don't do anything if we didn't lock the server...
	if (m_fLockedServerRunning == FALSE)
		return;

	// Get IRunnableObject and free lock...
	if (SUCCEEDED(m_pole->QueryInterface(IID_IRunnableObject, (void**)&pro)))
	{
		pro->LockRunning(FALSE, TRUE);
		pro->Release();
	}
	else if (SUCCEEDED(m_pole->QueryInterface(IID_IOleContainer, (void**)&pocnt)))
	{
		pocnt->LockContainer(FALSE);
		pocnt->Release();
	}

	m_fLockedServerRunning = FALSE;
	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SetRunningServerLock 
//
//  Sets an external lock on the running object server.
//
STDMETHODIMP CDsoDocObject::SetRunningServerLock(BOOL fLock)
{
	HRESULT hr;
	TRACE1("CDsoDocObject::SetRunningServerLock(%d)\n", (DWORD)fLock);

	// Word doesn't obey normal running lock, but does have method to explicitly
	// lock the server for mail, so we'll use that in the Word case...
	if (IsWordObject())
	{
		IClassFactory *pcf = NULL;
		interface ifoo : public IUnknown
		{
			STDMETHOD(_uncall)() PURE;
			STDMETHOD(lock)(BOOL f) PURE;
		} *pi = NULL;
		const GUID iidifoo = { 0x0006729A, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };

		SEH_TRY
			hr = CoGetClassObject(CLSID_WORD_DOCUMENT_DOC, CLSCTX_LOCAL_SERVER, NULL, IID_IClassFactory, (void**)&pcf);
		if (SUCCEEDED(hr) && (pcf))
		{
			hr = pcf->QueryInterface(iidifoo, (void**)&pi);
			if (SUCCEEDED(hr) && (pi))
			{
				hr = pi->lock(fLock);
				pi->Release();
			}
			pcf->Release();
		}
		SEH_EXCEPT(hr)
	}
	else
	{
		if (fLock)
		{
			hr = EnsureOleServerRunning(TRUE);
		}
		else
		{
			hr = (FreeRunningLock(), S_OK);
		}
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
//
// ActiveX Document Site Interfaces
//

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject IUnknown Interface Methods
//
//   STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
//   STDMETHODIMP_(ULONG) AddRef(void);
//   STDMETHODIMP_(ULONG) Release(void);
//
STDMETHODIMP CDsoDocObject::QueryInterface(REFIID riid, void** ppv)
{
	ODS("CDsoDocObject::QueryInterface\n");
	CHECK_NULL_RETURN(ppv, E_POINTER);

	HRESULT hr = S_OK;

	if (IID_IUnknown == riid)
	{
		*ppv = (IUnknown*)this;
	}
	else if (IID_IOleClientSite == riid)
	{
		*ppv = (IOleClientSite*)&m_xOleClientSite;
	}
	else if ((IID_IOleInPlaceSite == riid) || (IID_IOleWindow == riid))
	{
		*ppv = (IOleInPlaceSite*)&m_xOleInPlaceSite;
	}
	else if (IID_IOleDocumentSite == riid)
	{
		*ppv = (IOleDocumentSite*)&m_xOleDocumentSite;
	}
	else if ((IID_IOleInPlaceFrame == riid) || (IID_IOleInPlaceUIWindow == riid))
	{
		*ppv = (IOleInPlaceFrame*)&m_xOleInPlaceFrame;
	}
	else if (IID_IOleCommandTarget == riid)
	{
		*ppv = (IOleCommandTarget*)&m_xOleCommandTarget;
	}
	else if (IID_IServiceProvider == riid)
	{
		*ppv = (IServiceProvider*)&m_xServiceProvider;
	}
	else
	{
		*ppv = NULL;
		hr = E_NOINTERFACE;
	}

	if (NULL != *ppv)
		((IUnknown*)(*ppv))->AddRef();

	return hr;
}

ULONG CDsoDocObject::AddRef(void)
{
	TRACE1("CDsoDocObject::AddRef - %d\n", m_cRef + 1);
	return ++m_cRef;
}

ULONG CDsoDocObject::Release(void)
{
	TRACE1("CDsoDocObject::Release - %d\n", m_cRef - 1);
	return --m_cRef;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleClientSite -- IOleClientSite Implementation
//
//	 STDMETHODIMP SaveObject(void);
//	 STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhich, LPMONIKER* ppmk);
//	 STDMETHODIMP GetContainer(LPOLECONTAINER* ppContainer);
//	 STDMETHODIMP ShowObject(void);
//	 STDMETHODIMP OnShowWindow(BOOL fShow);
//	 STDMETHODIMP RequestNewObjectLayout(void);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleClientSite)

STDMETHODIMP CDsoDocObject::XOleClientSite::SaveObject(void)
{
	ODS("CDsoDocObject::XOleClientSite::SaveObject\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk)
{
	HRESULT hr = OLE_E_CANT_GETMONIKER;
	METHOD_PROLOGUE(CDsoDocObject, OleClientSite);

	TRACE2("CDsoDocObject::XOleClientSite::GetMoniker(%d, %d)\n", dwAssign, dwWhichMoniker);
	CHECK_NULL_RETURN(ppmk, E_POINTER);
	*ppmk = NULL;

	// Provide moniker for object opened by moniker...
	switch (dwWhichMoniker)
	{
	case OLEWHICHMK_OBJREL:
	case OLEWHICHMK_OBJFULL:
	{
		if (pThis->m_pmkSourceFile)
		{
			*ppmk = pThis->m_pmkSourceFile;
			hr = S_OK;
		}
		else if (dwAssign == OLEGETMONIKER_FORCEASSIGN)
		{
			// TODO: Should we allow force create of moniker if we don't start with it?
		}
	}
	break;
	}

	// Need to AddRef returned value to caller on success...
	if ((hr == S_OK) && (*ppmk))
		((IUnknown*)*ppmk)->AddRef();

	return hr;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::GetContainer(IOleContainer** ppContainer)
{
	ODS("CDsoDocObject::XOleClientSite::GetContainer\n");
	if (ppContainer) *ppContainer = NULL;
	return E_NOINTERFACE;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::ShowObject(void)
{
	ODS("CDsoDocObject::XOleClientSite::ShowObject\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::OnShowWindow(BOOL fShow)
{
	ODS("CDsoDocObject::XOleClientSite::OnShowWindow\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::RequestNewObjectLayout(void)
{
	ODS("CDsoDocObject::XOleClientSite::RequestNewObjectLayout\n");
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleInPlaceSite -- IOleInPlaceSite Implementation
//
//	 STDMETHODIMP GetWindow(HWND* phWnd);
//	 STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
//	 STDMETHODIMP CanInPlaceActivate(void);
//	 STDMETHODIMP OnInPlaceActivate(void);
//	 STDMETHODIMP OnUIActivate(void);
//	 STDMETHODIMP GetWindowContext(LPOLEINPLACEFRAME* ppIIPFrame, LPOLEINPLACEUIWINDOW* ppIIPUIWindow, LPRECT prcPos, LPRECT prcClip, LPOLEINPLACEFRAMEINFO pFI);
//	 STDMETHODIMP Scroll(SIZE sz);
//	 STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
//	 STDMETHODIMP OnInPlaceDeactivate(void);
//	 STDMETHODIMP DiscardUndoState(void);
//	 STDMETHODIMP DeactivateAndUndo(void);
//	 STDMETHODIMP OnPosRectChange(LPCRECT prcPos);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleInPlaceSite)

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::GetWindow(HWND* phwnd)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::GetWindow\n");
	if (phwnd) *phwnd = pThis->m_hwnd;
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::ContextSensitiveHelp(BOOL fEnterMode)
{
	ODS("CDsoDocObject::XOleInPlaceSite::ContextSensitiveHelp\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::CanInPlaceActivate(void)
{
	ODS("CDsoDocObject::XOleInPlaceSite::CanInPlaceActivate\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnInPlaceActivate(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::OnInPlaceActivate\n");

	if ((!pThis->m_pole) ||
		FAILED(pThis->m_pole->QueryInterface(IID_IOleInPlaceObject, (void **)&(pThis->m_pipobj))))
		return E_UNEXPECTED;

	pThis->m_fObjectIPActive = TRUE;
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnUIActivate(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::OnUIActivate\n");
	pThis->m_fObjectUIActive = TRUE;
	if (pThis->m_pipobj)
	{
		pThis->m_pipobj->GetWindow(&(pThis->m_hwndIPObject));
	}

	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::GetWindowContext(IOleInPlaceFrame** ppFrame,
	IOleInPlaceUIWindow** ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::GetWindowContext\n");

	if (ppFrame)
	{
		SAFE_SET_INTERFACE(*ppFrame, &(pThis->m_xOleInPlaceFrame));
	}

	if (ppDoc)
		*ppDoc = NULL;

	if (lprcPosRect)
		*lprcPosRect = pThis->m_rcViewRect;

	if (lprcClipRect)
		*lprcClipRect = *lprcPosRect;

	memset(lpFrameInfo, 0, sizeof(OLEINPLACEFRAMEINFO));
	lpFrameInfo->cb = sizeof(OLEINPLACEFRAMEINFO);
	lpFrameInfo->hwndFrame = pThis->m_hwnd;
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::Scroll(SIZE sz)
{
	ODS("CDsoDocObject::XOleInPlaceSite::Scroll\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnUIDeactivate(BOOL fUndoable)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::OnUIDeactivate\n");

	pThis->m_fObjectUIActive = FALSE;
	pThis->m_xOleInPlaceFrame.SetMenu(NULL, NULL, NULL);
	SetFocus(pThis->m_hwnd);

	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnInPlaceDeactivate(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::OnInPlaceDeactivate\n");

	pThis->m_fObjectIPActive = FALSE;
	pThis->m_hwndIPObject = NULL;
	SAFE_RELEASE_INTERFACE((pThis->m_pipobj));

	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::DiscardUndoState(void)
{
	ODS("CDsoDocObject::XOleInPlaceSite::DiscardUndoState\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::DeactivateAndUndo(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
	ODS("CDsoDocObject::XOleInPlaceSite::DeactivateAndUndo\n");
	if (pThis->m_pipobj) pThis->m_pipobj->InPlaceDeactivate();
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnPosRectChange(LPCRECT lprcPosRect)
{
	ODS("CDsoDocObject::XOleInPlaceSite::OnPosRectChange\n");
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleDocumentSite -- IOleDocumentSite Implementation
//
//	 STDMETHODIMP ActivateMe(IOleDocumentView* pView);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleDocumentSite)

STDMETHODIMP CDsoDocObject::XOleDocumentSite::ActivateMe(IOleDocumentView* pView)
{
	METHOD_PROLOGUE(CDsoDocObject, OleDocumentSite);
	ODS("CDsoDocObject::XOleDocumentSite::ActivateMe\n");

	HRESULT             hr = E_FAIL;
	IOleDocument*       pmsodoc;

	// If we're passed a NULL view pointer, then try to get one from
	// the document object (the object within us).
	if (pView)
	{
		// Make sure that the view has our client site
		hr = pView->SetInPlaceSite((IOleInPlaceSite*)&(pThis->m_xOleInPlaceSite));

		pView->AddRef(); // we will be keeping the object if successful..
	}
	else if (pThis->m_pole)
	{
		// Create a new view from the OleDocument...
		if (FAILED(pThis->m_pole->QueryInterface(IID_IOleDocument, (void **)&pmsodoc)))
			return E_FAIL;

		hr = pmsodoc->CreateView((IOleInPlaceSite*)&(pThis->m_xOleInPlaceSite),
			pThis->m_pstmview, 0, &pView);

		pmsodoc->Release();
	}

	// If we have the view, apply view state...
	if (SUCCEEDED(hr) && (pThis->m_pstmview))
		hr = pView->ApplyViewState(pThis->m_pstmview);

	// If any of the above failed, release the view and return...
	if (FAILED(hr))
	{
		SAFE_RELEASE_INTERFACE(pView);
		return hr;
	}

	// keep the view pointer...
	pThis->m_pdocv = pView;

	// Get a command target (if available)...
	pView->QueryInterface(IID_IOleCommandTarget, (void**)&(pThis->m_pcmdt));

	// Make sure that the view has our client site
	pView->SetInPlaceSite((IOleInPlaceSite*)&(pThis->m_xOleInPlaceSite));

	// This sets up toolbars and menus first    
	if (SUCCEEDED(hr = pView->UIActivate(TRUE)))
	{
		// Set the window size sensitive to new toolbars
		pView->SetRect(&(pThis->m_rcViewRect));

		// Makes it all active
		pView->Show(TRUE);

		pThis->m_fAppWindowActive = TRUE;

		// Toogle tools off if that's what user wants...
		if (!(pThis->m_fDisplayTools))
		{
			pThis->m_fDisplayTools = TRUE;
			pThis->OnNotifyChangeToolState(FALSE);
		}
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleInPlaceFrame -- IOleInPlaceFrame Implementation
//
//   STDMETHODIMP GetWindow(HWND* phWnd);
//   STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
//   STDMETHODIMP GetBorder(LPRECT prcBorder);
//   STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS pBW);
//   STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS pBW);
//   STDMETHODIMP SetActiveObject(LPOLEINPLACEACTIVEOBJECT pIIPActiveObj, LPCOLESTR pszObj);
//   STDMETHODIMP InsertMenus(HMENU hMenu, LPOLEMENUGROUPWIDTHS pMGW);
//   STDMETHODIMP SetMenu(HMENU hMenu, HOLEMENU hOLEMenu, HWND hWndObj);
//   STDMETHODIMP RemoveMenus(HMENU hMenu);
//   STDMETHODIMP SetStatusText(LPCOLESTR pszText);
//   STDMETHODIMP EnableModeless(BOOL fEnable);
//   STDMETHODIMP TranslateAccelerator(LPMSG pMSG, WORD wID);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleInPlaceFrame)

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::GetWindow(HWND* phWnd)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	ODS("CDsoDocObject::XOleInPlaceFrame::GetWindow\n");
	return pThis->m_xOleInPlaceSite.GetWindow(phWnd);
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::ContextSensitiveHelp(BOOL fEnterMode)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	ODS("CDsoDocObject::XOleInPlaceFrame::ContextSensitiveHelp\n");
	return pThis->m_xOleInPlaceSite.ContextSensitiveHelp(fEnterMode);
}


STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::GetBorder(LPRECT prcBorder)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	ODS("CDsoDocObject::XOleInPlaceFrame::GetBorder\n");
	CHECK_NULL_RETURN(prcBorder, E_POINTER);

	// If we don't allow Toolspace, and we are already active, give
	// no space for tools (ie, hide toolabrs), otherwise give as much we can...
	if (!(pThis->m_fDisplayTools) && (pThis->m_pipactive))
		SetRectEmpty(prcBorder);
	else
		GetClientRect(pThis->m_hwnd, prcBorder);

	TRACE_LPRECT("prcBorder", prcBorder);
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::RequestBorderSpace(LPCBORDERWIDTHS pBW)
{
	ODS("CDsoDocObject::XOleInPlaceFrame::RequestBorderSpace\n");
	CHECK_NULL_RETURN(pBW, E_POINTER);
	TRACE_LPRECT("pBW", pBW);
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetBorderSpace(LPCBORDERWIDTHS pBW)
{
	RECT rc;

	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	ODS("CDsoDocObject::XOleInPlaceFrame::SetBorderSpace\n");

	if (pBW) { TRACE_LPRECT("pBW", pBW); }

	GetClientRect(pThis->m_hwnd, &rc);
	SetRectEmpty((RECT*)&(pThis->m_bwToolSpace));

	if (pBW)
	{
		pThis->m_bwToolSpace = *pBW;
		rc.left += pBW->left;   rc.right -= pBW->right;
		rc.top += pBW->top;    rc.bottom -= pBW->bottom;
	}

	// Save the current view RECT (space minus tools)...
	pThis->m_rcViewRect = rc;

	// Update the active document (if alive)...
	if (pThis->m_pdocv)
		pThis->m_pdocv->SetRect(&(pThis->m_rcViewRect));

	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetActiveObject(LPOLEINPLACEACTIVEOBJECT pIIPActiveObj, LPCOLESTR pszObj)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	ODS("CDsoDocObject::XOleInPlaceFrame::SetActiveObject\n");

	SAFE_RELEASE_INTERFACE((pThis->m_pipactive));
	pThis->m_hwndUIActiveObj = NULL;
	pThis->m_dwObjectThreadID = 0;

	if (pIIPActiveObj)
	{
		SAFE_SET_INTERFACE(pThis->m_pipactive, pIIPActiveObj);
		pIIPActiveObj->GetWindow(&(pThis->m_hwndUIActiveObj));
		pThis->m_dwObjectThreadID = GetWindowThreadProcessId(pThis->m_hwndUIActiveObj, NULL);
	}

	return S_OK;
}

void CDsoDocObject::ReObtainActiveWindow()
{
	m_pipactive->GetWindow(&m_hwndUIActiveObj);
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::InsertMenus(HMENU hMenu, LPOLEMENUGROUPWIDTHS pMGW)
{
	ODS("CDsoDocObject::XOleInPlaceFrame::InsertMenus\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetMenu(HMENU hMenu, HOLEMENU hOLEMenu, HWND hWndObj)
{
	ODS("CDsoDocObject::XOleInPlaceFrame::SetMenu\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::RemoveMenus(HMENU hMenu)
{
	ODS("CDsoDocObject::XOleInPlaceFrame::RemoveMenus\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetStatusText(LPCOLESTR pszText)
{
	//METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	ODS("CDsoDocObject::XOleInPlaceFrame::SetStatusText\n");
	if ((pszText) && (*pszText)) { TRACE1(" Status Text = %S \n", pszText); }
	return S_OK; //  ((pThis->m_psiteCtl) ? pThis->m_psiteCtl->SetStatusText(pszText) : S_OK);
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::EnableModeless(BOOL fEnable)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
	TRACE1("CDsoDocObject::XOleInPlaceFrame::EnableModeless(%d)\n", fEnable);
	pThis->m_fObjectInModalCondition = !fEnable;
	SendMessage(pThis->m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_MODAL, (LPARAM)fEnable);
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::TranslateAccelerator(LPMSG pMSG, WORD wID)
{
	ODS("CDsoDocObject::XOleInPlaceFrame::TranslateAccelerator\n");
	return S_FALSE;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleCommandTarget -- IOleCommandTarget Implementation
//
//   STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
//   STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);            
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleCommandTarget)

STDMETHODIMP CDsoDocObject::XOleCommandTarget::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
	HRESULT hr = OLECMDERR_E_UNKNOWNGROUP;
	METHOD_PROLOGUE(CDsoDocObject, OleCommandTarget);
	if (pThis->m_pcmdCtl)
		hr = pThis->m_pcmdCtl->QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
	return hr;
}

STDMETHODIMP CDsoDocObject::XOleCommandTarget::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut)
{
	HRESULT hr = OLECMDERR_E_NOTSUPPORTED;
	METHOD_PROLOGUE(CDsoDocObject, OleCommandTarget);
	if (pThis->m_pcmdCtl)
		hr = pThis->m_pcmdCtl->Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
	return hr;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XServiceProvider -- IServiceProvider Implementation
//
//   STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, ServiceProvider)

STDMETHODIMP CDsoDocObject::XServiceProvider::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	ODS("CDsoDocObject::XServiceProvider::QueryService\n");

	return E_NOINTERFACE;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::FrameWindowProc
//
//  Site window procedure. Not much to do here except forward focus.
//
LRESULT CDsoDocObject::FrameWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	CDsoDocObject* pbndr = (CDsoDocObject*)GetWindowLong(hwnd, GWL_USERDATA);
	if (pbndr)
	{
		switch (msg)
		{
		case WM_NCDESTROY:
			SetWindowLong(hwnd, GWL_USERDATA, 0);
			pbndr->m_hwnd = NULL;
			break;

		case WM_SETFOCUS:
			if (pbndr->m_hwndUIActiveObj)
				SetFocus(pbndr->m_hwndUIActiveObj);
			return 0;

		case WM_ERASEBKGND:
			return 1;
		}
	}

	return DefWindowProc(hwnd, msg, wParam, lParam);
}