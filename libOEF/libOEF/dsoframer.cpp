/***************************************************************************
 * DSOFCONTROL.CPP
 *
 * CDsoFramerControl: The Base Control
 ***************************************************************************/

#include "dsoframer.h"


CDsoFramerControl::CDsoFramerControl()
{
	m_pDocObjFrame = nullptr;
}

CDsoFramerControl::~CDsoFramerControl(void)
{
	ODS("CDsoFramerControl::~CDsoFramerControl\n");
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
HRESULT CDsoFramerControl::Open(LPWSTR pwszDocument, BOOL fOpenReadOnly, LPWSTR pwszAltProgId, HWND hwndParent, RECT dstRect)
{
	HRESULT   hr;
	CLSID     clsidAlt = GUID_NULL;

	BIND_OPTS bopts = { sizeof(BIND_OPTS), BIND_MAYBOTHERUSER, 0, 10000 };

	TRACE2("CDsoFramerControl::Open(%S, %S)\n", pwszDocument, pwszAltProgId);

	// If the user passed the ProgId, find the alternative CLSID for server...
	if ((pwszAltProgId) && FAILED(CLSIDFromProgID(pwszAltProgId, &clsidAlt)))
	{
		return DSO_E_INVALIDPROGID;
	}

	// Let's make a doc frame for ourselves...
	if (!(m_pDocObjFrame = CDsoDocObject::CreateInstance(hwndParent, dstRect)))
	{
		return E_OUTOFMEMORY;
	}

	// Setup the bind options based on read-only flag....
	bopts.grfMode = (STGM_TRANSACTED | STGM_SHARE_DENY_WRITE | (fOpenReadOnly ? STGM_READ : STGM_READWRITE));

	SEH_TRY
		// Normally user gives a string that is path to file...
		hr = m_pDocObjFrame->CreateFromFile(pwszDocument, clsidAlt, &bopts);
		// If successful, we can activate the object...
		if (SUCCEEDED(hr))
		{
			hr = m_pDocObjFrame->IPActivateView();
		}
	SEH_EXCEPT(hr)
		// Force a close if an error occurred...
		if (FAILED(hr))
		{
			Close();
		}

	return hr;
}


//
void CDsoFramerControl::OnResize(RECT dstRect)
{
	ODS("CDsoFramerControl::OnResize\n");
	if (m_pDocObjFrame)
	{
		m_pDocObjFrame->OnNotifySizeChange(&dstRect);
	}
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

HWND CDsoFramerControl::GetActiveWindow()
{
	if (m_pDocObjFrame)
	{
		return m_pDocObjFrame->GetActiveWindow();
	}
	return NULL;
}

void CDsoFramerControl::ReobtainActiveFrame()
{
	if (m_pDocObjFrame)
	{
		return m_pDocObjFrame->ReObtainActiveWindow();
	}
}


////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::GetActiveDocument
//
//  Returns the automation object currently embedded.
//
//  Since we only support a single instance at a time, it might have been
//  better to call this property Object or simply Document, but VB reserves
//  the first name for use by the control extender, and IE reserves the second
//  in its extender, so we decided on the "Office sounding" name. ;-)
//
HRESULT CDsoFramerControl::GetActiveDocument(IDispatch** ppdisp)
{
	HRESULT hr = DSO_E_DOCUMENTNOTOPEN;
	IUnknown* punk;

	ODS("CDsoFramerControl::GetActiveDocument\n");
	CHECK_NULL_RETURN(ppdisp, E_POINTER); *ppdisp = NULL;

	// Get IDispatch from open document object.
	if ((m_pDocObjFrame) && (punk = (IUnknown*)(m_pDocObjFrame->GetActiveObject())))
	{
		// Ask ip active object for IDispatch interface. If it is not supported on
		// active object interface, try to get it from OLE object iface...
		if (FAILED(hr = punk->QueryInterface(IID_IDispatch, (void**)ppdisp)) &&
			(punk = (IUnknown*)(m_pDocObjFrame->GetOleObject())))
		{
			hr = punk->QueryInterface(IID_IDispatch, (void**)ppdisp);
		}
		ASSERT(SUCCEEDED(hr));
	}

	return hr;
}