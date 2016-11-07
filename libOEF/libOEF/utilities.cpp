/***************************************************************************
 * UTILITIES.CPP
 *
 * Shared helper functions and routines.
 ***************************************************************************/
#include "dsoframer.h"

////////////////////////////////////////////////////////////////////////
// Core Utility Functions
//
////////////////////////////////////////////////////////////////////////
// Heap Allocation (Private Heap)
//
extern HANDLE v_hPrivateHeap;
STDAPI_(LPVOID) DsoMemAlloc(DWORD cbSize)
{
    CHECK_NULL_RETURN(v_hPrivateHeap, NULL);
    return HeapAlloc(v_hPrivateHeap, HEAP_ZERO_MEMORY, cbSize);
}

STDAPI_(void) DsoMemFree(LPVOID ptr)
{
    if ((v_hPrivateHeap) && (ptr))
        HeapFree(v_hPrivateHeap, 0, ptr);
}

////////////////////////////////////////////////////////////////////////
// Global String Functions
//


//////////////////////////////////////////////////////////////////////
 //DsoConvertToBSTR

 // Takes a MBCS string and returns a BSTR. NULL is returned if the 
 // function fails or the string is empty.

STDAPI_(BSTR) DsoConvertToBSTR(LPWSTR pwsz)
{
	BSTR bstr = NULL;
	if (pwsz)
	{
	    bstr = SysAllocString(pwsz);
	    DsoMemFree(pwsz);
    }
	return bstr;
}

////////////////////////////////////////////////////////////////////////
// DsoConvertToLPOLESTR
//
//  Returns Unicode string in COM Task Memory (CoTaskMemAlloc).
//
STDAPI_(LPWSTR) DsoConvertToLPOLESTR(LPCWSTR pwszUnicodeString)
{
	LPWSTR pwsz;
	UINT cblen;

	CHECK_NULL_RETURN(pwszUnicodeString, NULL);
	cblen = lstrlenW(pwszUnicodeString);

    pwsz = (LPWSTR)CoTaskMemAlloc((cblen * sizeof(WCHAR)) + 2);
    if (pwsz)
    {
        memcpy(pwsz, pwszUnicodeString, (cblen * sizeof(WCHAR)));
        pwsz[cblen] = L'\0'; // Make sure it is NULL terminated.
    }

    return pwsz;
}

////////////////////////////////////////////////////////////////////////
// DsoCopyString
//
//  Duplicates the string into private heap string.
//
STDAPI_(LPWSTR) DsoCopyString(LPCWSTR pwszString)
{
	LPWSTR pwsz;
	UINT cblen;

	CHECK_NULL_RETURN(pwszString, NULL);
	cblen = lstrlenW(pwszString);

    pwsz = (LPWSTR)DsoMemAlloc((cblen * sizeof(WCHAR)) + 2);
    if (pwsz)
    {
        memcpy(pwsz, pwszString, (cblen * sizeof(WCHAR)));
        pwsz[cblen] = L'\0'; // Make sure it is NULL terminated.
    }

    return pwsz;
}

////////////////////////////////////////////////////////////////////////
// DsoCopyStringCat
//
STDAPI_(LPWSTR) DsoCopyStringCat(LPCWSTR pwszString1, LPCWSTR pwszString2)
{return DsoCopyStringCatEx(pwszString1, 1, &pwszString2);}

////////////////////////////////////////////////////////////////////////
// DsoCopyStringCatEx
//
//  Duplicates the string into private heap string and appends one or more
//  strings to the end (concatenation). 
//
STDAPI_(LPWSTR) DsoCopyStringCatEx(LPCWSTR pwszBaseString, UINT cStrs, LPCWSTR *ppwszStrs)
{
	LPWSTR pwsz;
	UINT i, cblenb, cblent;
    UINT *pcblens;

 // We assume you have a base string to start with. If not, we return NULL...
    if ((pwszBaseString == NULL) || 
        ((cblenb = lstrlenW(pwszBaseString)) < 1))
        return NULL;

 // If we have nothing to append, just do a plain copy...
    if ((cStrs == 0) || (ppwszStrs == NULL))
        return DsoCopyString(pwszBaseString);

 // Determine the size of the final string by finding the lengths
 // of each. We create an array of sizes to use later on...
    cblent = cblenb;
    pcblens = new UINT[cStrs];
    CHECK_NULL_RETURN(pcblens,  NULL);

    for (i = 0; i < cStrs; i++)
    {
        pcblens[i] =  lstrlenW(ppwszStrs[i]);
        cblent += pcblens[i];
    }

 // If we have data to append, create the new string and append the
 // data by copying them in place. We expect UTF-16 Unicode strings
 // for this to work, but this should be normal...
	if (cblent > cblenb)
    {
	    pwsz = (LPWSTR)DsoMemAlloc(((cblent + 1) * sizeof(WCHAR)));
	    CHECK_NULL_RETURN(pwsz,   NULL);

	    memcpy(pwsz, pwszBaseString, (cblenb * sizeof(WCHAR)));
        cblent = cblenb;

        for (i = 0; i < cStrs; i++)
        {
		    memcpy((pwsz + cblent), ppwszStrs[i], (pcblens[i] * sizeof(WCHAR)));
            cblent += pcblens[i];
        }
    }
    else pwsz = DsoCopyString(pwszBaseString);

    delete [] pcblens;
	return pwsz;
}

////////////////////////////////////////////////////////////////////////
// DsoCLSIDtoLPSTR
//
STDAPI_(LPWSTR) DsoCLSIDtoLPWSTR(REFCLSID clsid)
{
	LPWSTR psz = NULL;
	LPWSTR pwsz;
	if (SUCCEEDED(StringFromCLSID(clsid, &pwsz)))
	{
		psz = DsoCopyString(pwsz);
		CoTaskMemFree(pwsz);
	}
    return psz;
}


///////////////////////////////////////////////////////////////////////////////////
// DsoCompareStringsEx
//
//  Calls CompareString API using Unicode version (if available on OS). Otherwise,
//  we have to thunk strings down to MBCS to compare. This is fairly inefficient for
//  Win9x systems that don't handle Unicode, but hey...this is only a sample.
//
STDAPI_(UINT) DsoCompareStringsEx(LPCWSTR pwsz1, INT cch1, LPCWSTR pwsz2, INT cch2)
{
	UINT iret;
	LCID lcid = GetThreadLocale();
	UINT cblen1, cblen2;

 // Check that valid parameters are passed and then contain somethimg...
    if ((pwsz1 == NULL) || (cch1 == 0) || 
        ((cblen1 = ((cch1 > 0) ? cch1 : lstrlenW(pwsz1))) == 0))
		return CSTR_LESS_THAN;

	if ((pwsz2 == NULL) || (cch2 == 0) || 
        ((cblen2 = ((cch2 > 0) ? cch2 : lstrlenW(pwsz2))) == 0))
		return CSTR_GREATER_THAN;

 // If the string is of the same size, then we do quick compare to test for
 // equality (this is slightly faster than calling the API, but only if we
 // expect the calls to find an equal match)...
	if (cblen1 == cblen2)
	{
		for (iret = 0; iret < cblen1; iret++)
		{
			if (pwsz1[iret] == pwsz2[iret])
				continue;

			if (((pwsz1[iret] >= 'A') && (pwsz1[iret] <= 'Z')) &&
				((pwsz1[iret] + ('a' - 'A')) == pwsz2[iret]))
				continue;

			if (((pwsz2[iret] >= 'A') && (pwsz2[iret] <= 'Z')) &&
				((pwsz2[iret] + ('a' - 'A')) == pwsz1[iret]))
				continue;

			break; // don't continue if we can't quickly match...
		}

		// If we made it all the way, then they are equal...
		if (iret == cblen1)
			return CSTR_EQUAL;
	}

 // Now ask the OS to check the strings and give us its read. (We prefer checking
 // in Unicode since this is faster and we may have strings that can't be thunked
 // down to the local ANSI code page)...
	//if (v_fUnicodeAPI)
	//{
		iret = CompareStringW(lcid, NORM_IGNORECASE | NORM_IGNOREWIDTH, pwsz1, cblen1, pwsz2, cblen2);
	//}
	//else
	//{
	// // If we are on Win9x, we don't have much of choice (thunk the call)...
	//	LPSTR psz1 = DsoConvertToMBCS(pwsz1);
	//	LPSTR psz2 = DsoConvertToMBCS(pwsz2);
	//	iret = CompareStringA(lcid, NORM_IGNORECASE, psz1, -1, psz2, -1);
	//	DsoMemFree(psz2);
	//	DsoMemFree(psz1);
	//}

	return iret;
}


////////////////////////////////////////////////////////////////////////
// OLE Conversion Functions
//
#define HIMETRIC_PER_INCH   2540      // number HIMETRIC units per inch
#define PTS_PER_INCH        72        // number points (font size) per inch

#define MAP_PIX_TO_LOGHIM(x,ppli)   MulDiv(HIMETRIC_PER_INCH, (x), (ppli))
#define MAP_LOGHIM_TO_PIX(x,ppli)   MulDiv((ppli), (x), HIMETRIC_PER_INCH)

////////////////////////////////////////////////////////////////////////
// DsoHimetricToPixels
//
STDAPI_(void) DsoHimetricToPixels(LONG* px, LONG* py)
{
    HDC hdc = GetDC(NULL);
    if (px) *px = MAP_LOGHIM_TO_PIX(*px, GetDeviceCaps(hdc, LOGPIXELSX));
    if (py) *py = MAP_LOGHIM_TO_PIX(*py, GetDeviceCaps(hdc, LOGPIXELSY));
    ReleaseDC(NULL, hdc);
}

////////////////////////////////////////////////////////////////////////
// DsoPixelsToHimetric
//
STDAPI_(void) DsoPixelsToHimetric(LONG* px, LONG* py)
{
    HDC hdc = GetDC(NULL);
    if (px) *px = MAP_PIX_TO_LOGHIM(*px, GetDeviceCaps(hdc, LOGPIXELSX));
    if (py) *py = MAP_PIX_TO_LOGHIM(*py, GetDeviceCaps(hdc, LOGPIXELSY));
    ReleaseDC(NULL, hdc);
}

////////////////////////////////////////////////////////////////////////
// GDI Helper Functions
//
STDAPI_(HBITMAP) DsoGetBitmapFromWindow(HWND hwnd)
{
    HBITMAP hbmpold, hbmp = NULL;
    HDC hdcWin, hdcMem;
    RECT rc;
    INT x, y;

    if (!GetWindowRect(hwnd, &rc))
        return NULL;

    x = (rc.right - rc.left); if (x < 0) x = 1;
    y = (rc.bottom - rc.top); if (y < 0) y = 1;

	hdcWin = GetDC(hwnd);
	hdcMem = CreateCompatibleDC(hdcWin);

	hbmp = CreateCompatibleBitmap(hdcWin, x, y);
	hbmpold = (HBITMAP)SelectObject(hdcMem, hbmp);

	BitBlt(hdcMem, 0,0, x, y, hdcWin, 0,0, SRCCOPY);

	SelectObject(hdcMem, hbmpold);
	DeleteDC(hdcMem);
	ReleaseDC(hwnd, hdcWin);

    return hbmp;
}

////////////////////////////////////////////////////////////////////////
// Windows Helper Functions
//
STDAPI_(BOOL) IsWindowChild(HWND hwndParent, HWND hwndChild)
{
    HWND hwnd;

    if ((hwndChild == NULL) || !IsWindow(hwndChild))
        return FALSE;

    if (hwndParent == NULL)
        return TRUE;

    if (!IsWindow(hwndParent))
        return FALSE;

    hwnd = hwndChild;

    while (hwnd = GetParent(hwnd))
	    if (hwnd == hwndParent)
            return TRUE;

    return FALSE;
}

////////////////////////////////////////////////////////////////////////
// DsoDispatchInvoke
//
//  Wrapper for IDispatch::Invoke calls to event sinks, or late bound call
//  to embedded object to get ambient property.
//
STDAPI DsoDispatchInvoke(LPDISPATCH pdisp, LPOLESTR pwszname, DISPID dspid, WORD wflags, DWORD cargs, VARIANT* rgargs, VARIANT* pvtret)
{
    HRESULT    hr = S_FALSE;
    DISPID     dspidPut = DISPID_PROPERTYPUT;
    DISPPARAMS dspparm = {NULL, NULL, 0, 0};

	CHECK_NULL_RETURN(pdisp, E_POINTER);

    dspparm.rgvarg = rgargs;
    dspparm.cArgs = cargs;

	if ((wflags & DISPATCH_PROPERTYPUT) || (wflags & DISPATCH_PROPERTYPUTREF))
	{
		dspparm.rgdispidNamedArgs = &dspidPut;
		dspparm.cNamedArgs = 1;
	}

	SEH_TRY

	if (pwszname)
		hr = pdisp->GetIDsOfNames(IID_NULL, &pwszname, 1, LOCALE_USER_DEFAULT, &dspid);

    if (SUCCEEDED(hr))
        hr = pdisp->Invoke(dspid, IID_NULL, LOCALE_USER_DEFAULT, (WORD)(DISPATCH_METHOD | wflags), &dspparm, pvtret, NULL, NULL);

    SEH_EXCEPT(hr)

    return hr;
}



HRESULT OLEMethod(int nType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
{
	if (!pDisp) return E_FAIL;

	va_list marker;
	va_start(marker, cArgs);

	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	char szName[200];


	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		return hr;
	}
	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	// Extract arguments...
	for (int i = 0; i < cArgs; i++) {
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts!
	if (nType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call!
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, nType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr)) {
		return hr;
	}
	// End variable-argument section...
	va_end(marker);

	delete[] pArgs;
	return hr;
}

////////////////////////////////////////////////////////////////////////
// DsoReportError -- Report Error for both ComThreadError and DispError.
//
STDAPI DsoReportError(HRESULT hr, LPWSTR pwszCustomMessage, EXCEPINFO* peiDispEx)
{
    BSTR bstrSource, bstrDescription;
    ICreateErrorInfo* pcerrinfo;
    IErrorInfo* perrinfo;
    WCHAR szError[MAX_PATH];
    UINT nID = 0;

 // Don't need to do anything unless this is an error.
    if ((hr == S_OK) || SUCCEEDED(hr)) return hr;

 // Is this one of our custom errors (if so we will pull description from resource)...
    if ((hr > DSO_E_ERR_BASE) && (hr < DSO_E_ERR_MAX))
        nID = (hr & 0xFF);

 // Set the source name...
    bstrSource = SysAllocString(L"DsoFramerControl");

 // Set the error description...
    if (pwszCustomMessage)
    {
        bstrDescription = SysAllocString(pwszCustomMessage);
    }
    else if (((nID) && LoadString(v_hModule, nID, szError, sizeof(szError))) || 
             (FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, szError, sizeof(szError), NULL)))
    {
        bstrDescription = szError;
    }
    else bstrDescription = NULL;
    
 // Set ErrorInfo so that vtable clients can get rich error information...
	if (SUCCEEDED(CreateErrorInfo(&pcerrinfo)))
    {
		pcerrinfo->SetSource(bstrSource);
        pcerrinfo->SetDescription(bstrDescription);

		if (SUCCEEDED(pcerrinfo->QueryInterface(IID_IErrorInfo, (void**) &perrinfo)))
        {
			SetErrorInfo(0, perrinfo);
			perrinfo->Release();
		}
		pcerrinfo->Release();
	}

 // Fill-in DispException Structure for late-boud clients...
    if (peiDispEx)
    {
        peiDispEx->scode = hr;
        peiDispEx->bstrSource = SysAllocString(bstrSource);
        peiDispEx->bstrDescription = SysAllocString(bstrDescription);
    }

 // Free temp strings...
    SAFE_FREEBSTR(bstrDescription);
    SAFE_FREEBSTR(bstrSource);

 // We always return error passed (so caller can chain this in return call).
    return hr;
}

////////////////////////////////////////////////////////////////////////
// Variant Type Helpers (Fast Access to Variant Data)
//
VARIANT* __fastcall DsoPVarFromPVarRef(VARIANT* px)
{return ((px->vt == (VT_VARIANT|VT_BYREF)) ? (px->pvarVal) : px);}

BOOL __fastcall DsoIsVarParamMissing(VARIANT* px)
{return ((px->vt == VT_EMPTY) || ((px->vt & VT_ERROR) == VT_ERROR));}

LPWSTR __fastcall DsoPVarWStrFromPVar(VARIANT* px)
{return ((px->vt == VT_BSTR) ? px->bstrVal : ((px->vt == (VT_BSTR|VT_BYREF)) ? *(px->pbstrVal) : NULL));}

SAFEARRAY* __fastcall DsoPVarArrayFromPVar(VARIANT* px)
{return (((px->vt & (VT_BYREF|VT_ARRAY)) == (VT_BYREF|VT_ARRAY)) ? *(px->pparray) : (((px->vt & VT_ARRAY) == VT_ARRAY) ? px->parray : NULL));}

IUnknown* __fastcall DsoPVarUnkFromPVar(VARIANT* px)
{return (((px->vt == VT_DISPATCH) || (px->vt == VT_UNKNOWN)) ? px->punkVal : (((px->vt == (VT_DISPATCH|VT_BYREF)) || (px->vt == (VT_UNKNOWN|VT_BYREF))) ? *(px->ppunkVal) : NULL));}

SHORT __fastcall DsoPVarShortFromPVar(VARIANT* px, SHORT fdef)
{return (((px->vt & 0xFF) != VT_I2) ? fdef : ((px->vt & VT_BYREF) == VT_BYREF) ? *(px->piVal) : px->iVal);}

LONG __fastcall DsoPVarLongFromPVar(VARIANT* px, LONG fdef)
{return (((px->vt & 0xFF) != VT_I4) ? (LONG)DsoPVarShortFromPVar(px, (SHORT)fdef) : ((px->vt & VT_BYREF) == VT_BYREF) ? *(px->plVal) : px->lVal);}

BOOL __fastcall DsoPVarBoolFromPVar(VARIANT* px, BOOL fdef)
{return (((px->vt & 0xFF) != VT_BOOL) ? (BOOL)DsoPVarLongFromPVar(px, (LONG)fdef) : ((px->vt & VT_BYREF) == VT_BYREF) ? (BOOL)(*(px->pboolVal)) : (BOOL)(px->boolVal));}


////////////////////////////////////////////////////////////////////////
// Win32 Unicode API wrappers
//
//  This project is not compiled to Unicode in order for it to run on Win9x
//  machines. However, in order to try to keep the code language/locale neutral 
//  we use these wrappers to call the Unicode API functions when supported,
//  and thunk down strings to local code page if must run MBCS API.
//

////////////////////////////////////////////////////////////////////////
// FFileExists
//
//  Returns TRUE if the given file exists. Does not handle URLs.
//
STDAPI_(BOOL) FFileExists(WCHAR* wzPath)
{
    DWORD dw = 0xFFFFFFFF;
    //if (v_fUnicodeAPI)
    //{
        dw = GetFileAttributesW(wzPath);
 //   }
 //   else
 //   {
	//	LPSTR psz = DsoConvertToMBCS(wzPath);
 //       if (psz) dw = GetFileAttributesA(psz);
	//	DsoMemFree(psz);
	//}
    return (dw != 0xFFFFFFFF);
}

////////////////////////////////////////////////////////////////////////
// FOpenLocalFile
//
//  Returns TRUE if the file can be opened with the access required.
//  Use the handle for ReadFile/WriteFile operations as normal.
//
STDAPI_(BOOL) FOpenLocalFile(WCHAR* wzFilePath, DWORD dwAccess, DWORD dwShareMode, DWORD dwCreate, HANDLE* phFile)
{
    CHECK_NULL_RETURN(phFile, FALSE);
    *phFile = INVALID_HANDLE_VALUE;
    //if (v_fUnicodeAPI)
    //{
	    *phFile = CreateFileW(wzFilePath, dwAccess, dwShareMode, NULL, dwCreate, FILE_ATTRIBUTE_NORMAL, NULL);
    //}
    //else
    //{
    //    LPSTR psz = DsoConvertToMBCS(wzFilePath);
    //    if (psz) *phFile = CreateFileA(psz, dwAccess, dwShareMode, NULL, dwCreate, FILE_ATTRIBUTE_NORMAL, NULL);
    //    DsoMemFree(psz);
    //}
    return (*phFile != INVALID_HANDLE_VALUE);
}


////////////////////////////////////////////////////////////////////////
// FPerformShellOp
//
//  This function started as a wrapper for SHFileOperation, but that 
//  shell function had an enormous performance hit, especially on Win9x
//  and NT4. To speed things up we removed the shell32 call and are
//  using the kernel32 APIs instead. The only drawback is we can't
//  handle multiple files, but that is not critical for this project.
//
STDAPI_(BOOL) FPerformShellOp(DWORD dwOp, WCHAR* wzFrom, WCHAR* wzTo)
{
	BOOL f = FALSE;
    //if (v_fUnicodeAPI)
    //{
		switch (dwOp)
		{
		case FO_COPY:		f = CopyFileW(wzFrom, wzTo, FALSE);	break;
		case FO_MOVE: 
		case FO_RENAME:		f = MoveFileW(wzFrom, wzTo);		break;
		case FO_DELETE:		f = DeleteFileW(wzFrom);			break;
		}
	//}
 //   else
 //   {
	//    LPSTR pszFrom = DsoConvertToMBCS(wzFrom);
	//    LPSTR pszTo = DsoConvertToMBCS(wzTo);

	//	switch (dwOp)
	//	{
	//	case FO_COPY:		f = CopyFileA(pszFrom, pszTo, FALSE); break;
	//	case FO_MOVE:
	//	case FO_RENAME:		f = MoveFileA(pszFrom, pszTo);		break;
	//	case FO_DELETE:		f = DeleteFileA(pszFrom);			break;
	//	}

	//    if (pszFrom) DsoMemFree(pszFrom);
	//    if (pszTo) DsoMemFree(pszTo);
 //   }

	return f;
}


////////////////////////////////////////////////////////////////////////
// FDrawText
//
//  This is used by the control for drawing the caption in the titlebar.
//  Since a custom caption could contain Unicode characters only printable
//  on Unicode OS, we should try to use the Unicode version when available.
//
STDAPI_(BOOL) FDrawText(HDC hdc, WCHAR* pwsz, LPRECT prc, UINT fmt)
{
	BOOL f;
    //if (v_fUnicodeAPI)
    //{
        f = (BOOL)DrawTextW(hdc, pwsz, -1, prc, fmt);
  //  }
  //  else
  //  {
		//LPSTR psz = DsoConvertToMBCS(pwsz);
		//f = (BOOL)DrawTextA(hdc, psz, -1, prc, fmt);
		//DsoMemFree(psz);
  //  }
	return f;
}


////////////////////////////////////////////////////////////////////////
// FSetRegKeyValue
//
//  We use this for registration when dealing with the file path, since
//  that path may have Unicode characters on some systems. Win9x boxes
//  will have to be converted to ANSI.
//
STDAPI_(BOOL) FSetRegKeyValue(HKEY hk, WCHAR* pwsz)
{
	LONG lret;
    lret = RegSetValueExW(hk, NULL, 0, REG_SZ, (BYTE*)pwsz, (lstrlenW(pwsz) * sizeof(WCHAR)));

	return (lret == ERROR_SUCCESS);
}
