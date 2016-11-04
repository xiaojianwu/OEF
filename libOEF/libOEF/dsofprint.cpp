/***************************************************************************
 * DSOFPRINT.CPP
 *
 * CDsoDocObject: Print Code for CDsoDocObject object
 ***************************************************************************/
#include "dsoframer.h"

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::PrintDocument
//
//  We can use the IPrint interface for an ActiveX Document object to 
//  selectively print object to a given printer. 
//
STDMETHODIMP CDsoDocObject::PrintDocument(LPCWSTR pwszPrinter, LPCWSTR pwszOutput, UINT cCopies, UINT nFrom, UINT nTo, BOOL fPromptUser)
{
    HRESULT hr = S_OK;
 //   IPrint *print;
 //   HANDLE hPrint;
 //   DVTARGETDEVICE* ptd = NULL;

 //   ODS("CDsoDocObject::PrintDocument\n");
 //   CHECK_NULL_RETURN(m_pole, E_UNEXPECTED);

 //// First thing we need to do is ask object for IPrint. If it does not
 //// support it, we cannot continue. It is up to DocObj if this is allowed...
 //   hr = m_pole->QueryInterface(IID_IPrint, (void**)&print);
 //   RETURN_ON_FAILURE(hr);

 //// Now setup printer settings into DEVMODE for IPrint. Open printer 
 //// settings and gather default DEVMODE...
 //   if (FOpenPrinter(pwszPrinter, &hPrint))
 //   {
 //       LPDEVMODEW pdevMode     = NULL;
 //       LPWSTR pwszDefProcessor = NULL;
 //       LPWSTR pwszDefDriver    = NULL;
 //       LPWSTR pwszDefPort      = NULL;
 //       LPWSTR pwszPort;
 //       DWORD  cbDevModeSize;

 //       if (FGetPrinterSettings(hPrint, &pwszDefProcessor, 
 //               &pwszDefDriver, &pwszDefPort, &pdevMode, &cbDevModeSize) && (pdevMode))
 //       {
 //           DWORD cbPrintName, cbDeviceName, cbOutputName;
 //           DWORD cbDVTargetSize;

 //           pdevMode->dmFields |= DM_COPIES;
 //           pdevMode->dmCopies = (WORD)((cCopies) ? cCopies : 1);

 //           pwszPort = ((pwszOutput) ? (LPWSTR)pwszOutput : pwszDefPort);
 //        
 //        // We calculate the size we will need for the TARGETDEVICE structure...
 //           cbPrintName  = ((lstrlenW(pwszDefProcessor) + 1) * sizeof(WCHAR));
 //           cbDeviceName = ((lstrlenW(pwszDefDriver)    + 1) * sizeof(WCHAR));
 //           cbOutputName = ((lstrlenW(pwszPort)         + 1) * sizeof(WCHAR));

 //           cbDVTargetSize = sizeof(DWORD) + sizeof(DEVNAMES) + cbPrintName +
 //                           cbDeviceName + cbOutputName + cbDevModeSize;

 //        // Allocate new target device using COM Task Allocator...
 //           ptd = (DVTARGETDEVICE*)CoTaskMemAlloc(cbDVTargetSize);
 //           if (ptd)
 //           {
 //            // Copy all the data in the DVT...
 //               DWORD dwOffset = sizeof(DWORD) + sizeof(DEVNAMES);
 //               ptd->tdSize = cbDVTargetSize;

 //               ptd->tdDriverNameOffset = (WORD)dwOffset;
 //               memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pwszDefProcessor, cbPrintName);
 //               dwOffset += cbPrintName;

 //               ptd->tdDeviceNameOffset = (WORD)dwOffset;
 //               memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pwszDefDriver, cbDeviceName);
 //               dwOffset += cbDeviceName;

 //               ptd->tdPortNameOffset = (WORD)dwOffset;
 //               memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pwszPort, cbOutputName);
 //               dwOffset += cbOutputName;

 //               ptd->tdExtDevmodeOffset = (WORD)dwOffset;
 //               memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pdevMode, cbDevModeSize);
 //               dwOffset += cbDevModeSize;

 //               ASSERT(dwOffset == cbDVTargetSize);
 //           }
 //           else hr = E_OUTOFMEMORY;

 //        // We're done with the devmode...
 //           DsoMemFree(pdevMode);
 //       }
 //       else hr = E_WIN32_LASTERROR;

 //       SAFE_FREESTRING(pwszDefPort);
 //       SAFE_FREESTRING(pwszDefDriver);
 //       SAFE_FREESTRING(pwszDefProcessor);
 //       ClosePrinter(hPrint);
 //   }
 //   else hr = E_WIN32_LASTERROR;

 //// If we were successful in getting TARGETDEVICE struct, provide the page range
 //// for the print job and ask docobj server to print it...
 //   if (SUCCEEDED(hr))
 //   {
 //       PAGESET *ppgset;
 //       DWORD cbPgRngSize = sizeof(PAGESET) + sizeof(PAGERANGE);
 //       LONG cPages, cLastPage;
 //       DWORD grfPrintFlags;

 //    // Setup the page range to print...
 //       if ((ppgset = (PAGESET*)CoTaskMemAlloc(cbPgRngSize)) != NULL)
 //       {
 //           ppgset->cbStruct = cbPgRngSize;
 //           ppgset->cPageRange   = 1;
 //           ppgset->fOddPages    = TRUE;
 //           ppgset->fEvenPages   = TRUE;
 //           ppgset->cPageRange   = 1;
 //           ppgset->rgPages[0].nFromPage = ((nFrom) ? nFrom : 1);
 //           ppgset->rgPages[0].nToPage   = ((nTo) ? nTo : PAGESET_TOLASTPAGE);

 //        // Give indication we are waiting (on the print)...
 //           HCURSOR hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));

 //           SEH_TRY

 //       // Setup the initial page number (optional)...
 //           print->SetInitialPageNum(ppgset->rgPages[0].nFromPage);

 //           grfPrintFlags = (PRINTFLAG_MAYBOTHERUSER | PRINTFLAG_RECOMPOSETODEVICE);
 //           if (fPromptUser) grfPrintFlags |= PRINTFLAG_PROMPTUSER;
 //           if (pwszOutput)  grfPrintFlags |= PRINTFLAG_PRINTTOFILE;

 //        // Now ask server to print it using settings passed...
 //           hr = print->Print(grfPrintFlags, &ptd, &ppgset, NULL, (IContinueCallback*)&m_xContinueCallback, 
 //                   ppgset->rgPages[0].nFromPage, &cPages, &cLastPage);

 //           SEH_EXCEPT(hr)

 //           SetCursor(hCur);

 //           if (ppgset)
 //               CoTaskMemFree(ppgset);
 //       }
 //       else hr = E_OUTOFMEMORY;
 //   }

 //// We are done...
 //   if (ptd) CoTaskMemFree(ptd);
 //   print->Release();
    return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::StartPrintPreview
//
//  Ask embedded object to go into print preview (if supportted).
//
//STDMETHODIMP CDsoDocObject::StartPrintPreview()
//{
//    HRESULT hr;
//    IOleInplacePrintPreview *prev;
//    HCURSOR hCur;
//
//    ODS("CDsoDocObject::StartPrintPreview\n");
//    CHECK_NULL_RETURN(m_pole, E_UNEXPECTED);
//
// // No need to do anything if already in preview...
//    if (InPrintPreview()) return S_FALSE;
//
// // Otherwise, ask document server if it supports preview...
//    hr = m_pole->QueryInterface(IID_IOleInplacePrintPreview, (void**)&prev);
//    if (SUCCEEDED(hr))
//    {
//     // Tell user we waiting (switch to preview can be slow for very large docs)...
//        hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));
//
//     // If it does, make sure it can go into preview mode...
//        hr = prev->QueryStatus();
//        if (SUCCEEDED(hr))
//        {
//            SEH_TRY
//
//            if (m_hwndCtl) // Notify the control that preview started...
//                SendMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)FALSE);
//
//         // We will allow application to bother user and switch printers...
//            hr = prev->StartPrintPreview(
//                (PREVIEWFLAG_MAYBOTHERUSER | PREVIEWFLAG_PROMPTUSER | PREVIEWFLAG_USERMAYCHANGEPRINTER),
//                NULL, (IOlePreviewCallback*)&m_xPreviewCallback, 1);
//
//            SEH_EXCEPT(hr)
//
//         // If the call succeeds, we keep hold of interface to close preview later
//            if (SUCCEEDED(hr)) 
//			{ 
//				SAFE_SET_INTERFACE(m_pprtprv, prev);
//			}
//			else
//			{ // Otherwise, notify the control that preview failed...
//				if (m_hwndCtl) 
//					PostMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)TRUE);
//			}
//        }
//
//        SetCursor(hCur);
//        prev->Release();
//    }
//    else if (IsPPTObject() && (m_hwndUIActiveObj))
//    {
//     // PowerPoint doesn't have print preview, but it does have slide show mode, so we can use this to
//     // toggle the viewing into slideshow...
//        if (PostMessage(m_hwndUIActiveObj, WM_KEYDOWN, VK_F5, 0x00000001) &&
//            PostMessage(m_hwndUIActiveObj, WM_KEYUP,   VK_F5, 0xC0000001))
//        {
//            hr = S_OK; m_fAttemptPptPreview = TRUE;
//        }
//    }
//
//    return hr;
//}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::ExitPrintPreview
//
//  Drop out of print preview and restore items as needed.
//
STDMETHODIMP CDsoDocObject::ExitPrintPreview(BOOL fForceExit)
{
    TRACE1("CDsoDocObject::ExitPrintPreview(fForceExit=%d)\n", (DWORD)fForceExit);

 // Need to be in preview to run this function...
    if (!InPrintPreview()) return S_FALSE;

 // If the user closes the app or otherwise terminates the preview, we need
 // to notify the ActiveDocument server to leave preview mode...
    if (m_pprtprv)
    {
        if (fForceExit) // Tell docobj we want to end preview...
        {
            HRESULT hr = m_pprtprv->EndPrintPreview(TRUE);
            ASSERT(SUCCEEDED(hr)); (void)hr;
        }
    }
    else if (m_fInPptSlideShow)
    {
        if ((fForceExit) && (m_hwndUIActiveObj))
        {
         // HACK: When it goes into slideshow, PPT 2007 changes the active window but 
         // doesn't call SetActiveObject to update us with new object and window handle.
         // If we post VK_ESCAPE to the window handle we have, it can fail. It needs to go
         // to the slideshow window that PPT failed to give us handle for. As workaround,
         // setting focus to the UI window we have handle for will automatically forward
         // to the right window, so we can use that trick to get the right window and 
         // make the call in a way that that should succeed regardless of PPT version...
            SetFocus(m_hwndUIActiveObj);
            PostMessage(GetFocus(), WM_KEYDOWN, VK_ESCAPE, 0x00000001);
            PostMessage(GetFocus(), WM_KEYUP,   VK_ESCAPE, 0xC0000001);
        }
        m_fAttemptPptPreview = FALSE;
        m_fInPptSlideShow = FALSE;
    }

    if (m_hwndCtl) // Notify the control that preview ended...
        SendMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)TRUE);

 // Free our reference to preview interface...
    SAFE_RELEASE_INTERFACE(m_pprtprv);
    return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CheckForPPTPreviewChange
//
//  Used to update control when PPT goes in/out of slideshow asynchronously.
//
void CDsoDocObject::CheckForPPTPreviewChange()
{
    if (m_fInPptSlideShow)
    {
        ExitPrintPreview(FALSE);
    }
    else if (m_fAttemptPptPreview)
    {
        m_fInPptSlideShow = TRUE;
        m_fAttemptPptPreview = FALSE;
        if (m_hwndCtl) // Notify the control that preview started...
            SendMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)FALSE);
    }
}