' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTaskDialog.inc
' Purpose:  Windows common controls - Task dialog
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#if not defined(UNICODE)
   #define UNICODE
#endif
#if not defined(_WIN32_WINNT)
   #define _WIN32_WINNT &h0602
#endif

#pragma once
#include once "windows.bi"
#include once "win/commctrl.bi"
#include once "Afx2/CControl.inc"
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CTaskDialog class
' ########################################################################################

TYPE CTaskDialog EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE SUB ClickButton (BYVAL nButton AS DWORD)
   DECLARE SUB ClickRadioButton (BYVAL nRadioButton AS DWORD)
   DECLARE SUB ClickVerification (BYVAL fChecked AS BOOLEAN, BYVAL fFocused AS BOOLEAN)
   DECLARE SUB EnableButton (BYVAL nButton AS LONG)
   DECLARE SUB DisableButton (BYVAL nButton AS LONG)
   DECLARE SUB EnableRadioButton (BYVAL nRadioButton AS LONG)
   DECLARE SUB DisableRadioButton (BYVAL nRadioButton AS LONG)
   DECLARE SUB NavigatePage (BYREF TaskConfig AS TASKDIALOGCONFIG)
   DECLARE SUB SetButtonElevationRequiredState (BYVAL nButton AS LONG, BYVAL fElevation AS BOOLEAN)
   DECLARE SUB SetElementText (BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   DECLARE SUB SetMarqueeProgressBar (BYVAL fMarquee AS BOOLEAN)
   DECLARE SUB SetProgressBarMarquee (BYVAL fMarquee AS LONG, BYVAL nSpeed AS LONG)
   DECLARE FUNCTION SetProgressBarPos (BYVAL nNewPos AS LONG) AS LONG
   DECLARE FUNCTION SetProgressBarRange (BYVAL nMinRange AS SHORT, BYVAL nMaxRange AS SHORT) AS DWORD
   DECLARE FUNCTION SetProgressBarState (BYVAL nNewState AS LONG) AS LONG
   DECLARE SUB UpdateElementText (BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   DECLARE SUB UpdateIcon (BYVAL nIconElement AS LONG, BYVAL idResource AS DWORD)

END TYPE
' ========================================================================================

' ========================================================================================
' * CTaskDialog constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTaskDialog (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTaskDialog (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTaskDialog (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTaskDialog (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTaskDialog destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTaskDialog
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                   *** TASK DIALOG ***
' ########################################################################################

' ========================================================================================
' Simulates the action of a button click in the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.ClickButton (BYVAL nButton AS DWORD)
   SendMessageW(m_hCtl, TDM_CLICK_BUTTON, nButton, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Simulates the action of a radio button click in the TaskDialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.ClickRadioButton (BYVAL nRadioButton AS DWORD)
   SendMessageW(m_hCtl, TDM_CLICK_RADIO_BUTTON, nRadioButton, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Simulates a click on the verification checkbox of the Task Dialog, if it exists.
' ========================================================================================
PRIVATE SUB CTaskDialog.ClickVerification (BYVAL fChecked AS BOOLEAN, BYVAL fFocused AS BOOLEAN)
   SendMessageW(m_hCtl, TDM_CLICK_VERIFICATION, fChecked, fFocused)
END SUB
' ========================================================================================

' ========================================================================================
' Enables or disables a push button in the TaskDialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.EnableButton (BYVAL nButton AS LONG)
   SendMessageW(m_hCtl, TDM_ENABLE_BUTTON, nButton, CTRUE)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CTaskDialog.DisableButton (BYVAL nButton AS LONG)
   SendMessageW(m_hCtl, TDM_ENABLE_BUTTON, nButton, FALsE)
END SUB
' ========================================================================================

' ========================================================================================
' Enables or disables a radio button in the TaskDialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.EnableRadioButton (BYVAL nRadioButton AS LONG)
   SendMessageW(m_hCtl, TDM_ENABLE_RADIO_BUTTON, nRadioButton, CTRUE)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CTaskDialog.DisableRadioButton (BYVAL nRadioButton AS LONG)
   SendMessageW(m_hCtl, TDM_ENABLE_RADIO_BUTTON, nRadioButton, FALSE)
END SUB
' ========================================================================================

' ========================================================================================
' Refreshes the contents of the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.NavigatePage (BYREF TaskConfig AS TASKDIALOGCONFIG)
   IF VARPTR(TaskConfig) = 0 THEN EXIT SUB
   TaskConfig.cbSize = SIZEOF(TaskConfig)
   SendMessageW(m_hCtl, TDM_NAVIGATE_PAGE, 0, CAST(LPARAM, @TaskConfig))
END SUB
' ========================================================================================

' ========================================================================================
' Used by clients to designate whether a given Task Dialog button or command link should
' have a User Account Control (UAC) shield icon (in other words, whether the action
' invoked by the button requires elevation).
' ========================================================================================
PRIVATE SUB CTaskDialog.SetButtonElevationRequiredState (BYVAL nButton AS LONG, BYVAL fElevation AS BOOLEAN)
   SendMessageW(m_hCtl, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, nButton, fElevation)
END SUB
' ========================================================================================

' ========================================================================================
' Used by clients to update a text element on the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.SetElementText (BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   SendMessageW(m_hCtl, TDM_SET_ELEMENT_TEXT, eElement, CAST(LPARAM, pszNewText))
END SUB
' ========================================================================================

' ========================================================================================
' Indicates whether the hosted progress bar should be displayed in marquee mode.
' ========================================================================================
PRIVATE SUB CTaskDialog.SetMarqueeProgressBar (BYVAL fMarquee AS BOOLEAN)
   SendMessageW(m_hCtl, TDM_SET_MARQUEE_PROGRESS_BAR, cast(WPARAM, fMarquee), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Starts and stops the marquee display of the progress bar, and sets the speed of the marquee.
' ========================================================================================
PRIVATE SUB CTaskDialog.SetProgressBarMarquee (BYVAL fMarquee AS LONG, BYVAL nSpeed AS LONG)
   SendMessageW(m_hCtl, TDM_SET_PROGRESS_BAR_MARQUEE, fMarquee, nSpeed)
END SUB
' ========================================================================================

' ========================================================================================
' The TDM_SET_PROGRESS_BAR_POS message is used set the current position for a progress bar.
' Returns the previous position.
' ========================================================================================
PRIVATE FUNCTION CTaskDialog.SetProgressBarPos (BYVAL nNewPos AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TDM_SET_PROGRESS_BAR_POS, nNewPos, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The TDM_SET_PROGRESS_BAR_RANGE message is used set the minimum and maximum values for
' the hosted progress bar. Returns the previous minimum and maximum values, if successful,
' or zero otherwise. The LOWORD contains the minimum value, and the HIWORD contains the
' maximum value.
' ========================================================================================
PRIVATE FUNCTION CTaskDialog.SetProgressBarRange (BYVAL nMinRange AS SHORT, BYVAL nMaxRange AS SHORT) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TDM_SET_PROGRESS_BAR_RANGE, 0, MAKELONG(nMinRange, nMaxRange))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The TDM_SET_PROGRESS_BAR_POS message is used set the current position for a progress bar.
' If the function succeeds, the return value is non zero. If the function fails, the return
' value is zero. To get extended error information call GetLastError.
' ========================================================================================
PRIVATE FUNCTION CTaskDialog.SetProgressBarState (BYVAL nNewState AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TDM_SET_PROGRESS_BAR_STATE, nNewState, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The TDM_UPDATE_ELEMENT_TEXT message is available for use by clients to update a text
' element on the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.UpdateElementText (BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   SendMessageW(m_hCtl, TDM_UPDATE_ELEMENT_TEXT, eElement, CAST(LPARAM, pszNewText))
END SUB
' ========================================================================================

' ========================================================================================
' Refreshes the icon of a task dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.UpdateIcon (BYVAL nIconElement AS LONG, BYVAL idResource AS DWORD)
   SendMessageW(m_hCtl, TDM_UPDATE_ICON, nIconElement, idResource)
END SUB
' ========================================================================================

END NAMESPACE
