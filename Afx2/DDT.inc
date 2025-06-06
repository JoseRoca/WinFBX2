' ########################################################################################
' Platform: Microsoft Windows
' Filename: DDT.bas
' Purpose:  Wrapper module to use DDT-like syntax with CDialog.
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

#pragma once
#include once "Afx2/CDialog.inc"
#include once "Afx2/AfxWin2.inc"
#include once "Afx2/AfxCtl2.inc"
using Afx2

' ========================================================================================
' Macro for debug
' To allow debugging, define _DDT_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _DDT_DEBUG_
   #define _DDT_DEBUG_ 0
#ENDIF
#ifndef _DDT_DP_
   #define _DDT_DP_ 1
   #MACRO DDT_DP(st)
      #IF (_DDT_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE DDT

' ========================================================================================
' Returns a pointer to the CDialog class given the handle of the window created with it
' or the handle of any of it's children.
' ========================================================================================
PRIVATE FUNCTION GetCDialogPtr (BYVAL hwnd AS HWND) AS CDialog PTR
   IF hwnd = NULL THEN RETURN NULL
   DIM hRootOwner AS ..HWND = GetAncestor(hwnd, GA_ROOTOWNER)
   IF hRootOwner = NULL THEN RETURN NULL
   RETURN CAST(CDialog PTR, GetWindowLongPtrW(hRootOwner, DWLP_USER))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a pointer to the CDialog class given the handle of the window created with it
' or the handle of any of it's children.
' ========================================================================================
PRIVATE FUNCTION IsWin OVERLOAD (BYVAL hDlg AS HWND) AS BOOLEAN
   IF hDlg = NULL THEN RETURN FALSE
   RETURN IsWindow(hDlg)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION IsWin OVERLOAD (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS BOOLEAN
   IF hDlg = NULL THEN RETURN FALSE
   RETURN IsWindow(GetDlgItem(hDlg, id))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a dialog box using a memory template
' ========================================================================================
PRIVATE FUNCTION DialogNew OVERLOAD (BYVAL hParent AS HWND = HWND_DESKTOP, BYREF wszTitle AS WSTRING, _
BYVAL x AS LONG = -1, BYVAL y AS LONG = -1, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0, BYREF fontName AS WSTRING = "Segoe UI", _
BYVAL ptSize AS LONG = 9, BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) AS HWND
   DDT_DP(" - UNITS - Class #32770")
   ' // Creates a new instance of CDialog
   DIM pDlg AS CDIALOG PTR = NEW CDialog(fontName, ptSize, fontStyle, charset)
   IF pDlg THEN RETURN pDlg->DialogNew(hParent, wszTitle, x, y, cx, cy, dwStyle, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a custom dialog box using a memory template
' ========================================================================================
PRIVATE FUNCTION DialogNew OVERLOAD (BYREF wszClassName AS WSTRING, BYVAL hParent AS HWND = HWND_DESKTOP, _
BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = -1, BYVAL y AS LONG = -1, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0, BYREF fontName AS WSTRING = "Segoe UI", _
BYVAL ptSize AS LONG = 9, BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) AS HWND
   DDT_DP(" - UNITS - Class " & wszClassName)
   ' // Creates a new instance of CDialog
   DIM pDlg AS CDIALOG PTR = NEW CDialog(fontName, ptSize, fontStyle, charset)
   IF pDlg THEN RETURN pDlg->DialogNew(wszClassName, hParent, wszTitle, x, y, cx, cy, dwStyle, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a custom dialog box using a memory template
' ========================================================================================
PRIVATE FUNCTION DialogNewPixels OVERLOAD (BYVAL hParent AS HWND = HWND_DESKTOP, BYREF wszTitle AS WSTRING, _
BYVAL x AS LONG = -1, BYVAL y AS LONG = -1, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0, BYREF fontName AS WSTRING = "Segoe UI", _
BYVAL ptSize AS LONG = 9, BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) AS HWND
   DDT_DP(" - PIXELS - Class #32770")
   ' // Creates a new instance of CDialog
   DIM pDlg AS CDIALOG PTR = NEW CDialog(fontName, ptSize, fontStyle, charset)
   IF pDlg THEN RETURN pDlg->DialogNewPixels(hParent, wszTitle, x, y, cx, cy, dwStyle, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a custom dialog box using a memory template
' ========================================================================================
PRIVATE FUNCTION DialogNewPixels OVERLOAD (BYREF wszClassName AS WSTRING, BYVAL hParent AS HWND = HWND_DESKTOP, _
BYREF wszTitle AS WSTRING, BYVAL x AS LONG = -1, BYVAL y AS LONG = -1, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
BYVAL dwStyle AS LONG = 0, BYVAL dwExStyle AS LONG = 0, BYREF fontName AS WSTRING = "Segoe UI", _
BYVAL ptSize AS LONG = 9, BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) AS HWND
   DDT_DP(" - PIXELS - Class " & wszClassName)
   ' // Creates a new instance of CDialog
   DIM pDlg AS CDIALOG PTR = NEW CDialog(fontName, ptSize, fontStyle, charset)
   IF pDlg THEN RETURN pDlg->DialogNewPixels(wszClassName, hParent, wszTitle, x, y, cx, cy, dwStyle, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows the dialog as modal.
' ========================================================================================
PRIVATE FUNCTION DialogShowModal (BYVAL hDlg AS HWND, BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogShowModal(pCallback)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows the dialog as modeless.
' ========================================================================================
PRIVATE FUNCTION DialogShowModeless (BYVAL hDlg AS HWND, BYVAL pCallback AS ANY PTR = NULL) AS INT_PTR
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogShowModeless(pCallback)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Message pump for modeless dialogs
' ========================================================================================
PRIVATE FUNCTION DialogDoEvents (BYVAL dwMilliseconds AS DWORD = 1) AS BOOLEAN
   DIM uMsg AS tagMsg
   IF PeekMessageW(@uMsg, NULL, 0, 0, PM_REMOVE) THEN
      IF IsDialogMessageW(GetActiveWindow, @uMsg) = 0 THEN
         TranslateMessage @uMsg
         DispatchMessage @uMsg
      END IF
      RETURN TRUE
   ELSE
      Sleep_(dwMilliseconds)
      RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Centers a window on the screen or over another window.
' It also ensures that the placement is done within the work area.
' Parameters:
' - hwnd = Handle of the window.
' - hwndParent = [optional] Handle of the parent window.
' ========================================================================================
PRIVATE SUB DialogCenter (BYVAL hDlg AS HWND, BYVAL hwndParent AS HWND = NULL)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogCenter(hwndParent)
END SUB
' ========================================================================================

' ========================================================================================
' Ends the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogEnd (BYVAL hDlg AS HWND, BYVAL nResult AS LONG = 0) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg = NULL THEN RETURN FALSE
   pDlg->DialogEnd(nResult)
   Delete pDlg
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width and height of the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogGetSize (BYVAL hDlg AS HWND, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetSize(nWidth, nHeight)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the width of the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogGetWidth (BYVAL hDlg AS HWND) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetWidth
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the width of the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogGetHeight (BYVAL hDlg AS HWND) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width and height of the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogSetSize (BYVAL hDlg AS HWND, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogSetSize(nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounds of a window without the drop shadows.
' In Windows Vista and later, the Window Rect includes the area occupied by the drop shadow.
' Calling GetWindowRect will have different behavior depending on whether the window has
' ever been shown or not. If the window has not been shown before, GetWindowRect will not
' include the area of the drop shadow.
' To get the window bounds excluding the drop shadow, use DwmGetWindowAttribute, specifying
' DWMWA_EXTENDED_FRAME_BOUNDS. Note that unlike the Window Rect, the DWM Extended Frame Bounds
' are not adjusted for DPI. Getting the extended frame bounds can only be done after the
' window has been shown at least once.
' ========================================================================================
PRIVATE FUNCTION DialogGetBounds (BYVAL hDlg AS HWND) AS RECT
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetBounds
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the coordinates of the dialog's client area. The client coordinates specify the
' upper-left and lower-right corners of the client area. Because client coordinates are
' relative to the upper-left corner of a window's client area, the coordinates of the
' upper-left corner are (0,0).
' ========================================================================================
PRIVATE FUNCTION DialogGetClient (BYVAL hDlg AS HWND, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetClient(nWidth, nHeight)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DialogGetClientWidth (BYVAL hDlg AS HWND) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetClientWidth
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DialogGetClientHeight (BYVAL hDlg AS HWND) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetClientHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjusts the bounding rectangle of the dialog based on the desired size of the client area.
' Parameters:
' - nWidth = The new width of the client area of the dialog.
' - nHeight = The new height of the client area of the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogSetClient (BYVAL hDlg AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogSetClient(nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle of the font used by the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogGetFont (BYVAL hDlg AS HWND) AS HFONT
   RETURN CAST(HFONT, SendMessageW(hDlg, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the face name of the font used by the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogGetFontFaceName (BYVAL hDlg AS HWND) AS DWSTRING
   ' // Get the handle of the font used by the dialog
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hDlg, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN RETURN ""   
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   RETURN lfw.lfFaceName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the point size of the font used by the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogGetFontPointSize (BYVAL hDlg AS HWND) AS LONG
   ' // Get the handle of the font used by the dialog
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hDlg, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN RETURN 0
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(hDlg)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(hDlg, hDC)
   DIM ptSize AS LONG = MulDiv(lfw.lfHeight, 72, cyPixelsPerInch)
   IF ptSize < 0 THEN ptSize = -ptSize
   RETURN ptSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text in a dialog or window caption.
' ========================================================================================
PRIVATE FUNCTION DialogGetText (BYVAL hDlg AS HWND) AS DWSTRING
   DIM nLen AS LONG = SendMessageW(hDlg, WM_GETTEXTLENGTH, 0, 0)
   DIM dwsText AS DWSTRING = SPACE(nLen + 1)
   SendMessageW(hDlg, WM_GETTEXT, nLen + 1, cast(LPARAM, dwsText))
   RETURN LEFT(dwsText, LEN(dwsText) - 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in a dialog or window caption.
' ========================================================================================
PRIVATE FUNCTION DialogSetText (BYVAL hDlg AS HWND, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SetWindowTextW(hDlg, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the location of the top left corner of the window.
' The location is relative to the upper-left corner of the client area in the parent window.
' ========================================================================================
PRIVATE FUNCTION DialogGetLoc (BYVAL hDlg AS HWND, BYREF nLeft AS LONG, BYREF nTop AS LONG) AS BOOLEAN
   IF VARPTR(nLeft) = 0 OR VARPTR(nTop) = 0 THEN RETURN FALSE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogGetLoc(nLeft, nTop)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the location of the top left corner of the window.
' The location is relative to the upper-left corner of the client area in the parent window.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION DialogSetLoc (BYVAL hDlg AS HWND, BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogSetLoc(nLeft, nTop)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Retrieves a value from the user data area of the dialog.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE FUNCTION DialogGetUser (BYVAL hDlg AS HWND, BYVAL idx AS LONG) AS LONG_PTR
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogGetUser(idx)
END FUNCTION
' =====================================================================================
' =====================================================================================
' Sets a value in the user data area of a dialog.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE SUB DialogSetUser (BYVAL hDlg AS HWND, BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogGetUser(idx)
END SUB
' =====================================================================================

' =====================================================================================
' Sets the big and small icons of the dialog.
' =====================================================================================
PRIVATE SUB DialogSetIconEx (BYVAL hDlg AS HWND, BYVAL hIconBig AS HICON, BYVAL hIconSmall AS HICON)
   SendMessageW(hDlg, WM_SETICON, ICON_BIG, cast(LPARAM, hIconBig))
   SendMessageW(hDlg, WM_SETICON, ICON_SMALL, cast(LPARAM, hIconSmall))
END SUB
' =====================================================================================

' ========================================================================================
' Adds a control to the dialog.
' ========================================================================================
PRIVATE FUNCTION ControlAdd (BYREF wszClassName AS WSTRING, BYVAL hDlg AS HWND, BYVAL cId AS LONG_PTR = 0, _
BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL cx AS LONG = 0, BYVAL cy AS LONG = 0, _
BYVAL dwStyle AS LONG = -1, BYVAL dwExStyle AS LONG = -1, BYVAL lpParam AS LONG_PTR = 0, _
BYVAL pWndProc AS SUBCLASSPROC = NULL, BYVAL uIdSubclass AS UINT_PTR = &HFFFFFFFF, _
BYVAL dwRefData AS DWORD_PTR = NULL) AS HWND
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlAdd(wszClassName, cId, wszTitle, x, y, cx, cy, dwStyle, dwExStyle, _
      lpParam, pWndProc, uIdSubclass, dwRefData)
END FUNCTION
' ========================================================================================
' // For compatibility with DDT
#define ControlAddButton ControlAdd("Button")
#define ControlAddCheckState ControlAdd("CheckState")
#define ControlAddCheckBox ControlAdd("CheckBox")
#define ControlAddComboBox ControlAdd("ComboBox")
#define ControlAddFrame ControlAdd("GroupBox")
#define ControlAddImage ControlAdd("Image")
#define ControlAddImageX ControlAdd("ImageX")
#define ControlAddImgButton ControlAdd("ImgButton")
#define ControlAddImgButtonX ControlAdd("ImgButtonX")
#define ControlAddLabel ControlAdd("Label")
#define ControlAddLine ControlAdd("Line")
#define ControlAddListBox ControlAdd("ListBox")
#define ControlAddListBoxMultiSelect ControlAdd("ListBoxMultiSelect")
#define ControlAddListView ControlAdd("ListView")
#define ControlAddOption ControlAdd("Option")
#define ControlAddProgressBar ControlAdd("ProgressBar")
#define ControlAddScrollBar ControlAdd("ScrollBar")
#define ControlAddStatusBar ControlAdd("StatusBar")
#define ControlAddTab ControlAdd("Tab")
#define ControlAddTextBox ControlAdd("TextBox")
#define ControlAddToolBar ControlAdd("ToolBar")
#define ControlAddTreeView ControlAdd("TreView")
' // Additional controls
#define ControlAddAnimate ControlAdd("Animate")
#define ControlAddGroupBox ControlAdd("GroupBox")
#define ControlAddComboBoxEx ControlAdd("ComboBoxEx")
#define ControlAddHeader ControlAdd("Header")
#define ControlAddRebar ControlAdd("Rebar")
#define ControlAddDateTimePicker ControlAdd("DateTimePicker")
#define ControlAddMonthCalendar ControlAdd("MonthCalendar")
#define ControlAddIPAddress ControlAdd("IPAddress")
#define ControlAddHotkey ControlAdd("HotKey")
#define ControlAssSysLink ControlAdd("SysLink")
#define ControlAddPager ControlAdd("Pager")
#define ControlAddSizeBar ControlAdd("SizeBar")
#define ControlAddTrackBar ControlAdd("TrackBar")
#define ControlAddRichEdit ControlAdd("RichEdit")
#define ControlAddHScrollBar ControlAdd("HScrollBar")
#define ControlAddVScrollBar ControlAdd("VScrollBar")
#define ControlAddRadioButton ControlAdd("RadioButton")
#define ControlAddCustomButton ControlAdd("CustomButton")
#define ControlAddButtonOwnerDraw ControlAdd("CustomButton")
#define ControlAddCustomLabel ControlAdd("CustomLabel")
#define ControlAddLabelOwnerdraw ControlAdd("CustomLabel")
#define ControlAddFrameWindow ControlAdd("Frame")
#define ControlAddEditMultiline ControlAdd("EditMultiline")
#define ControlAddListBoxMultiSelect ControlAdd("ListBoxMultiSelect")
' ========================================================================================

' ========================================================================================
' Anchors the control
' ========================================================================================
PRIVATE FUNCTION ControlAnchor (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlAnchor(hDlg, cID, nAnchorMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus in the specified control of a dialog box.
' See https://devblogs.microsoft.com/oldnewthing/20040802-00/?p=38283
' ========================================================================================
PRIVATE SUB ControlSetFocus (BYVAL hDlg AS HWND, BYVAL cId AS LONG)
   SendMessageW(hDlg, WM_NEXTDLGCTL, cast(WPARAM, GetDlgItem(hDlg, cId)), CTRUE)
END SUB
' ========================================================================================

' ========================================================================================
' Disables the specified control.
' ========================================================================================
PRIVATE FUNCTION ControlDisable (BYVAL hDlg AS HWND, BYVAL cId AS LONG) AS BOOLEAN
   RETURN EnableWindow(GetDlgItem(hDlg, cId), FALSE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Enables the specified control.
' ========================================================================================
PRIVATE FUNCTION ControlEnable (BYVAL hDlg AS HWND, BYVAL cId AS LONG) AS BOOLEAN
   RETURN EnableWindow(GetDlgItem(hDlg, cId), CTRUE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the location of the top left corner of the window.
' The location is relative to the upper-left corner of the client area in the parent window.
' ========================================================================================
PRIVATE FUNCTION ControlGetLoc (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF nLeft AS LONG, BYREF nTop AS LONG) AS BOOLEAN
   IF VARPTR(nLeft) = 0 OR VARPTR(nTop) = 0 THEN RETURN FALSE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetLoc(cID, nLeft, nTop)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the location of the top left corner of the window.
' The location is relative to the upper-left corner of the client area in the parent window.
' Return value: TRUE or FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION ControlSetLoc (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetLoc(cID, nLeft, nTop)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text in a control or control caption.
' ========================================================================================
PRIVATE FUNCTION ControlGetText (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS DWSTRING
   DIM hCtl AS HWND = GetDlgItem(hDlg, cID)
   DIM nLen AS LONG = SendMessageW(hCtl, WM_GETTEXTLENGTH, 0, 0)
   DIM dwsText AS DWSTRING = WSPACE(nLen + 1)
   SendMessageW(hCtl, WM_GETTEXT, nLen + 1, cast(LPARAM, *dwsText))
   RETURN LEFT(dwsText, LEN(dwsText) - 1)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the text in a control or control caption.
' ========================================================================================
PRIVATE FUNCTION ControlSetText (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SetWindowTextW(GetDlgItem(hDlg, cID), pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a window handle for the specified control ID.
' ========================================================================================
PRIVATE FUNCTION ControlHandle (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS HWND
   RETURN GetDlgItem(hDlg, cID)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides the specified control.
' ========================================================================================
PRIVATE FUNCTION ControlHide (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS BOOLEAN
   RETURN ShowWindow(GetDlgItem(hDlg, cID), SW_HIDE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Makes visible the specified control.
' ========================================================================================
PRIVATE FUNCTION ControlNormalize (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS BOOLEAN
   RETURN ShowWindow(GetDlgItem(hDlg, cID), SW_SHOW)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys the specified control.
' ========================================================================================
PRIVATE FUNCTION ControlKill (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS BOOLEAN
   RETURN DestroyWindow(GetDlgItem(hDlg, cID))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the visible state of a control.
' ========================================================================================
PRIVATE FUNCTION ControlShowState (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL showState AS LONG) AS BOOLEAN
   RETURN ShowWindow(GetDlgItem(hDlg, cID), showState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redraws the specified control.
' ========================================================================================
PRIVATE FUNCTION ControlRedraw (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS BOOLEAN
   RETURN RedrawWindow(GetDlgItem(hDlg, cID), NULL, NULL, RDW_INVALIDATE OR RDW_UPDATENOW)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a logical font.
' Parameters:
' - wszFaceName = The typeface name.
' - lPointSize = The point size.
' - DPI = Dots per inch to calculate scaling. Default value = 96 (no scaling). If you pass -1
'   and the application is DPI aware, the DPI value used by the operating system will be used.
' - lWeight = The weight of the font in the range 0 through 1000. For example, 400 is normal
'      and 700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the font or NULL on failure.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' Usage examples:
'   hFont = FontNew("MS Sans Serif", 8, FW_NORMAL)
'   hFont = FontNew("Courier New", 10, FW_BOLD)
'   hFont = FontNew("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION FontNew (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, _
   BYVAL lWeight AS LONG = 0, BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, _
   BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT

   DIM tlfw AS LOGFONTW
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)

   ' // Font scaling
   DIM DPI AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   IF DPI > 0 THEN lPointSize = (lPointSize * DPI) \ GetDeviceCaps(hDC, LOGPIXELSY)

   tlfw.lfHeight         = -MulDiv(lPointSize, .GetDeviceCaps(hDC, LOGPIXELSY), 72)  ' logical font height
   tlfw.lfWidth          =  0                                                        ' average character width
   tlfw.lfEscapement     =  0                                                        ' escapement
   tlfw.lfOrientation    =  0                                                        ' orientation angles
   tlfw.lfWeight         =  lWeight                                                  ' font weight
   tlfw.lfItalic         =  bItalic                                                  ' italic(CTRUE/FALSE)
   tlfw.lfUnderline      =  bUnderline                                               ' underline(CTRUE/FALSE)
   tlfw.lfStrikeOut      =  bStrikeOut                                               ' strikeout(CTRUE/FALSE)
   tlfw.lfCharSet        =  bCharset                                                 ' character set
   tlfw.lfOutPrecision   =  OUT_TT_PRECIS                                            ' output precision
   tlfw.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                      ' clipping precision
   tlfw.lfQuality        =  DEFAULT_QUALITY                                          ' output quality
   tlfw.lfPitchAndFamily =  FF_DONTCARE                                              ' pitch and family
   tlfw.lfFaceName       =  wszFaceName                                              ' typeface name

   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = CreateFontIndirectW(@tlfw)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys a font when it is no longer needed.
' ========================================================================================
PRIVATE SUB FontEnd (BYVAL hFont AS HFONT)
   DeleteObject hFont
END SUB
' ========================================================================================

' ========================================================================================
' Gets the handle of the font used by the control.
' ========================================================================================
PRIVATE FUNCTION ControlGetFont (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS HFONT
   RETURN CAST(HFONT, SendMessageW(GetDLgItem(hDlg, cID), WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font that a control is to use when drawing text.
' ========================================================================================
PRIVATE SUB ControlSetFont (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL hFont AS HFONT, BYVAL fRedraw AS BOOLEAN = TRUE)
   DIM hCtl AS HWND = GetDlgItem(hDlg, cID)
   DIM hDC AS HDC = GetDC(hCtl)
   DeleteObject(SelectObject(hDC, CAST(HGDIOBJ, hFont)))
   ReleaseDC(hCtl, hDC)
   SendMessageW(hCtl, WM_SETFONT, CAST(WPARAM, hFont), fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the face name of the font used by the control.
' ========================================================================================
PRIVATE FUNCTION ControlGetFontFaceName (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS DWSTRING
   ' // Get the handle of the font used by the dialog
   DIM hCtl AS HWND = GetDlgItem(hDlg, cID)
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hCtl, WM_GETFONT, 0, 0))
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(hCtl)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(hCtl, hDC)
   RETURN lfw.lfFaceName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the point size of the font used by the control.
' ========================================================================================
PRIVATE FUNCTION ControlGetFontPointSize (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS LONG
   ' // Get the handle of the font used by the dialog
   DIM hCtl AS HWND = GetDlgItem(hDlg, cId)
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hCtl, WM_GETFONT, 0, 0))
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   DIM hDc AS HDC = GetDC(hCtl)
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC(hCtl, hDC)
   DIM ptSize AS LONG = MulDiv(lfw.lfHeight, 72, cyPixelsPerInch)
   IF ptSize < 0 THEN ptSize = -ptSize
   RETURN ptSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets/gets the return value of a message processed in the dialog box procedure.
' Return value: TRUE or FALSE.
' Must be set just before exiting the callback message (not after calling other functions)
' See https://forum.powerbasic.com/forum/user-to-user-discussions/programming/19130-info-dwl_msgresult
' Messages that may need the use of this method:
' WM_CTLCOLORMSGBOX, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_CTLCOLORBTN, WM_CTLCOLORDLG
' WM_CTLCOLORSCROLLBAR, WM_CTLCOLORSTATIC, WM_COMPAREITEM, WM_VKEYTOITEM, WM_CHARTOITEM
' WM_QUERYDRAGICON, WM_INITDIALOG.
' #define SetDlgMsgResult(hwnd,msg,result) (iif((msg)=WM_CTLCOLORMSGBOX orelse (msg)=WM_CTLCOLOREDIT
' orelse (msg)=WM_CTLCOLORLISTBOX orelse (msg)=WM_CTLCOLORBTN orelse (msg)=WM_CTLCOLORDLG
' orelse (msg)=WM_CTLCOLORSCROLLBAR orelse (msg)=WM_CTLCOLORSTATIC orelse (msg)=WM_COMPAREITEM
' orelse (msg)=WM_VKEYTOITEM orelse (msg)=WM_CHARTOITEM orelse (msg)=WM_QUERYDRAGICON
' orelse (msg)=WM_INITDIALOG, cast(WINBOOL, (result)), (SetWindowLongPtr((hwnd),DWLP_MSGRESULT,
' cast(LPARAM, cast(LRESULT, result)),TRUE))))
' See also this info message from Dominic Mitchell in the PowerBasic forum:
' https://forum.powerbasic.com/forum/user-to-user-discussions/programming/19130-info-dwl_msgresult
' And Rasymond Chen: https://devblogs.microsoft.com/oldnewthing/20031112-00/?p=41863
' Note: If you return a result using this function, your callback procedure must return TRUE.
' ========================================================================================
PRIVATE FUNCTION CBSetDlgMsgResult(BYVAL hDlg AS HWND, BYVAL result AS LONG_PTR) AS LONG_PTR
   RETURN SetWindowLongPtrW(hDlg, DWLP_MSGRESULT, result)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CBGetDlgMsgResult (BYVAL hDlg AS HWND) AS LONG_PTR
   RETURN GetWindowLongPtrW(hDlg, DWLP_MSGRESULT)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places (posts) a message in the message queue associated with the thread that created the
' specified dialog and returns without waiting for the thread to process the message.
' ========================================================================================
PRIVATE FUNCTION DialogPost (BYVAL hDlg AS HWND, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN PostMessageW(hDlg, Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the specified message to a window or windows. The SendMessage function calls the
' window procedure for the dialog and does not return until the dialog procedure has
' processed the message.
' ========================================================================================
PRIVATE FUNCTION DialogSend (BYVAL hDlg AS HWND, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN SendMessageW(hDlg, Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places (posts) a message in the message queue associated with the thread that created the
' specified window and returns without waiting for the thread to process the message.
' ========================================================================================
PRIVATE FUNCTION ControlPost (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN PostMessageW(GetDlgItem(hDlg, cID), Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sends the specified message to a window or windows. The SendMessage function calls the
' window procedure for the specified window and does not return until the window procedure
' has processed the message.
' ========================================================================================
PRIVATE FUNCTION ControlSend (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL Msg AS UINT, BYVAL wPrm AS LONG_PTR, BYVAL lPrm AS LONG_PTR) AS BOOLEAN
   RETURN SendDlgItemMessageW(hDlg, cID, Msg, cast(WPARAM, wPrm), cast(LPARAM, lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width and height of the control.
' ========================================================================================
PRIVATE FUNCTION ControlGetSize (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetSize(cID, nWidth, nHeight)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the width of the control.
' ========================================================================================
PRIVATE FUNCTION ControlGetWidth (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetWidth(cID)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the height of the control.
' ========================================================================================
PRIVATE FUNCTION ControlGetHeight (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetHeight(cID)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size of the specified window.
' - nWidth = The new width of the window.
' - nHeight = The new height of the window.
' ========================================================================================
PRIVATE FUNCTION ControlSetSize (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetSize(cID, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the coordinates of the control's client area. The client coordinates specify the
' upper-left and lower-right corners of the client area. Because client coordinates are
' relative to the upper-left corner of a window's client area, the coordinates of the
' upper-left corner are (0,0).
' ========================================================================================
PRIVATE FUNCTION ControlGetClient (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF nWidth AS LONG, BYREF nHeight AS LONG) AS BOOLEAN
   IF VARPTR(nWidth) = 0 OR VARPTR(nHeight) = 0 THEN RETURN FALSE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetClient(cID, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the the control's client area
' ========================================================================================
PRIVATE FUNCTION ControlGetClientWidth (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetClientWidth(cID)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the height of the the control's client area
' ========================================================================================
PRIVATE FUNCTION ControlGetClientHeight (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS LONG
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlGetClientHeight(cID)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjusts the bounding rectangle of the dialog based on the desired size of the client area.
' Parameters:
' - nWidth = The new width of the client area of the dialog.
' - nHeight = The new height of the client area of the dialog.
' ========================================================================================
PRIVATE FUNCTION ControlSetClient (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetClient(cID, nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the control to the center of the dialog horizontally
' - cId = Identifier of the control
' ========================================================================================
PRIVATE SUB ControlCenterHoriz (BYVAL hDlg AS HWND, BYVAL cID AS LONG)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->ControlCenterHoriz(cID)
END SUB
' ========================================================================================

' ========================================================================================
' Moves the control to the center of the dialog vertically
' - cId = Identifier of the control
' ========================================================================================
PRIVATE SUB ControlCenterVert (BYVAL hDlg AS HWND, BYVAL cID AS LONG)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->ControlCenterVert(cID)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the check state of a radio button or check box. 
' ========================================================================================
PRIVATE FUNCTION ControlGetCheck (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS LONG
   DIM hCtrl AS HWND = GetDlgItem(hDlg, cID)
   RETURN SendMessageW(hCtrl, BM_GETCHECK, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the check state of a radio button or check box.
' ========================================================================================
PRIVATE SUB ControlSetCheck (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL checkState AS LONG)
   DIM hCtrl AS HWND = GetDlgItem(hDlg, cID)
   SendMessageW(hCtrl, BM_SETCHECK, checkState, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the check state for an Option (radio) control, and unsets the check state for other Option buttons in a group.
' ========================================================================================
PRIVATE FUNCTION ControlSetOption (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL minId AS LONG, BYVAL maxId AS LONG) AS BOOLEAN
   RETURN ..CheckRadioButton(GetDlgItem(hDlg, cId), minId, maxId, cId)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the identifier of the default push button control for a dialog box.
' ========================================================================================
PRIVATE FUNCTION GetDefId (BYVAL hDlg AS HWND) AS INT_PTR
   RETURN LOWORD(SendMessageW(hDlg, DM_GETDEFID, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the identifier of the default push button for a dialog box.
' ========================================================================================
PRIVATE FUNCTION SetDefId (BYVAL hDlg AS HWND, BYVAL cID AS LONG) AS BOOLEAN
   ' // Removes the default button because 32768 is not a valid control ID.
   SendMessageW(hDlg, DM_SETDEFID, 32768, 0)
   ' // Sets the specified button as the default
   RETURN SendMessageW(hDlg, DM_SETDEFID, cID, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Repositions a top-level dialog box so that it fits within the desktop area. 
' ========================================================================================
PRIVATE FUNCTION DialogReposition (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN SendMessageW(hDlg, DM_REPOSITION, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Change the visible state of a dialog.
' ========================================================================================
PRIVATE FUNCTION DialogShowState (BYVAL hDlg AS HWND, BYVAL showState AS LONG) AS BOOLEAN
   RETURN ShowWindow(hDlg, showState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maximizes the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogMaximize (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN ShowWindow(hDlg, SW_MAXIMIZE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Minimizes the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogMinimize (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN ShowWindow(hDlg, SW_MINIMIZE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogHide (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN ShowWindow(hDlg, SW_HIDE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Makes the dialog visible at its normal size and position.
' ========================================================================================
PRIVATE FUNCTION DialogNormalize (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN ShowWindow(hDlg, SW_SHOWNORMAL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Disables the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogDisable (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN EnableWindow(hDlg, FALSE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogEnable (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN EnableWindow(hDlg, CTRUE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redraws the dialog.
' ========================================================================================
PRIVATE FUNCTION DialogRedraw (BYVAL hDlg AS HWND) AS BOOLEAN
   RETURN RedrawWindow(hDlg, NULL, NULL, RDW_INVALIDATE OR RDW_ERASE OR RDW_UPDATENOW)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If you use dual (or even triple/quad) displays then you have undoubtedly encountered the
' following situation: You change the physical order of your displays, or otherwise
' reconfigure the logical ordering using your display software. This sometimes has the
' side-effect of changing your desktop coordinates from zero-based to negative starting
' coordinates (i.e. the top-left coordinate of your desktop changes from 0,0 to -1024,-768).
' This effects many Windows programs which restore their last on-screen position whenever
' they are started. Should the user reorder their display configuration this can sometimes
' result in a Windows program subsequently starting in an off-screen position (i.e. at a
' location that used to be visible) - and is now effectively invisible, preventing the
' user from closing it down or otherwise moving it back on-screen.
' The DialogForceVisibility method can be called at program start-time right after the
' main window has been created and positioned 'on-screen'. Should the window be positioned
' in an off-screen position, it is forced back onto the nearest display to its last
' position. The user will be unaware this is happening and won't even realise to thank you
' for keeping their user-interface visible, even though they changed their display
' settings.
' Source: http://www.catch22.net/tuts/tips2
' ========================================================================================
PRIVATE SUB DialogForceVisibility (BYVAL hDlg AS HWND)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogForceVisibility
END SUB
' ========================================================================================

' ========================================================================================
' Makes a dialog stabilized (non-closeable).
' ========================================================================================
PRIVATE FUNCTION DialogStabilize (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogStabilize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Makes a dialog non stable (closeable).
' ========================================================================================
PRIVATE FUNCTION DialogNonStable (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogNonStable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns true if the dialog is stabilized; otherwise, false.
' ========================================================================================
PRIVATE FUNCTION IsDialogStabilized (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->IsDialogStabilized
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns true if the dialog is stabilized; otherwise, false.
' ========================================================================================
PRIVATE FUNCTION IsDialogNonStable (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->IsDialogNonStable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns true if the dialog uses pixels; otherwise, false.
' ========================================================================================
PRIVATE FUNCTION DialogUsesPixels (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->UsesPixels
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if the dialog uses units; otherwise, false.
' ========================================================================================
PRIVATE FUNCTION DialogUsesUnits (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->UsesUnits
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns true if the dialog uses units; otherwise, false.
' ========================================================================================
PRIVATE FUNCTION DialogEndResult (BYVAL hDlg AS HWND) AS INT_PTR
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogEndResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Makes a dialog scrollable by shrinking its client area size
' ========================================================================================
PRIVATE FUNCTION DialogSetViewPort (BYVAL hDlg AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogSetViewPort(nWidth, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes both the dialog icon in the caption, and the icon shown in the ALT+TAB task list.
' The old image handle is released.
' ========================================================================================
PRIVATE SUB DialogSetIcon (BYVAL hDlg AS HWND, BYREF wszImage AS WSTRING)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogSetIcon(wszImage)
END SUB
' ========================================================================================

' =====================================================================================
' Finds the resource type given it's identifier or name.
' =====================================================================================
PRIVATE FUNCTION FindResourceType (BYVAL hDlg AS HWND, BYREF wszResourceName AS WSTRING) AS STRING
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->FindResourceType(wszResourceName)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the icon or bitmap displayed in a label.
' The size of the label is adjusted to the size of the image.
' The old image handle is released.
' ========================================================================================
PRIVATE FUNCTION ControlSetImage (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszImage AS WSTRING) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetImage(cID, wszImage)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the icon or bitmap displayed in a label
' The size of the image is streched to fit the size of the control.
' The old image handle is released.
' ========================================================================================
PRIVATE FUNCTION ControlSetImageX (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszImage AS WSTRING) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetImageX(cID, wszImage)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the icon or bitmap displayed in a button.
' The size of the button does not change; The image is clipped if it does nt fit in the button.
' The old image handle is released.
' ========================================================================================
PRIVATE FUNCTION ControlSetImgButton (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszImage AS WSTRING) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetImgButton(cID, wszImage)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the icon or bitmap displayed in a button.
' The size of the button does not change; The image is clipped if it does nt fit in the button.
' The old image handle is released.
' ========================================================================================
PRIVATE FUNCTION ControlSetImgButtonX (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszImage AS WSTRING) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetImgButtonX(cID, wszImage)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the colors of the control
' ========================================================================================
PRIVATE FUNCTION ControlSetColor (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL foreColor AS COLORREF, BYVAL backColor AS COLORREF) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ControlSetColor(cID, foreColor, backColor)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of the dialog
' ========================================================================================
PRIVATE FUNCTION DialogSetColor (BYVAL hDlg AS HWND, BYVAL foreColor AS COLORREF, BYVAL backColor AS COLORREF) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->DialogSetColor(foreColor, backColor)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enable/disable dialog repainting during resizing
' ========================================================================================
PRIVATE SUB DialogEnableRepaint (BYVAL hDlg AS HWND, BYVAL fRepaint AS BOOLEAN)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogEnableRepaint(fRepaint)
END SUB
' ========================================================================================
' ========================================================================================
' Checks if repainting is enabled during resizing
' ========================================================================================
PRIVATE FUNCTION IsDialogRepaintDisabled (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->IsDialogRepaintDisabled
END FUNCTION
' ========================================================================================
' ========================================================================================
' Enable/disable dialog repainting during resizing
' ========================================================================================
PRIVATE SUB DialogDisableRepaintOnResize (BYVAL hDlg AS HWND, BYVAL fRepaint AS BOOLEAN)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DialogDisableRepaintOnResize(fRepaint)
END SUB
' ========================================================================================
' ========================================================================================
' Check if repainting is enabled during resizing
' ========================================================================================
PRIVATE FUNCTION IsDialogRepaintDisabledOnResize (BYVAL hDlg AS HWND) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->IsDialogRepaintDisabledOnResize
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                 *** COMBOBOX ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION ComboBoxAdd (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, CB_ADDSTRING, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxDelete (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, CB_DELETESTRING, item - 1, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxFind (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   IF item <= 0 THEN RETURN 0
   DIM start aS LONG = item
   IF item = 1 THEN start = 2
   DIM res AS LONG = SendDlgItemMessageW(hDlg, id, CB_FINDSTRING, start - 2, CAST(LPARAM, pwszText))
   IF item = 1 THEN RETURN res ELSE RETURN res + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxFindExact (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   IF item <= 0 THEN RETURN 0
   DIM start aS LONG = item
   IF item = 1 THEN start = 2
   DIM res AS LONG = SendDlgItemMessageW(hDlg, id, CB_FINDSTRINGEXACT, start - 2, CAST(LPARAM, pwszText))
   RETURN res + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxGetCount (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, CB_GETCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxGetSelCount (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS LONG
   DIM item AS LONG = SendDlgItemMessageW(hDlg, id, CB_GETCURSEL, 0, 0)
   IF item <> CB_ERR THEN RETURN 1 ELSE RETURN 0
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxGetSelect (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, CB_GETCURSEL, 0, 0) + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxGetState (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG) AS BOOLEAN
   DIM selItem AS LONG = SendDlgItemMessageW(hDlg, id, CB_GETCURSEL, 0, 0)
   IF selItem + 1 = item THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxGetText (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG = 0) AS DWSTRING
   DIM ItemText AS DWSTRING
   DIM cbLen AS LONG = SendDlgItemMessageW(hDlg, id, CB_GETLBTEXTLEN, item - 1, 0)
   IF cbLen < 1 THEN
      DIM selItem AS LONG = SendDlgItemMessageW(hDlg, id, CB_GETCURSEL, 0, 0)
      IF selItem >= 0 THEN
         cbLen = SendDlgItemMessageW(hDlg, id, CB_GETLBTEXTLEN, selItem, 0)
         itemText = WSPACE(cbLen + 1)
         SendDlgItemMessageW(hDlg, id, CB_GETLBTEXT, selItem, cast(LPARAM, STRPTR(itemText)))
      END IF
   ELSE
      itemText = WSPACE(cbLen + 1)
      SendDlgItemMessageW(hDlg, id, CB_GETLBTEXT, item - 1, cast(LPARAM, STRPTR(itemText)))
   END IF
   RETURN LEFT(itemText, cbLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxInsert (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   IF item < 1 THEN RETURN CB_ERR
   RETURN SendDlgItemMessageW(hDlg, id, CB_INSERTSTRING, item - 1, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB ComboBoxReset (BYVAL hDlg AS HWND, BYVAL id AS LONG)
   SendDlgItemMessageW hDlg, id, CB_RESETCONTENT, 0, 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxSelect (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, CB_SETCURSEL, item - 1, 0) + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxSetText (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   SendDlgItemMessageW(hDlg, id, CB_DELETESTRING, item - 1, 0)
   SendDlgItemMessageW(hDlg, id, CB_INSERTSTRING, item - 1, cast(LPARAM, pwszText))
   RETURN SendDlgItemMessageW(hDlg, id, CB_SETCURSEL, item - 1, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxUnselect (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, CB_SETCURSEL, -1, 0)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                   *** LISTBOX ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION ListBoxAdd (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, LB_ADDSTRING, 0, CAST(LPARAM, pwszText)) + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBoxDelete (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, LB_DELETESTRING, item - 1, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBoxFind (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   IF item <= 0 THEN RETURN 0
   DIM start aS LONG = item
   IF item = 1 THEN start = 2
   DIM res AS LONG = SendDlgItemMessageW(hDlg, id, LB_FINDSTRING, start - 2, CAST(LPARAM, pwszText))
   IF item = 1 THEN RETURN res ELSE RETURN res + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBoxFindExact (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   IF item <= 0 THEN RETURN 0
   DIM start aS LONG = item
   IF item = 1 THEN start = 2
   DIM res AS LONG = SendDlgItemMessageW(hDlg, id, LB_FINDSTRINGEXACT, start - 2, CAST(LPARAM, pwszText))
   RETURN res + 1
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the number of items in a list box.
' ========================================================================================
PRIVATE FUNCTION ListBoxGetCount (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, LB_GETCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the total number of selected items in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION ListBoxGetSelCount (BYVAL hDlg AS HWND, BYVAL id AS LONG) AS LONG
   RETURN SendDlgItemMessageW(hDlg, id, LB_GETSELCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' The listbox is searched to find the first selected item. If the item parameter is included,
' searching starts at that position to facilitate retrieving multiple selected items.
' If item is omitted, the search starts at the first data item. The index number of the
' selected item is returned as the result of the function. If no item is selected, the
' value -1 is returned.
' ========================================================================================
PRIVATE FUNCTION ListBoxGetSelect (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG = 0) AS LONG
   DIM hCtl AS HWND = GetDlgItem(hDlg, id)
   IF (GetWindowLongPtrW(hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      ' // Search for the first selected item starting from startIndex (zero-based)
      DIM nCount AS LONG = SendMessageW(hCtl, LB_GETCOUNT, 0, 0)
      IF item = 0 THEN item = 1
      FOR selectedIndex AS LONG = item - 1 TO nCount - 1
         IF SendMessageW(hCtl, LB_GETSEL, item, 0) <> 0 THEN
            RETURN item + 1   ' one-based inxdes
         END IF
      NEXT
      ' // If no item is selected, return 0
      RETURN 0
   ELSE
      '// Single slection listbox
      RETURN SendMessageW(hCtl, IIF((GetWindowLong(hCtl, GWL_STYLE) AND LBS_MULTIPLESEL), LB_GETCARETINDEX, LB_GETCURSEL), 0, 0) + 1
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
' Checks if the item is selected or not.
' ========================================================================================
PRIVATE FUNCTION ListBoxGetState (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG) AS BOOLEAN
   RETURN SendDlgItemMessageW(hDlg, id, LB_GETSEL, item - 1, 0) > 0
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets a string from a list box.
' ========================================================================================
PRIVATE FUNCTION ListBoxGetText (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG = 0) AS DWSTRING
   DIM nLen AS LONG, buffer AS DWSTRING
   IF item = 0 THEN item = ListBoxGetSelect(hDlg, id, 0)
   nLen = SendDlgItemMessageW(hDlg, id, LB_GETTEXTLEN, item - 1, 0)
   SendDlgItemMessageW hDlg, id, LB_GETTEXT, item - 1, CAST(LPARAM, STRPTR(buffer))
   RETURN LEFT(buffer, nLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Removes all items from a list box.
' ========================================================================================
PRIVATE SUB ListBoxReset (BYVAL hDlg AS HWND, BYVAL id AS LONG)
   SendDlgItemMessageW hDlg, id, LB_RESETCONTENT, 0, 0
END SUB
' ========================================================================================
' ========================================================================================
' * Selects the specified string in a list box.
' ========================================================================================
PRIVATE FUNCTION ListBoxSelect (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL index AS LONG = 0) AS LONG
   DIM hCtl AS HWND = GetDlgItem(hDlg, id)
   IF (GetWindowLongPtrW(hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      IF index = 0 THEN RETURN SendMessageW(hCtl, LB_SETSEL, CTRUE, index) + 1
      RETURN SendMessageW(hCtl, LB_SETSEL, CTRUE, index - 1) + 1
   ELSE
      '// Single slection listbox
      IF index = 0 THEN RETURN SendMessageW(hCtl, LB_SETSEL, CTRUE, index) + 1
      RETURN SendMessageW(hCtl, LB_SETCURSEL, index - 1, 0) + 1
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBoxSetText (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   '// Delete the string
   DIM lRes AS LONG = SendDlgItemMessageW(hDlg, id, LB_DELETESTRING, item - 1, 0)
   IF lRes = LB_ERR THEN RETURN lRes
   ' // Insert the new string
   SendDlgItemMessageW(hDlg, id, LB_INSERTSTRING, item - 1, cast(LPARAM, pwszText))
   IF item = LB_ERR OR item = LB_ERRSPACE THEN RETURN item
   DIM hCtl AS HWND = GetDlgItem(hDlg, id)
   RETURN SendMessageW(hCtl, IIF((GetWindowLongPtrW(hCtl, GWL_STYLE) AND LBS_MULTIPLESEL), LB_SETSEL, LB_SETCURSEL), item - 1, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Unselects the specified string in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION ListBoxUnselect (BYVAL hDlg AS HWND, BYVAL id AS LONG, BYVAL item AS LONG = 0) AS LONG
   IF (GetWindowLongPtrW(GetDlgItem(hDlg, id), GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      RETURN SendDlgItemMessageW(hDlg, id, LB_SETSEL, FALSE, item - 1)
   ELSE
      ' // Single selection listbox
      RETURN SendDlgItemMessageW(hDlg, id, LB_SETCURSEL, -1, 0)
   END IF
END FUNCTION
' ========================================================================================

'+++++++++++++++++


' ########################################################################################
'                                 *** ACCELERATORS ***
' ########################################################################################

' =====================================================================================
' Gets the accelerator table handle
' =====================================================================================
PRIVATE FUNCTION AccelHandle (BYVAL hDlg AS HWND) AS HACCEL
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->AccelHandle
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB AccelAttach (BYVAL hDlg AS HWND, BYVAL hAccel AS HACCEL)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->AccelAttach(hAccel)
END SUB
' =====================================================================================
' =====================================================================================
' Adds an accelerator key to the table.
' =====================================================================================
PRIVATE SUB AddAccelerator OVERLOAD (BYVAL hDlg AS HWND, BYVAL fvirt AS UBYTE, BYVAL wKey AS WORD, BYVAL cmd AS WORD)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->AddAccelerator(fvirt, wKey, cmd)
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB AddAccelerator OVERLOAD (BYVAL hDlg AS HWND, BYVAL fvirt AS UBYTE, BYREF wszKey AS WSTRING, BYVAL cmd AS WORD)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->AddAccelerator(fvirt, wszKey, cmd)
END SUB
' =====================================================================================
' =====================================================================================
' Creates the accelerator table.
' =====================================================================================
PRIVATE FUNCTION CreateAccelTable (BYVAL hDlg AS HWND) AS HACCEL
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->CreateAccelTable
END FUNCTION
' =====================================================================================
' =====================================================================================
' Destroys the accelerator table.
' =====================================================================================
PRIVATE SUB DestroyAccelTable (BYVAL hDlg AS HWND)
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN pDlg->DestroyAccelTable
END SUB
' =====================================================================================

' ########################################################################################
'                              *** DPI AND CONVERSIONS ***
' ########################################################################################

' ========================================================================================
' Returns the DPI scaling ratios
' ========================================================================================
PRIVATE FUNCTION rxRatioForDpi (BYVAL hDlg AS HWND) AS SINGLE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->rxRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ryRatioForDpi (BYVAL hDlg AS HWND) AS SINGLE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ryRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales a horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION ScaleXForDpi (BYVAL hDlg AS HWND, BYVAL cx AS SINGLE) AS SINGLE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ScaleX(cx)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Scales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION ScaleYForDpi (BYVAL hDlg AS HWND, BYVAL cy AS SINGLE) AS SINGLE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ScaleY(cy)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales a horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION UnScaleXForDpi (BYVAL hDlg AS HWND, BYVAL cx AS SINGLE) AS SINGLE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->UnScaleX(cx)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Unscales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION UnScaleYForDpi (BYVAL hDlg AS HWND, BYVAL cy AS SINGLE) AS SINGLE
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->UnScaleY(cy)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales a RECT structure
' ========================================================================================
PRIVATE FUNCTION ScaleRectForDpi (BYVAL hDlg AS HWND, BYVAL rc AS RECT) AS RECT
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->ScaleRect(rc)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Unscales a RECT structure.
' ========================================================================================
PRIVATE FUNCTION UnScaleRectForDpi (BYVAL hDlg AS HWND, BYVAL rc AS RECT) AS RECT
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->UnScaleRect(rc)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the size, position, and Z order of a child, pop-up, or top-level window.
' ========================================================================================
PRIVATE FUNCTION SetWindowPosForDPI OVERLOAD (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL hWndInsertAfter AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, _
   BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL uFlags AS UINT = SWP_NOZORDER) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->SetWindowPosForDPI(GetDlgItem(hDlg, cID), hWndInsertAfter, x, y, cx, cy, uFlags)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION SetWindowPosForDPI OVERLOAD (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL x AS LONG, BYVAL y AS LONG, _
   BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL uFlags AS UINT = SWP_NOZORDER) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->SetWindowPosForDPI(GetDlgItem(hDlg, cID), NULL, x, y, cx, cy, uFlags)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the position and dimensions of the specified window.
' ========================================================================================
PRIVATE FUNCTION MoveWindowForDPI (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL bRepaint AS BOOLEAN) AS BOOLEAN
   DIM pDlg AS CDialog PTR = GetCDialogPtr(hDlg)
   IF pDlg THEN RETURN pDlg->MoveWindowForDPI(GetDlgItem(hDlg, cID), x, y, nWidth, nHeight, bRepaint)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts the specified dialog box units to screen units (pixels).
' ========================================================================================
PRIVATE FUNCTION DialogUnitsToPixels (BYVAL hDlg AS HWND, BYVAL dluX AS LONG, BYVAL dluY AS LONG, BYREF pixelX AS LONG, BYREF pixelY AS LONG) AS BOOLEAN
   IF VARPTR(pixelX) = 0 OR VARPTR(pixelY) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   rc.Right = dluX : rc.Bottom = dluY
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(pixelX) THEN pixelX = rc.Right
   IF VARPTR(pixelY) THEN pixelY = rc.Bottom
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratios from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION DialogUnitsToPixelsRatios (BYVAL hDlg AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   IF VARPTR(ratioX) = 0 OR VARPTR(ratioY) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(ratioX) THEN ratioX = rc.Right / 4
   IF VARPTR(ratioY) THEN ratioY = rc.Bottom / 8
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION DluToPixRX (BYVAL hDlg AS HWND) AS SINGLE
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   RETURN rc.Right / 4
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION DluToPixRY (BYVAL hDlg AS HWND) AS SINGLE
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   RETURN rc.Bottom / 8
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts the specified screen units (pixels) to dialog box units.
' ========================================================================================
PRIVATE FUNCTION PixelsToDialogUnits (BYVAL hDlg AS HWND, BYVAL pixelX AS LONG, BYVAL pixelY AS LONG, BYREF dluX AS LONG, BYREF dluY AS LONG) AS BOOLEAN
   IF VARPTR(pixelX) = 0 OR VARPTR(pixelY) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   rc.Right = 4
   rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(dluX) THEN dluX = pixelX * (4.0 / rc.right)
   IF VARPTR(dluY) THEN dluY = pixelY * (8.0 / rc.bottom)
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratios from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION PixelsToDialogUnitsRatios (BYVAL hDlg AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   IF VARPTR(ratioX) = 0 OR VARPTR(ratioY) = 0 THEN RETURN FALSE
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(ratioX) THEN ratioX = 4.0 / rc.right
   IF VARPTR(ratioY) THEN ratioY = 8.0 / rc.bottom
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION PixToDluRX (BYVAL hDlg AS HWND) AS SINGLE
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN 4.0 / rc.right
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION PixToDluRY (BYVAL hDlg AS HWND) AS SINGLE
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN 8.0 / rc.bottom
END FUNCTION
' ========================================================================================

END NAMESPACE

