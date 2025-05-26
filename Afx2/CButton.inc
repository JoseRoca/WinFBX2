' ########################################################################################
' Platform: Microsoft Windows
' Filename: CButton.inc
' Purpose:  Windows common controls
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
#include once "windows.bi"
#include once "win/commctrl.bi"
#include once "Afx2/CControl.inc"
' // conflicting define in windowsx.bi
#undef DeleteBitmap
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CButton class
' ########################################################################################

TYPE CButton EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION CheckDlgButton (BYVAL cIDButton AS LONG, BYVAL uCheck AS DWORD) AS BOOLEAN
   DECLARE FUNCTION CheckRadioButton (BYVAL cIDFirstButton AS LONG, BYVAL cIDLastButton AS LONG, BYVAL cIDCheckButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsDlgButtonChecked (BYVAL cIDButton AS LONG) AS UINT
   DECLARE FUNCTION Disable () AS BOOLEAN
   DECLARE FUNCTION Enable () AS BOOLEAN
   DECLARE SUB Click ()
   DECLARE FUNCTION GetTextLength () AS LONG
   DECLARE FUNCTION GetText () AS DWSTRING
   DECLARE FUNCTION SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION GetImage (BYVAL imageType aS LONG) AS HANDLE
   DECLARE FUNCTION GetBitmap () AS HANDLE
   DECLARE FUNCTION GetIcon () AS HANDLE
   DECLARE FUNCTION SetImage (BYVAL ImageType AS DWORD, BYVAL hImage AS HANDLE) AS HANDLE
   DECLARE FUNCTION SetBitmap (BYVAL hbmp AS HBITMAP) AS HBITMAP
   DECLARE FUNCTION SetIcon (BYVAL hIcon AS HICON) AS HICON
   DECLARE FUNCTION GetImageList () AS BUTTON_IMAGELIST PTR
   DECLARE FUNCTION SetImageList (BYVAL pbuttonImagelist AS BUTTON_IMAGELIST PTR) AS BOOLEAN
   DECLARE FUNCTION SetImageList (BYVAL himl AS HIMAGELIST, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG, BYVAL uALign AS DWORD = 0) AS BOOLEAN
   DECLARE FUNCTION DeleteBitmap () AS BOOLEAN
   DECLARE FUNCTION DeleteIcon () AS BOOLEAN
   DECLARE FUNCTION DeleteImage (BYVAL ImageType AS DWORD) AS BOOLEAN
   DECLARE FUNCTION GetCheck () AS DWORD
   DECLARE SUB SetCheck (BYVAL checkState AS LONG)
   DECLARE SUB Check ()
   DECLARE SUB Uncheck ()
   DECLARE SUB Gray ()
   DECLARE FUNCTION GetState () AS DWORD
   DECLARE SUB SetState (BYVAL bstate AS BOOLEAN)
   DECLARE FUNCTION GetStyle () AS DWORD
   DECLARE SUB SetStyle (BYVAL dwStyle AS DWORD, BYVAL bRedraw AS BOOLEAN)
   DECLARE FUNCTION GetIdealSize (BYREF tSize AS SIZE) AS BOOLEAN
   DECLARE FUNCTION GetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
   DECLARE FUNCTION SetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
#if _WIN32_WINNT >= &h0600
   DECLARE FUNCTION SetDropDownState (BYVAL bDropDown AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION GetSplitInfo (BYREF info AS BUTTON_SPLITINFO) AS BOOLEAN
   DECLARE FUNCTION SetSplitInfo (BYREF info AS BUTTON_SPLITINFO) AS BOOLEAN
   DECLARE SUB SetDontClick (BYVAL bState AS BOOLEAN)
   DECLARE FUNCTION GetNoteLength () AS LONG
   DECLARE FUNCTION GetNote () AS DWSTRING
   DECLARE FUNCTION SetNote (BYREF pwsz AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION SetElevationRequiredState (BYVAL bRequired AS BOOLEAN) AS LONG
#endif

END TYPE
' ========================================================================================

' ========================================================================================
' * CButton constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CButton (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CButton (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CButton (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CButton (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CButton destructor
' ========================================================================================
PRIVATE DESTRUCTOR CButton
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Enables the control.
' Return value: False if the control was previous enabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CButton.Enable () AS BOOLEAN
   RETURN EnableWindow(m_hCtl, TRUE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Disables the control.
' Return value: False if the control was previous disabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CButton.Disable () AS BOOLEAN
   DIM res AS LONG = EnableWindow(m_hCtl, FALSE)
   DIM bEnable AS BOOLEAN
   IF res = 0 THEN bEnable = TRUE ELSE bEnable = FALSE
   RETURN bEnable
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the length of the text in a button control.
' ========================================================================================
PRIVATE FUNCTION CButton.GetTextLength () AS LONG
   RETURN GetWindowTextLengthW(m_hCtl)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the text in a button control.
' ========================================================================================
PRIVATE FUNCTION CButton.GetText () AS DWSTRING
   DIM nLen AS LONG = GetWindowTextLengthW(m_hCtl)
   DIM dwsText AS DWSTRING = WSPACE(nLen + 1)
   ..GetWindowTextW(m_hCtl, dwsText, nLen +  1)
   RETURN dwsText
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Sets the text of a button.
' ========================================================================================
PRIVATE FUNCTION CButton.SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SetWindowTextW(m_hCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Simulates the user clicking a button. This message causes the button to receive the
' WM_LBUTTONDOWN and WM_LBUTTONUP messages, and the button's parent window to receive a
' BN_CLICKED notification code.
' Usage example: CButton(pDlg, 105).Click (where 105 is the identifier of the button).
' ========================================================================================
PRIVATE SUB CButton.Click ()
   SendMessageW(m_hCtl, BM_CLICK, 0, 0)
END SUB
' ========================================================================================

#if _WIN32_WINNT >= &h0600
' ========================================================================================
' Sets a flag on a radio button that controls the generation of BN_CLICKED messages when
' the button receives focus.
' - bState = A BOOLEAN that specifies the state. TRUE to set the flag, otherwise FALSE.
' ========================================================================================
PRIVATE SUB CButton.SetDontClick (BYVAL bState AS BOOLEAN)
   SendMessageW(m_hCtl, BM_SETDONTCLICK, bState, 0)
END SUB
' ========================================================================================
#endif

' ========================================================================================
' Retrieves a handle to the image (icon or bitmap) associated with the button.
' ========================================================================================
PRIVATE FUNCTION CButton.GetImage (BYVAL imageType aS LONG) AS HANDLE
   RETURN CAST(HANDLE, SendMessageW(m_hCtl, BM_GETIMAGE, imageType, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves a handle to the bitmap associated with the button.
' ========================================================================================
PRIVATE FUNCTION CButton.GetBitmap () AS HANDLE
   RETURN CAST(HANDLE, SendMessageW(m_hCtl, BM_GETIMAGE, IMAGE_BITMAP, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves a handle to the icon associated with the button.
' ========================================================================================
PRIVATE FUNCTION CButton.GetIcon () AS HANDLE
   RETURN CAST(HANDLE, SendMessageW(m_hCtl, BM_GETIMAGE, IMAGE_ICON, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Associates a new bitmap with a button. The return value is a handle to the image previously
' associated with the button, if any; otherwise, it is NULL.
' ========================================================================================
PRIVATE FUNCTION CButton.SetBitmap (BYVAL hbmp AS HBITMAP) AS HBITMAP
   RETURN cast(HBITMAP, SendMessageW(m_hCtl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hbmp)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Associates a new icon with the button. The return value is a handle to the image previously
' associated with the button, if any; otherwise, it is NULL.
' ========================================================================================
PRIVATE FUNCTION CButton.SetIcon (BYVAL hIcon AS HICON) AS HICON
   RETURN cast(..HICON, SendMessageW(m_hCtl, BM_SETIMAGE, IMAGE_ICON, cast(LPARAM, hIcon)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Associates a new image (icon or bitmap) with the button. The return value is a handle to
' the image previously associated with the button, if any; otherwise, it is NULL.
' ========================================================================================
PRIVATE FUNCTION CButton.SetImage (BYVAL ImageType AS DWORD, BYVAL hImage AS HANDLE) AS HANDLE
   RETURN cast(HANDLE, SendMessageW(m_hCtl, BM_SETIMAGE, ImageType, cast(LPARAM, hImage)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Deletes a bitmap associated with a button. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CButton.DeleteBitmap () AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, NULL))))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Deletes an icon associated with a button. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CButton.DeleteIcon () AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, BM_SETIMAGE, IMAGE_ICON, cast(LPARAM, NULL))))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Deletes an image (icon or bitmap) associated a button. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CButton.DeleteImage (BYVAL ImageType AS DWORD) AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, BM_SETIMAGE, ImageType, cast(LPARAM, NULL))))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the check state of a radio button or check box. 
' ========================================================================================
PRIVATE FUNCTION CButton.GetCheck () AS DWORD
   RETURN SendMessageW(m_hCtl, BM_GETCHECK, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the check state of a radio button or check box.
' ========================================================================================
PRIVATE SUB CButton.SetCheck (BYVAL checkState AS LONG)
   SendMessageW(m_hCtl, BM_SETCHECK, checkState, 0)
END SUB
' ========================================================================================
' ========================================================================================
' Sets the button state to checked.
' ========================================================================================
PRIVATE SUB CButton.Check ()
   SendMessageW m_hCtl, BM_SETCHECK, BST_CHECKED, 0
END SUB
' ========================================================================================
' ========================================================================================
' Sets the button state to cleared.
' ========================================================================================
PRIVATE SUB CButton.Uncheck ()
   SendMessageW m_hCtl, BM_SETCHECK, BST_UNCHECKED, 0
END SUB
' ========================================================================================
' ========================================================================================
' Sets the button state to grayed, indicating an indeterminate state. Use this value only
' if the button has the BS_3STATE or BS_AUTO3STATE style.
' ========================================================================================
PRIVATE SUB CButton.Gray ()
   SendMessageW m_hCtl, BM_SETCHECK, BST_INDETERMINATE, 0
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the state of a button or check box.
' ========================================================================================
PRIVATE FUNCTION CButton.GetState () AS DWORD
   RETURN SendMessageW(m_hCtl, BM_GETSTATE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the highlight state of a button. The highlight state indicates whether the button
' is highlighted as if the user had pushed it. 
' ========================================================================================
PRIVATE SUB CButton.SetState (BYVAL bstate AS BOOLEAN)
   SendMessageW(m_hCtl, BM_SETSTATE, bstate, 0)
END SUB
' ========================================================================================
' ========================================================================================
' Sets the style of a button.
' ========================================================================================
PRIVATE SUB CButton.SetStyle (BYVAL dwStyle AS DWORD, BYVAL bRedraw AS BOOLEAN)
   SendMessageW(m_hCtl, BM_SETSTYLE, dwStyle, bRedraw)
END SUB
' ========================================================================================
' ========================================================================================
' Retrieves the style of button.
' ========================================================================================
PRIVATE FUNCTION CButton.GetStyle () AS DWORD
   RETURN GetWindowLongPtrW(m_hCtl, GWL_STYLE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the size of the button that best fits its text and image, if an image list is present.
' ========================================================================================
PRIVATE FUNCTION CButton.GetIdealSize (BYREF tSize AS SIZE) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_GETIDEALSIZE, 0, CAST(LPARAM, @tSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns an image list to a button control.
' Parameters:
' - hImageList = The Handle to the image list
' - nLeft      = The x-coordinate of the upper-left corner of the rectangle for the image.
' - nTop       = The y-coordinate of the upper-left corner of the rectangle for the image.
' - nRight     = The x-coordinate of the lower-right corner of the rectangle for the image.
' - nBottom    = The y-coordinate of the lower-right corner of the rectangle for the image.
' - uAlign     = The alignment to use. This parameter can be one of the following values:
'                BUTTON_IMAGELIST_ALIGN_LEFT   Align the image with the left margin.
'                BUTTON_IMAGELIST_ALIGN_RIGHT  Align the image with the right margin
'                BUTTON_IMAGELIST_ALIGN_TOP    Align the image with the top margin
'                BUTTON_IMAGELIST_ALIGN_BOTTOM Align the image with the bottom margin
'                BUTTON_IMAGELIST_ALIGN_CENTER Center the image.
'                The default value is BUTTON_IMAGELIST_ALIGN_LEFT.
' Return value:  If the function succeeds, it returns TRUE. Otherwise it returns FALSE.
' Note:  To use this function, you must provide a manifest specifying Comclt32.dll version 6.0.
' ========================================================================================
PRIVATE FUNCTION CButton.SetImageList (BYVAL himl AS HIMAGELIST, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG, BYVAL uALign AS DWORD = 0) AS BOOLEAN
   DIM bi AS BUTTON_IMAGELIST = (himl, (nLeft, nTop, nRight, nBottom), uAlign)
   RETURN SendMessageW(m_hCtl, BCM_SETIMAGELIST, 0, cast(LPARAM, @bi))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns an image list to a button control.
' ========================================================================================
PRIVATE FUNCTION CButton.SetImageList (BYVAL pbuttonImagelist AS BUTTON_IMAGELIST PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_SETIMAGELIST, 0, CAST(LPARAM, pbuttonImagelist))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the BUTTON_IMAGELIST structure that describes the image list assigned to a button control. 
' ========================================================================================
PRIVATE FUNCTION CButton.GetImageList () AS BUTTON_IMAGELIST PTR
   DIM pbuttonImagelist AS BUTTON_IMAGELIST PTR
   SendMessageW(m_hCtl, BCM_GETIMAGELIST, 0, CAST(LPARAM, @pbuttonImagelist))
   RETURN pbuttonImagelist
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the margins for drawing text in a button control.
' ========================================================================================
PRIVATE FUNCTION CButton.SetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_SETTEXTMARGIN, 0, CAST(LPARAM, @tmargin))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the margins used to draw text in a button control.
' ========================================================================================
PRIVATE FUNCTION CButton.GetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_GETTEXTMARGIN, 0, CAST(LPARAM, @tMargin))
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT >= &h0600
 ' ========================================================================================
' Sets the drop down state for a specified button with style of BS_SPLITBUTTON.
' ========================================================================================
PRIVATE FUNCTION CButton.SetDropDownState (BYVAL bDropDown AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_SETDROPDOWNSTATE, bDropDown, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets information for a specified split button control.
' ========================================================================================
PRIVATE FUNCTION CButton.SetSplitInfo (BYREF info AS BUTTON_SPLITINFO) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_SETSPLITINFO, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets information for a specified split button control.
' ========================================================================================
PRIVATE FUNCTION CButton.GetSplitInfo (BYREF info AS BUTTON_SPLITINFO) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, BCM_GETSPLITINFO, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the length of the note text that may be displayed in the description for a command link.
' ========================================================================================
PRIVATE FUNCTION CButton.GetNoteLength () AS LONG
   RETURN SendMessageW(m_hCtl, BCM_GETNOTELENGTH, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the text of the note associated with the Command Link button.
' ========================================================================================
PRIVATE FUNCTION CButton.GetNote () AS DWSTRING
   DIM nLen AS LONG = SendMessageW(m_hCtl, BCM_GETNOTELENGTH, 0, 0)
   DIM dws AS DWSTRING = WSPACE(nLen + 1)
   SendMessageW(m_hCtl, BCM_GETNOTE, nLen + 1, CAST(LPARAM, *dws))
   SetResult(GetLastError)
   RETURN dws
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the text of the note associated with a specified command link button.
' ========================================================================================
PRIVATE FUNCTION CButton.SetNote (BYREF pwsz AS WSTRING) AS BOOLEAN
   RETURN SendMessage(m_hCtl, BCM_SETNOTE, 0, CAST(LPARAM, @pwsz))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the elevation required state for a specified button or command link to display an
' elevated icon.
' ========================================================================================
PRIVATE FUNCTION CButton.SetElevationRequiredState (BYVAL bRequired AS BOOLEAN) AS LONG
   RETURN SendMessageW(m_hCtl, BCM_SETSHIELD, 0, bRequired)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Changes the check state of a button control.
' ========================================================================================
PRIVATE FUNCTION CButton.CheckDlgButton (BYVAL cIDButton AS LONG, BYVAL uCheck AS UINT) AS BOOLEAN
   RETURN ..CheckDlgButton(GetParent(m_hCtl), cIDButton, uCheck)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Adds a check mark to (checks) a specified radio button in a group and removes a check mark
' from (clears) all other radio buttons in the group.
' ========================================================================================
PRIVATE FUNCTION CButton.CheckRadioButton (BYVAL cIDFirstButton AS LONG, BYVAL cIDLastButton AS LONG, BYVAL cIDCheckButton AS LONG) AS BOOLEAN
   RETURN ..CheckRadioButton(GetParent(m_hCtl), cIDFirstButton, cIDLastButton, cIDCheckButton)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Determines whether a button control is checked or whether a three-state button control is
' checked, unchecked, or indeterminate.
' ========================================================================================
PRIVATE FUNCTION CButton.IsDlgButtonChecked (BYVAL cIDButton AS LONG) AS UINT
   RETURN ..IsDlgButtonChecked(GetParent(m_hCtl), cIDButton)
END FUNCTION
' ========================================================================================

END NAMESPACE
