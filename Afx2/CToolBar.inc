' ########################################################################################
' Platform: Microsoft Windows
' Filename: CToolBar.inc
' Purpose:  Windows common controls - ToolBar
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
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CToolBar class
' ########################################################################################

TYPE CToolBar EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION AddBitmap (BYVAL nButtons AS DWORD, BYVAL tabm AS TBADDBITMAP PTR) AS LONG
   DECLARE FUNCTION AddButtons (BYVAL nButtons AS DWORD, BYVAL lpButtons AS TBBUTTON PTR) AS BOOLEAN
   DECLARE FUNCTION AddButton (BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
                    BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION AddSeparator (BYVAL nWidth AS LONG = 0) AS BOOLEAN
   DECLARE FUNCTION InsertButton OVERLOAD (BYVAL idxButton AS LONG, BYVAL lpButton AS TBBUTTON PTR) AS BOOLEAN
   DECLARE FUNCTION InsertButton OVERLOAD (BYVAL idxButton AS LONG, BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
                    BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION InsertSeparator (BYVAL idxButton AS LONG, BYVAL nWidth AS LONG = 0) AS BOOLEAN
   DECLARE FUNCTION AddString (BYVAL hInst AS HINSTANCE, BYVAL idString AS INT_PTR) AS LONG
   DECLARE SUB AutoSize ()
   DECLARE FUNCTION ButtonCount () AS DWORD
   DECLARE FUNCTION GetCount () AS DWORD   '// For PowerBasic's compatibility
   DECLARE SUB ButtonStructSize (BYVAL cb AS LONG)
   DECLARE FUNCTION ChangeBitmap (BYVAL idButton AS LONG, BYVAL iBitmap AS WORD) AS BOOLEAN
   DECLARE FUNCTION CheckButton (BYVAL idButton AS LONG, BYVAL fCheck AS SHORT) AS BOOLEAN
   DECLARE FUNCTION CommandToIndex (BYVAL idButton AS LONG) AS LONG
   DECLARE SUB Customize ()
   DECLARE FUNCTION DeleteButton (BYVAL nButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION EnableButton (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION DisableButton (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetAnchorHighlight () AS BOOLEAN
   DECLARE FUNCTION GetBitmap (BYVAL idButton AS LONG) AS LONG
   DECLARE FUNCTION GetBitmapFlags () AS DWORD
   DECLARE FUNCTION GetButton (BYVAL nButton AS LONG, BYVAL ptbb AS TBBUTTON PTR) AS BOOLEAN
   DECLARE FUNCTION GetButtonHeight () AS DWORD
   DECLARE FUNCTION GetButtonWidth () AS DWORD
   DECLARE FUNCTION GetButtonSize () AS DWORD
   DECLARE FUNCTION GetButtonInfo (BYVAL iID AS LONG, BYVAL lptbbi AS TBBUTTONINFOW PTR) AS LONG
   DECLARE FUNCTION GetColorScheme (BYVAL lpcs AS COLORSCHEME PTR) AS BOOLEAN
   DECLARE FUNCTION GetImageList () AS HIMAGELIST
   DECLARE FUNCTION GetDisabledImageList () AS HIMAGELIST
   DECLARE FUNCTION GetHotImageList () AS HIMAGELIST
   DECLARE FUNCTION GetExtendedStyle () AS DWORD
   DECLARE FUNCTION GetHotItem () AS DWORD
   DECLARE FUNCTION GetInsertMark (BYVAL ptbim AS TBINSERTMARK PTR) AS LONG
   DECLARE FUNCTION GetInsertMarkColor () AS COLORREF
   DECLARE FUNCTION GetItemRect (BYVAL nButton AS LONG, BYVAL lprc AS RECT PTR) AS BOOLEAN
   DECLARE FUNCTION GetMaxSize (BYVAL psize AS SIZE PTR) AS BOOLEAN
   DECLARE SUB GetMetrics (BYVAL ptbMetrics AS TBMETRICS PTR)
   DECLARE FUNCTION GetPadding () AS DWORD
   DECLARE FUNCTION GetHorizontalPadding () AS DWORD
   DECLARE FUNCTION GetVerticalPadding () AS DWORD
   DECLARE FUNCTION GetRect (BYVAL iID AS LONG, BYVAL lprc AS RECT PTR) AS BOOLEAN
   DECLARE FUNCTION GetRows () AS DWORD
   DECLARE FUNCTION GetState (BYVAL idButton AS LONG) AS DWORD
   DECLARE FUNCTION GetString (BYVAL iString AS SHORT, BYVAL pBuffer AS WSTRING PTR, BYVAL cchMax AS SHORT) AS LONG
   DECLARE FUNCTION GetStyle () AS DWORD
   DECLARE FUNCTION GetTextRows () AS DWORD
   DECLARE FUNCTION GetTooltips () AS HWND
   DECLARE FUNCTION HideButton (BYVAL idButton AS LONG, BYVAL fShow AS SHORT) AS BOOLEAN
   DECLARE FUNCTION HitTest (BYVAL pptHitTest AS POINT PTR) AS LONG
   DECLARE FUNCTION Indeterminate (BYVAL idButton AS LONG, BYVAL fIndeterminate AS SHORT) AS BOOLEAN
   DECLARE FUNCTION InsertMarkHitTest (BYVAL lppt AS POINT PTR, BYVAL lptbim AS TBINSERTMARK PTR) AS BOOLEAN
   DECLARE FUNCTION IsButtonChecked (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsButtonEnabled (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsButtonHidden (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsButtonHighlighted (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsButtonIndeterminate (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION IsButtonPressed (BYVAL idButton AS LONG) AS BOOLEAN
   DECLARE FUNCTION LoadImages (BYVAL iBitmapID AS DWORD, BYVAL hinst AS HINSTANCE) AS LONG
   DECLARE FUNCTION MapAccelerator (BYVAL chAccel AS DWORD, BYVAL pIDBtn AS DWORD PTR) AS BOOLEAN
   DECLARE FUNCTION MarkButton (BYVAL idButton AS LONG, BYVAL fHighlight AS SHORT) AS BOOLEAN
   DECLARE FUNCTION MoveButton (BYVAL uOldPos AS DWORD, BYVAL uNewPos AS DWORD) AS BOOLEAN
   DECLARE FUNCTION PressButton (BYVAL idButton AS LONG, BYVAL fPress AS SHORT) AS BOOLEAN
   DECLARE FUNCTION ReplaceBitmap (BYVAL ptbrb AS TBREPLACEBITMAP PTR) AS BOOLEAN
   DECLARE SUB SaveRestore (BYVAL fSave AS LONG, BYVAL ptbsp AS TBSAVEPARAMSW PTR)
   DECLARE FUNCTION SetAnchorHighlight (BYVAL fAnchor AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION SetBitmapSize (BYVAL dxBitmap AS SHORT, BYVAL dyBitmap AS SHORT) AS BOOLEAN
   DECLARE FUNCTION SetButtonInfo (BYVAL iID AS LONG, BYVAL lptbbi AS TBBUTTONINFOW PTR) AS BOOLEAN
   DECLARE FUNCTION SetButtonSize (BYVAL dxButton AS SHORT, BYVAL dyButton AS SHORT) AS BOOLEAN
   DECLARE FUNCTION SetButtonWidth (BYVAL cxMin AS SHORT, BYVAL cxMax AS SHORT) AS BOOLEAN
   DECLARE FUNCTION SetCmdId (BYVAL index AS SHORT, BYVAL cmdid AS SHORT) AS BOOLEAN
   DECLARE SUB SetColorScheme (BYVAL lpcs AS COLORSCHEME PTR)
   DECLARE FUNCTION SetImageList (BYVAL himlNew AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetHotImageList (BYVAL himlNewHot AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetDisabledImageList (BYVAL himlNewDisabled AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetDrawTextFlags (BYVAL dwMask AS DWORD, BYVAL dwDTFlags AS DWORD) AS DWORD
   DECLARE FUNCTION SetExtendedStyle (BYVAL dwExStyle AS DWORD) AS DWORD
   DECLARE FUNCTION SetHotItem (BYVAL iHot AS LONG) AS LONG
   DECLARE FUNCTION SetIndent (BYVAL iIndent AS LONG) AS BOOLEAN
   DECLARE SUB SetInsertMark (BYVAL lptbim AS TBINSERTMARK PTR)
   DECLARE FUNCTION SetInsertMarkColor (BYVAL clrInsertMark AS COLORREF) AS COLORREF
   DECLARE SUB SetListGap (BYVAL gap AS DWORD)
   DECLARE FUNCTION SetMaxTextRows (BYVAL iMaxRows AS LONG) AS BOOLEAN
   DECLARE SUB SetMetrics (BYVAL ptbMetrics AS TBMETRICS PTR)
   DECLARE FUNCTION SetPadding (BYVAL cx AS SHORT, BYVAL cy AS SHORT) AS DWORD
   DECLARE FUNCTION SetParent (BYVAL hwndParent AS HWND) AS HWND
   DECLARE FUNCTION SetPressedImageList (BYVAL iImageID AS DWORD, BYVAL himlNew AS HIMAGELIST) AS HIMAGELIST
   DECLARE SUB SetRows (BYVAL cRows AS WORD, BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR)
   DECLARE FUNCTION SetState (BYVAL idButton AS LONG, BYVAL fState AS SHORT) AS BOOLEAN
   DECLARE SUB SetStyle (BYVAL dwStyle AS DWORD)
   DECLARE SUB SetToolTips (BYVAL hwndToolTip AS HWND)
   DECLARE SUB SetWindowTheme (BYVAL pwszTheme AS WSTRING PTR)
   DECLARE SUB MakeOpaque ()
   DECLARE FUNCTION GetObject () AS IDropTarget PTR
   DECLARE FUNCTION GetButtonText (BYVAL idButton AS LONG, BYVAL pBuffer AS WSTRING PTR) AS LONG
   DECLARE FUNCTION GetButtonTextLength (BYVAL idButton AS LONG) AS LONG
   DECLARE FUNCTION ButtonGetTextLength (BYVAL idButton AS LONG) AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' * CToolBar constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CToolBar (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CToolBar (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CToolBar (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CToolBar (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CToolBar destructor
' ========================================================================================
PRIVATE DESTRUCTOR CToolBar
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                    *** TOOLBAR ***
' ########################################################################################

' ========================================================================================
' Adds one or more images to the list of button images available for a toolbar.
' Returns the index of the first new image if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.AddBitmap (BYVAL nButtons AS DWORD, BYVAL tabm AS TBADDBITMAP PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_ADDBITMAP, cast(WPARAM, nButtons), CAST(LPARAM, tabm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds one or more buttons to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.AddButtons (BYVAL nButtons AS DWORD, BYVAL lpButtons AS TBBUTTON PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ADDBUTTONSW, cast(WPARAM, nButtons), CAST(LPARAM, lpButtons))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a button to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.AddButton (BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
   IF fsState = 0 THEN fsState = TBSTATE_ENABLED
   DIM idxString AS INT_PTR
   IF pwszText <> NULL THEN idxString = IIF(LEN(*pwszText) = 0, -1, CAST(INT_PTR, pwszText))
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0, 0, 0, 0, 0}, dwData, idxString)
#else
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0}, dwData, idxString)
#endif
   FUNCTION = SendMessageW(m_hCtl, TB_ADDBUTTONSW, 1, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a separator to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.AddSeparator (BYVAL nWidth AS LONG = 0) AS BOOLEAN
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0, 0, 0, 0, 0}, 0, -1)
#else
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0}, 0, -1)
#endif
   FUNCTION = SendMessageW(m_hCtl, TB_ADDBUTTONSW, 1, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.InsertButton OVERLOAD (BYVAL idxButton AS LONG, BYVAL lpButton AS TBBUTTON PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_INSERTBUTTONW, cast(WPARAM, idxButton), CAST(LPARAM, lpButton))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CToolBar.InsertButton OVERLOAD (BYVAL idxButton AS LONG, BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
   IF fsState = 0 THEN fsState = TBSTATE_ENABLED
   DIM idxString AS INT_PTR
   IF pwszText <> NULL THEN idxString = IIF(LEN(*pwszText) = 0, -1, CAST(INT_PTR, pwszText))
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0, 0, 0, 0, 0}, dwData, idxString)
#else
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0}, dwData, idxString)
#endif
   FUNCTION = SendMessageW(m_hCtl, TB_INSERTBUTTONW, idxButton, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a separator in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.InsertSeparator (BYVAL idxButton AS LONG, BYVAL nWidth AS LONG = 0) AS BOOLEAN
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0, 0, 0, 0, 0}, 0, -1)
#else
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0}, 0, -1)
#endif
   FUNCTION = SendMessageW(m_hCtl, TB_INSERTBUTTONW, idxButton, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a new string to the toolbar's string pool.
' Returns the index of the first new string if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.AddString (BYVAL hInst AS HINSTANCE, BYVAL idString AS INT_PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_ADDSTRINGW, CAST(WPARAM, hInst), IIF(hInst = NULL, idString, MAKELONG(idString, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes a toolbar to be resized.
' ========================================================================================
PRIVATE SUB CToolBar.AutoSize ()
   SendMessageW(m_hCtl, TB_AUTOSIZE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves a count of the buttons currently in the toolbar.
' ========================================================================================
PRIVATE FUNCTION CToolBar.ButtonCount () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_BUTTONCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' PowerBasic uses the less explicit GetCount.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetCount () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_BUTTONCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the size of the TBBUTTON structure.
' ========================================================================================
PRIVATE SUB CToolBar.ButtonStructSize (BYVAL cb AS LONG)
   SendMessageW(m_hCtl, TB_BUTTONSTRUCTSIZE, cb, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Changes the bitmap for a button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.ChangeBitmap (BYVAL idButton AS LONG, BYVAL iBitmap AS WORD) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_CHANGEBITMAP, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(iBitmap, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Checks or unchecks a given button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.CheckButton (BYVAL idButton AS LONG, BYVAL fCheck AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_CHECKBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(fCheck, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the zero-based index for the button associated with the specified command
' identifier. Returns the zero-based index for the button or -1 if the specified command
' identifier is invalid.
' ========================================================================================
PRIVATE FUNCTION CToolBar.CommandToIndex (BYVAL idButton AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_COMMANDTOINDEX, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays the Customize Toolbar dialog box.
' ========================================================================================
PRIVATE SUB CToolBar.Customize ()
   SendMessageW(m_hCtl, TB_CUSTOMIZE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Deletes a button from the toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.DeleteButton (BYVAL nButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_DELETEBUTTON, cast(WPARAM, nButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.EnableButton (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ENABLEBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(CTRUE, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Disables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.DisableButton (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ENABLEBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(FALSE, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the anchor highlight setting for a toolbar.
' Returns a Boolean value that indicates if anchor highlighting is set. If this value is
' nonzero, anchor highlighting is enabled. If this value is zero, it is disabled.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetAnchorHighlight () AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_GETANCHORHIGHLIGHT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the index of the bitmap associated with a button in a toolbar.
' Returns the index of the bitmap if successful, or zero otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetBitmap (BYVAL idButton AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_GETBITMAP, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the flags that describe the type of bitmap to be used.
' Returns a DWORD value that describes the type of bitmap that should be used. If this
' return value has the TBBF_LARGE flag set, applications should use large bitmaps (24 x 24);
' otherwise, applications should use small bitmaps (16 x 16). All other bits are reserved.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetBitmapFlags () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETBITMAPFLAGS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetButton (BYVAL nButton AS LONG, BYVAL ptbb AS TBBUTTON PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_GETBUTTON, CAST(WPARAM, nButton), CAST(LPARAM, ptbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current width and height of toolbar buttons, in pixels.
' Returns a DWORD value that contains the width and height values in the low word and high
' word, respectively.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetButtonSize () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETBUTTONSIZE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current height of toolbar buttons, in pixels.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetButtonHeight () AS DWORD
   FUNCTION = HIWORD(SendMessageW(m_hCtl, TB_GETBUTTONSIZE, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current width of toolbar buttons, in pixels.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetButtonWidth () AS DWORD
   FUNCTION = LOWORD(SendMessageW(m_hCtl, TB_GETBUTTONSIZE, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves extended information for a button in a toolbar.
' Returns the zero-based index of the button, or -1 if an error occurs.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetButtonInfo (BYVAL iID AS LONG, BYVAL lptbbi AS TBBUTTONINFOW PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_GETBUTTONINFOW, cast(WPARAM, iID), cast(LPARAM, lptbbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the color scheme information from the toolbar control.
' - lpcs: Pointer to a COLORSCHEME structure that will receive the color scheme information.
'   You must set the cbSize member of this structure to sizeof(COLORSCHEME) before sending
'   this message.
' Returns nonzero if successful, or zero otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetColorScheme (BYVAL lpcs AS COLORSCHEME PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_GETCOLORSCHEME, 0, CAST(LPARAM, lpcs))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the image list that a toolbar control uses to display buttons in their default
' state. A toolbar control uses this image list to display buttons when they are not hot
' or disabled.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetImageList () AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_GETIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the image list that a toolbar control uses to display inactive buttons.
' A toolbar control uses this image list to display buttons when they are disabled.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetDisabledImageList () AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_GETDISABLEDIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the image list that a toolbar control uses to display hot buttons.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetHotImageList () AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_GETHOTIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the extended styles for a toolbar control.
' Returns a DWORD that represents the styles currently in use for the toolbar control.
' This value can be a combination of extended styles.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetExtendedStyle () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the horizontal padding for a toolbar control.
' Minimum operating systems Windows 2000, Windows NT 4.0 with Internet Explorer 4.0,
' Windows 98, Windows 95 with Internet Explorer 4.0
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetHorizontalPadding () AS DWORD
   FUNCTION = LOWORD(SendMessageW(m_hCtl, TB_GETPADDING, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the vertical padding for a toolbar control.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetVerticalPadding () AS DWORD
   FUNCTION = HIWORD(SendMessageW(m_hCtl, TB_GETPADDING, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the index of the hot item in a toolbar.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetHotItem () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETHOTITEM, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current insertion mark for the toolbar.
' Always returns TRUE.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetInsertMark (BYVAL ptbim AS TBINSERTMARK PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_GETINSERTMARK, 0, CAST(LPARAM, ptbim))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the color used to draw the insertion mark for the toolbar.
' Returns a COLORREF value that contains the current insertion mark color.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetInsertMarkColor () AS COLORREF
   FUNCTION = SendMessageW(m_hCtl, TB_GETINSERTMARKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle of a button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetItemRect (BYVAL nButton AS LONG, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_GETITEMRECT, cast(WPARAM, nButton), CAST(LPARAM, lprc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the total size of all of the visible buttons and separators in the toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetMaxSize (BYVAL psize AS SIZE PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_GETMAXSIZE, 0, CAST(LPARAM, psize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the metrics of a toolbar control.
' Note: To use this API, you must provide a manifest specifying Comclt32.dll version 6.0.
' ========================================================================================
PRIVATE SUB CToolBar.GetMetrics (BYVAL ptbMetrics AS TBMETRICS PTR)
   SendMessageW(m_hCtl, TB_GETMETRICS, 0, CAST(LPARAM, ptbMetrics))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the padding for a toolbar control.
' Returns a DWORD value that contains the horizontal padding in the low word and the
' vertical padding in the high word, in pixels.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetPadding () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETPADDING, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle for a specified toolbar button.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetRect (BYVAL iID AS LONG, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_GETRECT, cast(WPARAM, iID), CAST(LPARAM, lprc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of rows of buttons in a toolbar with the TBSTYLE_WRAPABLE style.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetRows () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETROWS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the state of the specified button in a toolbar, such as
' whether it is enabled, pressed, or checked.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetState (BYVAL idButton AS LONG) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETSTATE, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a string from a toolbar's string pool.
' - iString: Index of the string
' - pBuffer: Pointer to a buffer used to return the string.
' - cchMax: Length of the buffer in bytes.
' Returns the string length if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetString (BYVAL iString AS SHORT, BYVAL pBuffer AS WSTRING PTR, BYVAL cchMax AS SHORT) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_GETSTRINGW, MAKELONG(cchMax, iString), CAST(LPARAM, pbuffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the styles currently in use for a toolbar control.
' Returns a DWORD value that is a combination of toolbar control styles.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetStyle () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of rows of buttons in a toolbar with the TBSTYLE_WRAPABLE style.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetTextRows () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_GETTEXTROWS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the ToolTip control, if any, associated with the toolbar.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetTooltips () AS HWND
   FUNCTION = cast(HWND, SendMessageW(m_hCtl, TB_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides or shows the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.HideButton (BYVAL idButton AS LONG, BYVAL fShow AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_HIDEBUTTON, idButton, MAKELONG(fShow, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines where a point lies in a toolbar control.
' Returns an integer value. If the return value is zero or a positive value, it is the
' zero-based index of the nonseparator item in which the point lies. If the return value
' is negative, the point does not lie within a button. The absolute value of the return
' value is the index of a separator item or the nearest nonseparator item.
' ========================================================================================
PRIVATE FUNCTION CToolBar.HitTest (BYVAL pptHitTest AS POINT PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_HITTEST, 0, CAST(LPARAM, pptHitTest))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets or clears the indeterminate state of the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.Indeterminate (BYVAL idButton AS LONG, BYVAL fIndeterminate AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_INDETERMINATE, cast(WPARAM, idButton), MAKELONG(fIndeterminate, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the insertion mark information for a point in a toolbar.
' Minimum operating systems Windows 2000, Windows NT 4.0 with Internet Explorer 4.0,
' Windows 98, Windows 95 with Internet Explorer 4.0
' ========================================================================================
PRIVATE FUNCTION CToolBar.InsertMarkHitTest (BYVAL lppt AS POINT PTR, BYVAL lptbim AS TBINSERTMARK PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_INSERTMARKHITTEST, CAST(WPARAM, lppt), CAST(LPARAM, lptbim))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified button in a toolbar is checked.
' Returns TRUE if the point is an insertion mark, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.IsButtonChecked (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ISBUTTONCHECKED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified button in a toolbar is enabled.
' Returns TRUE if the button is enabled, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.IsButtonEnabled (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ISBUTTONENABLED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified button in a toolbar is hidden.
' Returns TRUE if the button is hidden, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.IsButtonHidden (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ISBUTTONHIDDEN, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Checks the highlight state of a toolbar button.
' Returns TRUE if the button is highlighted, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.IsButtonHighlighted (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ISBUTTONHIGHLIGHTED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified button in a toolbar is indeterminate.
' Returns TRUE if the button is in a indetermides state, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.IsButtonIndeterminate (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ISBUTTONINDETERMINATE, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified button in a toolbar is pressed.
' Returns TRUE if the button is pressed, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.IsButtonPressed (BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_ISBUTTONPRESSED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads system-defined button images into a toolbar control's image list.
' The count of images in the image list. Returns zero if the toolbar has no image list or
' if the existing image list is empty.
' ========================================================================================
PRIVATE FUNCTION CToolBar.LoadImages (BYVAL iBitmapID AS DWORD, BYVAL hinst AS HINSTANCE) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_LOADIMAGES, cast(WPARAM, iBitmapID), CAST(LPARAM, hinst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the ID of the button that corresponds to the specified accelerator character.
' Returns TRUE if one of the buttons has wParam as its accelerator character, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.MapAccelerator (BYVAL chAccel AS DWORD, BYVAL pIDBtn AS DWORD PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_MAPACCELERATOR, cast(WPARAM, chAccel), CAST(LPARAM, pIDBtn))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the highlight state of a given button in a toolbar control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.MarkButton (BYVAL idButton AS LONG, BYVAL fHighlight AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_MARKBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(fHighlight, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves a button from one index to another.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.MoveButton (BYVAL uOldPos AS DWORD, BYVAL uNewPos AS DWORD) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_MOVEBUTTON, cast(WPARAM, uOldPos), cast(LPARAM, uNewPos))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Presses or releases the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.PressButton (BYVAL idButton AS LONG, BYVAL fPress AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_PRESSBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(fPress, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces an existing bitmap with a new bitmap.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.ReplaceBitmap (BYVAL ptbrb AS TBREPLACEBITMAP PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_REPLACEBITMAP, 0, CAST(LPARAM, ptbrb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Send this message to initiate saving or restoring a toolbar state.
' Minimum operating systems Windows NT 3.51, Windows 95
' ========================================================================================
PRIVATE SUB CToolBar.SaveRestore (BYVAL fSave AS LONG, BYVAL ptbsp AS TBSAVEPARAMSW PTR)
   SendMessageW(m_hCtl, TB_SAVERESTOREW, cast(WPARAM, fSave), CAST(LPARAM, ptbsp))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the anchor highlight setting for a toolbar.
' Returns the previous anchor highlight setting. If this value is nonzero, anchor highlighting
' was enabled. If this value is zero, anchor highlighting was disabled.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetAnchorHighlight (BYVAL fAnchor AS BOOLEAN) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETANCHORHIGHLIGHT, fAnchor, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size of the bitmapped images to be added to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetBitmapSize (BYVAL dxBitmap AS SHORT, BYVAL dyBitmap AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETBITMAPSIZE, 0, MAKELONG(dxBitmap, dyBitmap))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the information for an existing button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetButtonInfo (BYVAL iID AS LONG, BYVAL lptbbi AS TBBUTTONINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETBUTTONINFOW, cast(WPARAM, iID), CAST(LPARAM, lptbbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size of the buttons to be added to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetButtonSize (BYVAL dxButton AS SHORT, BYVAL dyButton AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETBUTTONSIZE, 0, cast(LPARAM, MAKELONG(dxButton, dyButton)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum and maximum button widths in the toolbar control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetButtonWidth (BYVAL cxMin AS SHORT, BYVAL cxMax AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETBUTTONWIDTH, 0, cast(LPARAM, MAKELONG(cxMin, cxMax)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the command identifier of a toolbar button.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetCmdId (BYVAL index AS SHORT, BYVAL cmdid AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETCMDID, cast(WPARAM, index), cast(LPARAM, cmdid))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the color scheme information for the toolbar control.
' ========================================================================================
PRIVATE SUB CToolBar.SetColorScheme (BYVAL lpcs AS COLORSCHEME PTR)
   SendMessageW(m_hCtl, TB_SETCOLORSCHEME, 0, CAST(LPARAM, lpcs))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the image list that the toolbar will use to display buttons that are in their
' default state. Returns the handle to the image list previously used to display buttons
' in their default state, or NULL if no image list was previously set.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetImageList (BYVAL himlNew AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_SETIMAGELIST, 0, CAST(LPARAM, himlNew)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the image list that the toolbar control will use to display hot buttons.
' Returns the handle to the image list previously used to display hot buttons, or NULL if
' no image list was previously set.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetHotImageList (BYVAL himlNewHot AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_SETHOTIMAGELIST, 0, CAST(LPARAM, himlNewHot)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the image list that the toolbar control will use to display disabled buttons.
' Returns the handle to the image list previously used to display disabled buttons, or NULL
' if no image list was previously set.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetDisabledImageList (BYVAL himlNewDisabled AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_SETDISABLEDIMAGELIST, 0, CAST(LPARAM, himlNewDisabled)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the image list that the toolbar uses to display buttons that are in a pressed state.
' Returns the handle to the image list previously used to display buttons in their pressed
' state, or NULL if no such image list was previously set.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetPressedImageList (BYVAL iImageID AS DWORD, BYVAL himlNew AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(m_hCtl, TB_SETPRESSEDIMAGELIST, cast(WPARAM, iImageID), CAST(LPARAM, himlNew)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text drawing flags for the toolbar.
' Returns the previous text drawing flags.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetDrawTextFlags (BYVAL dwMask AS DWORD, BYVAL dwDTFlags AS DWORD) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_SETDRAWTEXTFLAGS, cast(WPARAM, dwMask), cast(LPARAM, dwDTFlags))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the extended styles for a toolbar control.
' Returns a DWORD that represents the previous extended styles. This value can be a
' combination of extended styles.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetExtendedStyle (BYVAL dwExStyle AS DWORD) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_SETEXTENDEDSTYLE, 0, cast(LPARAM, dwExStyle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the hot item in a toolbar.
' iHot: Index of the item that will be made hot. If this value is -1, none of the items will be hot.
' Returns the index of the previous hot item, or -1 if there was no hot item.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetHotItem (BYVAL iHot AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_SETHOTITEM, cast(WPARAM, iHot), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the indentation for the first button in a toolbar control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetIndent (BYVAL iIndent AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETINDENT, cast(WPARAM, iIndent), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current insertion mark for the toolbar.
' ========================================================================================
PRIVATE SUB CToolBar.SetInsertMark (BYVAL lptbim AS TBINSERTMARK PTR)
   SendMessageW(m_hCtl, TB_SETINSERTMARK, 0, CAST(LPARAM, lptbim))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the color used to draw the insertion mark for the toolbar.
' clrInsertMark: COLORREF value that contains the new insertion mark color.
' Returns a COLORREF value that contains the previous insertion mark color.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetInsertMarkColor (BYVAL clrInsertMark AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(m_hCtl, TB_SETINSERTMARKCOLOR, cast(WPARAM, clrInsertMark), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the distance between the toolbar buttons on a specific toolbar.
' ========================================================================================
PRIVATE SUB CToolBar.SetListGap (BYVAL gap AS DWORD)
   SendMessageW(m_hCtl, TB_SETLISTGAP, gap, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the maximum number of text rows displayed on a toolbar button.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetMaxTextRows (BYVAL iMaxRows AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETMAXTEXTROWS, cast(WPARAM, iMaxRows), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the metrics of a toolbar control.
' ptbMetrics: TBMETRICS structure that contains the toolbar metrics to set.
' ========================================================================================
PRIVATE SUB CToolBar.SetMetrics (BYVAL ptbMetrics AS TBMETRICS PTR)
   SendMessageW(m_hCtl, TB_SETMETRICS, 0, CAST(LPARAM, ptbmetrics))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the padding for a toolbar control.
' Returns a DWORD value that contains the previous horizontal padding in the LOWORD and
' the previous vertical padding in the HIWORD, in pixels.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetPadding (BYVAL cx AS SHORT, BYVAL cy AS SHORT) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, TB_SETPADDING, 0, MAKELONG(cx, cy))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the window to which the toolbar control sends notification messages.
' The return value is a handle to the previous notification window, or NULL if there is no
' previous notification window.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetParent (BYVAL hwndParent AS HWND) AS HWND
   FUNCTION = CAST(HWND, SendMessageW(m_hCtl, TB_SETPARENT, CAST(WPARAM, hwndParent), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the number of rows of buttons in a toolbar.
' ========================================================================================
PRIVATE SUB CToolBar.SetRows (BYVAL cRows AS WORD, BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR)
   SendMessageW(m_hCtl, TB_SETROWS, MAKELONG(cRows, fLarger), CAST(LPARAM, lprc))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the state for the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CToolBar.SetState (BYVAL idButton AS LONG, BYVAL fState AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, TB_SETSTATE, cast(WPARAM, idButton), MAKELONG(fState, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the style for a toolbar control.
' ========================================================================================
PRIVATE SUB CToolBar.SetStyle (BYVAL dwStyle AS DWORD)
   SendMessageW(m_hCtl, TB_SETSTYLE, 0, dwStyle)
END SUB
' ========================================================================================

' ========================================================================================
' Associates a ToolTip control with a toolbar.
' ========================================================================================
PRIVATE SUB CToolBar.SetToolTips (BYVAL hwndToolTip AS HWND)
   SendMessageW(m_hCtl, TB_SETTOOLTIPS, CAST(WPARAM, hwndToolTip), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the visual style of a toolbar control.
' Note: To use this API, you must provide a manifest specifying Comclt32.dll version 6.0.
' Minimum Operating System: Windows XP.
' ========================================================================================
PRIVATE SUB CToolBar.SetWindowTheme (BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(m_hCtl, TB_SETWINDOWTHEME, CAST(WPARAM, pwszTheme), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Makes the toolbar opaque by removing the TBSTYLE_TRANSPARENT style.
' ========================================================================================
PRIVATE SUB CToolBar.MakeOpaque ()
   SetWindowLongPtrW(m_hCtl, GWL_STYLE, GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND NOT TBSTYLE_TRANSPARENT)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the IDropTarget for a toolbar control.
' Minimum operating systems Windows 2000, Windows NT 4.0 with Internet Explorer 4.0,
' Windows 98, Windows 95 with Internet Explorer 4.0
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetObject () AS IDropTarget PTR
   DIM ppvObject AS IDropTarget PTR
   DIM IID_IDropTarget AS GUID = (&h00000122, &h0000, &h0000, {&hC0, &h00, &h00, &h00, &h00, &h00, &h00, &h46})
   DIM hr AS HRESULT = SendMessageW(m_hCtl, TB_GETOBJECT, CAST(WPARAM, @IID_IDropTarget), CAST(LPARAM, ppvObject))
   IF hr <> S_OK THEN FUNCTION = ppvObject
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the display text of a button on a toolbar.
' Returns the length, in characters, of the string pointed to by lParam. The length does
' not include the terminating null character. If unsuccessful, the return value is -1.
' idButton = Command identifier of the button whose text is to be retrieved.
' pBuffer = Pointer to a buffer that receives the button text.
' Minimum operating systems Windows NT 3.51, Windows 95
' Security Warning: Using this message incorrectly might compromise the security of your
' program. This message does not provide a way for you to know the size of the buffer.
' If you use this message, first call the message passing NULL in the lParam, this returns
' the number of characters, excluding NULL that are required. Then call the message a
' second time to retrieve the string.
' The returned string corresponds to the text that is currently displayed by the button.
' ========================================================================================
PRIVATE FUNCTION CToolBar.GetButtonText (BYVAL idButton AS LONG, BYVAL pBuffer AS WSTRING PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_GETBUTTONTEXTW, cast(WPARAM, idButton), CAST(LPARAM, pbuffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of the display text of a button on a toolbar.
' Returns the length, in characters, of the string pointed to by lParam. The length does
' not include the terminating null character. If unsuccessful, the return value is -1.
' ========================================================================================
PRIVATE FUNCTION CToolBar.ButtonGetTextLength (BYVAL idButton AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, TB_GETBUTTONTEXTW, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================

END NAMESPACE
