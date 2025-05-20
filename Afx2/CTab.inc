' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTab.inc
' Purpose:  Windows common controls - Tab control
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
' CTab class
' ########################################################################################

TYPE CTab EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE SUB AdjustRect (BYVAL fLarger AS BOOLEAN, BYREF rc AS RECT)
   DECLARE FUNCTION DeleteAllItems () AS BOOLEAN
   DECLARE FUNCTION DeleteItem (BYVAL iItem AS LONG) AS BOOLEAN
   DECLARE SUB DeselectAll (BYVAL fExcludeFocus AS BOOLEAN)
   DECLARE FUNCTION GetCurFocus () AS LONG
   DECLARE FUNCTION GetCurSel () AS LONG
   DECLARE FUNCTION GetExtendedStyle () AS DWORD
   DECLARE FUNCTION GetImageIndex (BYVAL iItem AS DWORD) AS LONG
   DECLARE FUNCTION GetImageList () AS HIMAGELIST
   DECLARE FUNCTION GetItem (BYVAL iItem AS DWORD, BYREF pItem AS TCITEMW) AS BOOLEAN
   DECLARE FUNCTION GetItemCount () AS LONG
   DECLARE FUNCTION GetItemRect (BYVAL iItem AS DWORD, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetItemRect (BYVAL iItem AS DWORD) AS RECT
   DECLARE FUNCTION GetRowCount () AS LONG
   DECLARE FUNCTION GetText (BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetText (BYVAL nTabIndex AS DWORD) AS DWSTRING
   DECLARE FUNCTION GetToolTips () AS HWND
   DECLARE FUNCTION GetUnicodeFormat () AS BOOLEAN
   DECLARE FUNCTION HighlightItem (BYVAL idItem AS LONG, BYVAL fHighlight AS WORD) AS LONG
   DECLARE FUNCTION HitTest (BYREF info AS TC_HITTESTINFO) AS LONG
   DECLARE FUNCTION InsertItem (BYVAL iItem AS DWORD, BYREF item AS TCITEMW) AS LONG
   DECLARE FUNCTION InsertTab (BYVAL nTabIndex AS DWORD, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0) AS LONG
   DECLARE FUNCTION AddTab (BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0) AS LONG
   DECLARE SUB RemoveImage (BYVAL iImage AS LONG)
   DECLARE SUB SetCurFocus (BYVAL iItem AS LONG)
   DECLARE FUNCTION SetCurSel (BYVAL iItem AS LONG) AS LONG
   DECLARE FUNCTION SetExtendedStyle (BYVAL dwExStyle AS DWORD) AS DWORD
   DECLARE FUNCTION SetImageList (BYVAL himl AS HIMAGELIST) AS BOOLEAN
   DECLARE FUNCTION SetImageIndex (BYVAL iItem AS LONG, BYVAL iImage AS LONG) AS BOOLEAN
   DECLARE FUNCTION SetItem (BYVAL iItem AS LONG, BYREF item AS TCITEMW) AS LONG
   DECLARE FUNCTION SetItemExtra (BYVAL cb AS LONG) AS BOOLEAN
   DECLARE FUNCTION SetItemSize (BYVAL x AS LONG, BYVAL y AS LONG) AS DWORD
   DECLARE FUNCTION SetMinTabWidth (BYVAL cx AS LONG) AS LONG
   DECLARE SUB SetPadding (BYVAL cx AS INTEGER, BYVAL cy AS INTEGER)
   DECLARE FUNCTION SetText (BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE SUB SetToolTips (BYVAL hwndTT AS HWND)
   DECLARE FUNCTION SetUnicodeFormat (BYVAL fUnicode AS BOOLEAN) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CTab constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTab (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTab (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTab (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTab (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTab destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTab
END DESTRUCTOR
' ========================================================================================

' ########################################################################################
'                                      *** TAB ***
' ########################################################################################

' ========================================================================================
' Calculates a tab control's display area given a window rectangle, or calculates the
' window rectangle that would correspond to a specified display area.
' ========================================================================================
PRIVATE SUB CTab.AdjustRect (BYVAL fLarger AS BOOLEAN, BYREF rc AS RECT)
   SendMessageW m_hCtl, TCM_ADJUSTRECT, fLarger, CAST(LPARAM, @rc)
END SUB
' ========================================================================================

' ========================================================================================
' Removes all items from a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.DeleteAllItems () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_DELETEALLITEMS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Removes an item from a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.DeleteItem (BYVAL iItem AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_DELETEITEM, iItem, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets items in a tab control, clearing any that were set to the TCIS_BUTTONPRESSED state.
' ========================================================================================
PRIVATE SUB CTab.DeselectAll (BYVAL fExcludeFocus AS BOOLEAN)
   SendMessageW m_hCtl, TCM_DESELECTALL, fExcludeFocus, 0
END SUB
' ========================================================================================

' ========================================================================================
' Returns the index of the item that has the focus in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetCurFocus () AS LONG
   RETURN SendMessageW(m_hCtl, TCM_GETCURFOCUS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the currently selected tab in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetCurSel () AS LONG
   RETURN SendMessageW(m_hCtl, TCM_GETCURSEL, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the extended styles that are currently in use for the tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetExtendedStyle () AS DWORD
   RETURN SendMessageW(m_hCtl, TCM_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the 0-based index in the tab control's image list.
' Returns the index of the image or -1 if not found.
' ========================================================================================
PRIVATE FUNCTION CTab.GetImageIndex (BYVAL iItem AS DWORD) AS LONG
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_IMAGE
   FUNCTION = -1
   IF SendMessageW(m_hCtl, TCM_GETITEMW, cast(WPARAM, iItem), cast(LPARAM, @ttc_item)) THEN
      FUNCTION = ttc_item.iImage
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the image list associated with a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetImageList () AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, TCM_GETIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about a tab in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetItem (BYVAL iItem AS DWORD, BYREF pItem AS TCITEMW) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_GETITEMW, iItem, CAST(LPARAM, @pItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of tabs in the tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetItemCount () AS LONG
   RETURN SendMessageW(m_hCtl, TCM_GETITEMCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle for a tab in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetItemRect (BYVAL iItem AS DWORD, BYREF rc AS RECT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_GETITEMRECT, iItem, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTab.GetItemRect (BYVAL iItem AS DWORD) AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, TCM_GETITEMRECT, iItem, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of rows of tabs in the tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetRowCount () AS LONG
   RETURN SendMessageW(m_hCtl, TCM_GETROWCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the name of a tab in a Tab control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTab.GetText (BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT
   ttc_item.pszText = pwszText
   ttc_item.cchTextMax = cchTextMax
   RETURN SendMessageW(m_hCtl, TCM_GETITEMW, cast(WPARAM, nTabIndex), cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTab.GetText (BYVAL nTabIndex AS DWORD) AS DWSTRING
   DIM wszText AS WSTRING * 260
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT
   ttc_item.pszText = @wszText
   ttc_item.cchTextMax = SIZEOF(wszText)
   SendMessageW(m_hCtl, TCM_GETITEMW, cast(WPARAM, nTabIndex), cast(LPARAM, @ttc_item))
   RETURN wszText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the ToolTip control associated with a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetToolTips () AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TCM_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the UNICODE character format flag for the control.
' ========================================================================================
PRIVATE FUNCTION CTab.GetUnicodeFormat () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the highlight state of a tab item.
' ========================================================================================
PRIVATE FUNCTION CTab.HighlightItem (BYVAL idItem AS LONG, BYVAL fHighlight AS WORD) AS LONG
   RETURN SendMessageW(m_hCtl, TCM_HIGHLIGHTITEM, idItem, MAKELONG(fHighlight, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which tab, if any, is at a specified screen position.
' ========================================================================================
PRIVATE FUNCTION CTab.HitTest (BYREF info AS TC_HITTESTINFO) AS LONG
   RETURN SendMessageW(m_hCtl, TCM_HITTEST, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new tab in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.InsertItem (BYVAL iItem AS DWORD, BYREF item AS TCITEMW) AS LONG
   RETURN SendMessageW(m_hCtl, TCM_INSERTITEMW, iItem, CAST(LPARAM, @Item))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts a new tab in a tab control.
' - hTab: Handle to the tab control.
' - nTabInxex = Zero-based index of the tab
' - iImage = Zero-based index of the image in the image list or -1 for no image.
' - pwszText = Pointer to a null-terminated string that contains the tab text when item
'     information is being set.
' - lParam: Application-defined data associated with the tab control item.
' Returns the index of the new tab if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CTab.InsertTab (BYVAL nTabIndex AS DWORD, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0) AS LONG
   IF pwszText = NULL THEN RETURN -1
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT OR TCIF_IMAGE OR TCIF_RTLREADING OR TCIF_PARAM OR TCIF_STATE OR TCIS_BUTTONPRESSED OR TCIS_HIGHLIGHTED
   ttc_item.pszText = pwszText
   ttc_item.cchTextMax = LEN(*pwszText)
   ttc_item.iImage = iImage
   ttc_item.lParam = lParam
   RETURN SendMessageW(m_hCtl, TCM_INSERTITEMW, cast(WPARAM, nTabIndex), cast(..LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of tabs and calls InsertTab to add a new tab.
' ========================================================================================
PRIVATE FUNCTION CTab.AddTab (BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0) AS LONG
   IF pwszText = NULL THEN RETURN -1
   DIM nTabIndex AS LONG = SendMessageW(m_hCtl, TCM_GETITEMCOUNT, 0, 0)
   RETURN InsertTab(nTabIndex, iImage, pwszText, lParam)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an image from a tab control's image list.
' ========================================================================================
PRIVATE SUB CTab.RemoveImage (BYVAL iImage AS LONG)
   SendMessageW m_hCtl, TCM_REMOVEIMAGE, iImage, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the focus to a specified tab in a tab control.
' ========================================================================================
PRIVATE SUB CTab.SetCurFocus (BYVAL iItem AS LONG)
   SendMessageW m_hCtl, TCM_SETCURFOCUS, iItem, 0
END SUB
' ========================================================================================

' ========================================================================================
' Selects a tab in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.SetCurSel (BYVAL iItem AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, TCM_SETCURSEL, iItem, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the extended styles that the tab control will use.
' ========================================================================================
PRIVATE FUNCTION CTab.SetExtendedStyle (BYVAL dwExStyle AS DWORD) AS DWORD
   RETURN SendMessageW(m_hCtl, TCM_SETEXTENDEDSTYLE, 0, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an image list to a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.SetImageList (BYVAL himl AS HIMAGELIST) AS BOOLEAN
    RETURN SendMessageW(m_hCtl, TCM_SETIMAGELIST, 0, CAST(LPARAM, himl))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the zero-based index in the tab control's image list.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTab.SetImageIndex (BYVAL iItem AS LONG, BYVAL iImage AS LONG) AS BOOLEAN
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_IMAGE
   ttc_item.iImage = iImage
   RETURN SendMessageW(m_hCtl, TCM_SETITEMW, iItem, cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets some or all of a tab's attributes.
' ========================================================================================
PRIVATE FUNCTION CTab.SetItem (BYVAL iItem AS LONG, BYREF item AS TCITEMW) AS LONG
   RETURN SendMessageW(m_hCtl, TCM_SETITEMW, iItem, CAST(LPARAM, @item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the number of bytes per tab reserved for application-defined data in a tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.SetItemExtra (BYVAL cb AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_SETITEMEXTRA, cb, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width and height of tabs in a fixed-width or owner-drawn tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.SetItemSize (BYVAL x AS LONG, BYVAL y AS LONG) AS DWORD
   RETURN SendMessageW(m_hCtl, TCM_SETITEMSIZE, 0, MAKELONG(x, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width and height of tabs in a fixed-width or owner-drawn tab control.
' ========================================================================================
PRIVATE FUNCTION CTab.SetMinTabWidth (BYVAL cx AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, TCM_SETMINTABWIDTH, 0, cx)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the amount of space (padding) around each tab's icon and label in a tab control.
' ========================================================================================
PRIVATE SUB CTab.SetPadding (BYVAL cx AS INTEGER, BYVAL cy AS INTEGER)
   SendMessageW m_hCtl, TCM_SETPADDING, 0, MAKELONG(cx, cy)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the name of a tab in a Tab control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTab.SetText (BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT
   ttc_item.pszText = pwszText
   ttc_item.cchTextMax = LEN(*pwszText)
   RETURN SendMessageW(m_hCtl, TCM_SETITEMW, cast(WPARAM, nTabIndex), cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a ToolTip control to a tab control. You can use this macro or send the
' TCM_SETTOOLTIPS message explicitly.
' ========================================================================================
PRIVATE SUB CTab.SetToolTips (BYVAL hwndTT AS HWND)
   SendMessageW m_hCtl, TCM_SETTOOLTIPS, CAsT(WPARAM, hwndTT), 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the Unicode character format flag for the control. This message allows you to
' change the character set used by the control at run time rather than having to re-create
' the control.
' ========================================================================================
FUNCTION CTab.SetUnicodeFormat (BYVAL fUnicode AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TCM_SETUNICODEFORMAT, fUnicode, 0)
END FUNCTION
' ========================================================================================

END NAMESPACE
