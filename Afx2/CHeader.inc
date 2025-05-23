' ########################################################################################
' Platform: Microsoft Windows
' Filename: CHeader.inc
' Purpose:  Windows common controls - Header control
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
' CHeader class
' ########################################################################################

TYPE CHeader EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION ClearAllFilters() AS LONG
   DECLARE FUNCTION ClearFilter (BYVAL i AS DWORD) AS LONG
   DECLARE FUNCTION CreateDragImage (BYVAL i AS DWORD) AS DWORD
   DECLARE FUNCTION DeleteItem (BYVAL i AS DWORD) AS LONG
   DECLARE FUNCTION EditFilter (BYVAL i AS DWORD, BYVAL fDiscardChanges AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION GetBitmapMargin () AS LONG
   DECLARE FUNCTION GetFocusedItem () AS LONG
   DECLARE FUNCTION GetImageList () AS HIMAGELIST
   DECLARE FUNCTION GetItem (BYVAL iItem AS DWORD, BYREF hdi AS HDITEMW) AS BOOLEAN
   DECLARE FUNCTION GetItemCount () AS LONG
   DECLARE FUNCTION GetItemDropDownRect (BYVAL iItem AS LONG, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetItemDropDownRect (BYVAL iItem AS LONG) AS RECT
   DECLARE FUNCTION GetItemOrder (BYVAL nItem AS LONG) AS LONG
   DECLARE FUNCTION GetItemRect (BYVAL iItem AS LONG, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetItemRect (BYVAL iItem AS LONG) AS RECT
   DECLARE FUNCTION GetItemText (BYVAL nItem AS LONG, BYVAL pszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetItemWidth (BYVAL nItem AS LONG) AS LONG
   DECLARE FUNCTION GetOrderArray (BYVAL iCount AS LONG, BYVAL lpiArray AS LONG PTR) AS LONG
   DECLARE FUNCTION GetOverflowRect (BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetOverflowRect () AS RECT
   DECLARE FUNCTION GetStateImageList () AS HIMAGELIST
   DECLARE FUNCTION GetText (BYVAL iCol AS LONG, BYVAL cchTextMax AS LONG = 0) AS DWSTRING
   DECLARE FUNCTION HitTest (BYREF hdhti AS HDHITTESTINFO) AS LONG
   DECLARE FUNCTION HitTest () AS HDHITTESTINFO
   DECLARE FUNCTION InsertItem (BYVAL i AS DWORD, BYREF hdi AS HDITEM) AS LONG
   DECLARE FUNCTION Layout (BYREF hdl AS HDLAYOUT) AS BOOLEAN
   DECLARE FUNCTION Layout () AS HDLAYOUT
   DECLARE FUNCTION OrderToIndex (BYVAL i AS LONG) AS LONG
   DECLARE FUNCTION SetBitmapMargin (BYVAL iWidth AS DWORD) AS LONG
   DECLARE FUNCTION SetFilterChangeTimeout (BYVAL i AS LONG) AS LONG
   DECLARE FUNCTION SetFocusedItem (BYVAL iItem AS LONG) AS BOOLEAN
   DECLARE FUNCTION SetHotDivider (BYVAL fPos AS BOOLEAN, BYVAL dwInputValue AS LONG) AS LONG
   DECLARE FUNCTION SetImageList (BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetItem (BYVAL i AS DWORD, BYREF hdi AS HDITEMW) AS LONG
   DECLARE FUNCTION SetItemWidth (BYVAL iItem AS LONG, BYVAL nWidth AS LONG) AS BOOLEAN
   DECLARE FUNCTION SetItemText (BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION SetItemOrder (BYVAL nItem AS LONG, BYVAL iOrder AS LONG) AS LONG
   DECLARE FUNCTION SetOrderArray (BYVAL iSize AS DWORD, BYVAL lpiArray AS LONG PTR) AS LONG
   DECLARE FUNCTION SetStateImageList (BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetText (BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CHeader constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CHeader (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CHeader (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CHeader (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CHeader (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CHeader destructor
' ========================================================================================
PRIVATE DESTRUCTOR CHeader
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                      *** HEADER ***
' ########################################################################################

' ========================================================================================
' Clears all of the filters for a given header control.
' ========================================================================================
PRIVATE FUNCTION CHeader. ClearAllFilters() AS LONG
   RETURN SendMessageW(m_hCtl, HDM_CLEARFILTER, -1, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears the filter for a given header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.ClearFilter (BYVAL i AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_CLEARFILTER, i, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a transparent version of an item image within an existing header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.CreateDragImage (BYVAL i AS DWORD) AS DWORD
   RETURN SendMessageW(m_hCtl, HDM_CREATEDRAGIMAGE, i, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes an item from a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.DeleteItem (BYVAL i AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_DELETEITEM, i, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the input focus to the edit box when a filter button has the focus.
' ========================================================================================
PRIVATE FUNCTION CHeader.EditFilter (BYVAL i AS DWORD, BYVAL fDiscardChanges AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, HDM_EDITFILTER, i, fDiscardChanges)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the margin (in pixels) of a bitmap in an existing header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetBitmapMargin () AS LONG
   RETURN SendMessageW(m_hCtl, HDM_GETBITMAPMARGIN, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the item in a header control that has the focus.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetFocusedItem () AS LONG
   RETURN SendMessageW(m_hCtl, HDM_GETFOCUSEDITEM, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle to the image list that has been set for an existing header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetImageList () AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, HDM_GETIMAGELIST, HDSIL_NORMAL, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about an item in a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItem (BYVAL iItem AS DWORD, BYREF hdi AS HDITEMW) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, HDM_GETITEMW, iItem, CAST(LPARAM, @hdi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a count of the items in a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemCount () AS LONG
   RETURN SendMessageW(m_hCtl, HDM_GETITEMCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the coordinates of the drop-down button for a specfied item in a header control.
' The header control must be of type HDF_SPLITBUTTON.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemDropDownRect (BYVAL iItem AS LONG, BYREF rc AS RECT) AS BOOLEAN
   IF VARPTR(rc) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, HDM_GETITEMDROPDOWNRECT, iItem, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemDropDownRect (BYVAL iItem AS LONG) AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, HDM_GETITEMDROPDOWNRECT, iItem, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the bounding rectangle for a given item in a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemRect (BYVAL iItem AS LONG, BYREF rc AS RECT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, HDM_GETITEMRECT, iItem, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemRect (BYVAL iItem AS LONG) AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, HDM_GETITEMRECT, iItem, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the order of the specified item of a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemOrder (BYVAL nItem AS LONG) AS LONG
   DIM hdi AS HDITEMW
   hdi.mask = HDI_ORDER
   IF SendMessageW(m_hCtl, HDM_GETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi))) THEN RETURN hdi.iOrder
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the text of the specified column.
' - m_hCtl: Handle to the header control.
' - nItem: The index of the item for which information is to be retrieved.
' - pwszText: A pointer to an item string. If the text is being retrieved from the control,
'   this member must be initialized to point to a character buffer. If this member is set
'   to LPSTR_TEXTCALLBACK, the control will request text information for this item by sending
'   an HDN_GETDISPINFO notification code. Note that although the header control allows a
'   string of any length to be stored as item text, only the first 260 characters are displayed.
' - cchTextMax : Maximum number of characters to return. Both this value and the size of the
'   buffer pointed by pwszText must be one character bigger that the wanted length of the text
'   to return to make room for the null character terminator.
' Return value: Returns TRUE or FALSE.
' Usage example:
' DIM wszText AS WSTRING * 260
' CHeader.GetItemText(hTreeView, hItem, @wszText, 260)
' AfxMsg(wszText)
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemText (BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM hdi AS HDITEMW
   hdi.mask = HDI_TEXT
   hdi.psztext = pwszText
   hdi.cchtextmax = cchtextmax
   RETURN SendMessageW(m_hCtl, HDM_GETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the specified item of a header control.
' Returns the index of the item at the specified position, if any, or – 1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetItemWidth (BYVAL nItem AS LONG) AS LONG
   DIM hdi AS HDITEMW
   hdi.mask = HDI_WIDTH
   IF SendMessageW(m_hCtl, HDM_GETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi))) THEN RETURN hdi.cxy
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current left-to-right order of items in a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetOrderArray (BYVAL iCount AS LONG, BYVAL lpiArray AS LONG PTR) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_GETORDERARRAY, iCount, CAST(LPARAM, lpiArray))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the coordinates of the drop-down overflow area for a specified header control.
' The header control must be of type HDF_SPLITBUTTON.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetOverflowRect (BYREF rc AS RECT) AS BOOLEAN
   IF VARPTR(rc) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, HDM_GETOVERFLOWRECT, 0, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CHeader.GetOverflowRect () AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, HDM_GETOVERFLOWRECT, 0, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle to the image list that has been set for an existing header control state.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetStateImageList () AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, HDM_GETIMAGELIST, HDSIL_STATE, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of the specified column.
' ========================================================================================
PRIVATE FUNCTION CHeader.GetText (BYVAL iCol AS LONG, BYVAL cchTextMax AS LONG = 0) AS DWSTRING
   DIM wszText AS DWSTRING
   IF cchTextMax < 1 THEN cchTextMax = MAX_PATH + 1
   wszText = WSPACE(cchTextMax + 1)
   DIM hdi AS HDITEMW
   hdi.mask = HDI_TEXT
   hdi.cchTextMax = cchTextMax + 1
   hdi.pszText = STRPTR(wszText)
   IF SendMessageW(m_hCtl, HDM_GETITEMW, iCol, CAST(LPARAM, @hdi)) THEN RETURN wszText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tests a point to determine which header item, if any, is at the specified point.
' ========================================================================================
PRIVATE FUNCTION CHeader.HitTest (BYREF hdhti AS HDHITTESTINFO) AS LONG
   IF VARPTR(hdhti) = 0 THEN RETURN 0
   RETURN SendMessageW(m_hCtl, HDM_HITTEST, 0, CAST(LPARAM, @hdhti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CHeader.HitTest () AS HDHITTESTINFO
   DIM hdhti AS HDHITTESTINFO
   SendMessageW(m_hCtl, HDM_HITTEST, 0, CAST(LPARAM, @hdhti))
   RETURN hdhti
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new item into a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.InsertItem (BYVAL i AS DWORD, BYREF hdi AS HDITEM) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_INSERTITEM, i, CAST(LPARAM, @hdi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the correct size and position of a header control within the parent window.
' ========================================================================================
PRIVATE FUNCTION CHeader.Layout (BYREF hdl AS HDLAYOUT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, HDM_LAYOUT, 0, CAST(LPARAM, @hdl))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CHeader.Layout () AS HDLAYOUT
   DIM hdl AS HDLAYOUT
   SendMessageW(m_hCtl, HDM_LAYOUT, 0, CAST(LPARAM, @hdl))
   RETURN hdl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves an index value for an item based on its order in the header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.OrderToIndex (BYVAL i AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_ORDERTOINDEX, i, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width of the margin for a bitmap in an existing header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetBitmapMargin (BYVAL iWidth AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_SETBITMAPMARGIN, iWidth, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the timeout interval between the time a change takes place in the filter
' attributes and the posting of an HDN_FILTERCHANGE notification.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetFilterChangeTimeout (BYVAL i AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_SETFILTERCHANGETIMEOUT, 0, i)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus to a specified item in a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetFocusedItem (BYVAL iItem AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, HDM_SETFOCUSEDITEM, 0, iItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the color of a divider between header items to indicate the destination of an
' external drag-and-drop operation.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetHotDivider (BYVAL fPos AS BOOLEAN, BYVAL dwInputValue AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_SETHOTDIVIDER, fPos, dwInputValue)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an image list to an existing header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetImageList (BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   RETURN CAsT (HIMAGELIST, SendMessageW(m_hCtl, HDM_SETIMAGELIST, 0, CAST(LPARAM, himl)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the attributes of the specified item in a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetItem (BYVAL i AS DWORD, BYREF hdi AS HDITEMW) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_SETITEMW, i, CAST(LPARAM, @hdi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the order of the specified item of a header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetItemOrder (BYVAL nItem AS LONG, BYVAL iOrder AS LONG) AS LONG
   DIM hdi AS HDITEMW
   hdi.mask = HDI_ORDER
   hdi.iOrder = iOrder
   RETURN SendMessageW(m_hCtl, HDM_SETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified item. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetItemText (BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM hdi AS HDITEMW
   hdi.mask = HDI_TEXT
   hdi.cchTextMax = LEN(*pwszText)
   hdi.pszText = pwszText
   RETURN SendMessageW(m_hCtl, HDM_SETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the text of the specified item. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetText (BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM hdi AS HDITEMW
   hdi.mask = HDI_TEXT
   hdi.cchTextMax = LEN(*pwszText)
   hdi.pszText = pwszText
   RETURN SendMessageW(m_hCtl, HDM_SETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width of the specified item of a header control.
' Returns nonzero upon success, or zero otherwise.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetItemWidth (BYVAL nItem AS LONG, BYVAL nWidth AS LONG) AS BOOLEAN
   DIM hdi AS HDITEMW
   hdi.mask = HDI_WIDTH
   hdi.cxy = nWidth
   RETURN SendMessageW(m_hCtl, HDM_SETITEMW, nItem, CAST(LPARAM, CAST(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the left-to-right order of header items.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetOrderArray (BYVAL iSize AS DWORD, BYVAL lpiArray AS LONG PTR) AS LONG
   RETURN SendMessageW(m_hCtl, HDM_SETORDERARRAY, iSize, CAST(LPARAM, lpiArray))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an image list to an existing header control.
' ========================================================================================
PRIVATE FUNCTION CHeader.SetStateImageList (BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, HDM_SETIMAGELIST, HDSIL_STATE, CAST(LPARAM, himl)))
END FUNCTION
' ========================================================================================

END NAMESPACE
