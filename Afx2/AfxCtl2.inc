' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxCtl2.inc
' Contents: Common controls macros and functions
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
#include once "win/richedit.bi"
#include once "win/uxtheme.bi"
#include once "win/windowsx.bi"

' Button
DECLARE FUNCTION Button_SetImageListXY (BYVAL hButton AS HWND, BYVAL himl AS HIMAGELIST, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG, BYVAL uALign AS DWORD = 0) AS BOOLEAN

' Combobox
DECLARE FUNCTION Combobox_CalcHorizontalExtent (BYVAL hComboBox AS HWND) AS LONG
DECLARE FUNCTION ComboBox_GetEditBoxHandle (BYVAL hComboBox AS HWND) AS HWND
DECLARE FUNCTION ComboBox_GetListBoxHandle (BYVAL hComboBox AS HWND) AS HWND

' Header
DECLARE FUNCTION Header_GetItemWidth (BYVAL hwndHD AS HWND, BYVAL nItem AS LONG) AS LONG
DECLARE FUNCTION Header_SetItemWidth (BYVAL hwndHD AS HWND, BYVAL iItem AS LONG, BYVAL nWidth AS LONG) AS BOOLEAN
DECLARE FUNCTION Header_GetItemOrder (BYVAL hwndHD AS HWND, BYVAL nItem AS LONG) AS LONG
DECLARE FUNCTION Header_SetItemOrder (BYVAL hwndHD AS HWND, BYVAL nItem AS LONG, BYVAL iOrder AS LONG) AS LONG
DECLARE FUNCTION Header_GetItemText (BYVAL hwndHD AS HWND, BYVAL nItem AS LONG, BYVAL pszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
DECLARE FUNCTION Header_SetItemText (BYVAL hwndHD AS HWND, BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN

' Listbox
DECLARE FUNCTION ListBox_ReplaceString (BYVAL hCtl AS HWND, BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
DECLARE FUNCTION Listbox_CalcHorizontalExtent (BYVAL hListBox AS HWND) AS LONG

' Listview
DECLARE SUB ListView_SelectAllItems (BYVAL hCtl AS HWND)
DECLARE SUB ListView_MakeHeaderFlat (BYVAL hCtl AS HWND)
DECLARE FUNCTION ListView_AutoSizeColumns (BYVAL hCtl AS HWND) AS BOOLEAN
DECLARE FUNCTION ListView_AutoSizeHeader (BYVAL hCtl AS HWND) AS BOOLEAN
DECLARE FUNCTION ListView_GetColumnCount (BYVAL hCtl AS HWND) AS LONG
DECLARE FUNCTION ListView_GetColumnOrder (BYVAL hCtl AS HWND, BYVAL nCol AS LONG) AS LONG
DECLARE FUNCTION ListView_GetOrderArray (BYVAL hCtl AS HWND, BYVAL nSize AS DWORD, BYVAL lpiArray AS LONG PTR) AS LONG
DECLARE FUNCTION ListView_GetTooltipsFont (BYVAL hCtl AS HWND) AS HFONT
DECLARE SUB ListView_SetTooltipsFont (BYVAL hCtl AS HWND, BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
DECLARE FUNCTION ListView_GetColumnText (BYVAL hCtl AS HWND, BYVAL nCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
DECLARE FUNCTION ListView_GetHeaderText (BYVAL hCtl AS HWND, BYVAL nCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS LONG
DECLARE FUNCTION ListView_AddColumn (BYVAL hCtl AS HWND, BYVAL nCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL nWidth AS LONG, BYVAL nFormat AS LONG = 0) AS LONG
DECLARE FUNCTION ListView_AddItem (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL nImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
DECLARE FUNCTION ListView_AppendItem (BYVAL hCtl AS HWND, BYVAL nImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG

' Progress bar
DECLARE FUNCTION ProgressBar_IncPos (BYVAL hProgress AS HWND, BYVAL nIncVal AS LONG) AS LONG

' Status bar
DECLARE FUNCTION StatusBar_GetHeight (BYVAL hStatusBar AS HWND) AS LONG

' Tab control
DECLARE FUNCTION TabCtrl_InsertTab (BYVAL hwndTab AS HWND, BYVAL nTabIndex AS DWORD, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0) AS LONG
DECLARE FUNCTION TabCtrl_AddTab (BYVAL hwndTab AS HWND, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0) AS LONG
DECLARE FUNCTION TabCtrl_GetText (BYVAL hwndTab AS HWND, BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
DECLARE FUNCTION TabCtrl_SetText (BYVAL hwndTab AS HWND, BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
DECLARE FUNCTION TabCtrl_GetImageIndex (BYVAL hwndTab AS HWND, BYVAL iItem AS DWORD) AS LONG
DECLARE FUNCTION TabCtrl_SetImageIndex (BYVAL hwndTab AS HWND, BYVAL iItem AS DWORD, BYVAL iImage AS LONG) AS BOOLEAN

' Toolbar
DECLARE FUNCTION Toolbar_AddButton (BYVAL hToolBar AS HWND, BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
                 BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
DECLARE FUNCTION Toolbar_AddSeparator (BYVAL hToolBar AS HWND, BYVAL nWidth AS LONG = 0) AS BOOLEAN
DECLARE FUNCTION Toolbar_InsertButton OVERLOAD (BYVAL hToolBar AS HWND, BYVAL idxButton AS LONG, BYVAL lpButton AS TBBUTTON PTR) AS BOOLEAN
DECLARE FUNCTION Toolbar_InsertButton OVERLOAD (BYVAL hToolBar AS HWND, BYVAL idxButton AS LONG, BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
                 BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
DECLARE FUNCTION Toolbar_InsertSeparator (BYVAL hToolBar AS HWND, BYVAL idxButton AS LONG, BYVAL nWidth AS LONG = 0) AS BOOLEAN
DECLARE FUNCTION Toolbar_GetObject (BYVAL hToolBar AS HWND) AS IDropTarget PTR

' Tooltips
DECLARE FUNCTION Tooltip_Add (BYVAL hwndCtrl AS HWND, BYREF wszText AS CONST WSTRING = "", BYVAL bBalloon AS BOOLEAN = FALSE, BYVAL bCentered AS BOOLEAN = FALSE) AS HWND
DECLARE SUB Tooltip_SetText (BYVAL hTooltip AS HWND, BYVAL hwndCtrl AS HWND, BYREF wszText AS CONST WSTRING)
DECLARE SUB Tooltip_Delete (BYVAL hTooltip AS HWND, BYVAL hwndCtrl AS HWND)

' Treeview
DECLARE FUNCTION TreeView_GetItemEx (BYVAL hwndTV AS HWND, BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
DECLARE FUNCTION TreeView_AddItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, BYVAL hInsertAfter AS HTREEITEM, _
   BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
DECLARE FUNCTION TreeView_AppendItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, _
   BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
DECLARE FUNCTION TreeView_InsertFirstItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, _
   BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
DECLARE FUNCTION TreeView_AddRootItem (BYVAL hwndTV AS HWND, BYVAL pwszText AS WSTRING PTR, _
   BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
DECLARE FUNCTION TreeView_InsertSortedItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, _
   BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
DECLARE FUNCTION TreeView_GetBold (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM) AS BOOLEAN
DECLARE FUNCTION TreeView_IsItemChecked (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM) AS BOOLEAN
DECLARE FUNCTION TreeView_GetISearchStringLen (BYVAL hwndTV AS HWND) AS LONG
DECLARE FUNCTION TreeView_GetlParam (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM) AS LPARAM
DECLARE SUB TreeView_SetBold (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL fBold AS BOOLEAN)
DECLARE FUNCTION TreeView_SetItemEx (BYVAL hwndTV AS HWND, BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
DECLARE FUNCTION TreeView_SetItemText (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
DECLARE SUB TreeView_ExpandChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB TreeView_ExpandSiblings (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB TreeView_ExpandAllChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB TreeView_ExpandAllItems (BYVAL hwndTV AS HWND)
DECLARE SUB TreeView_CollapseChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB TreeView_CollapseSiblings (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB TreeView_CollapseAllChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
DECLARE SUB TreeView_CollapseAllItems (BYVAL hwndTV AS HWND)
DECLARE FUNCTION TreeView_GetItemText (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchtextmax AS LONG) AS BOOLEAN
DECLARE FUNCTION TreeView_GetParentText (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
DECLARE FUNCTION TreeView_ItemExists (BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
DECLARE FUNCTION TreeView_ItemExistsIgnoreCase (BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN


' ########################################################################################
'                                     *** BUTTON ***
' ########################################################################################

#define Button_Click(hCtl) SendMessageW(hCtl, BM_CLICK, 0, 0)
#define Button_GetImage(hCtl, imageType) CAST(HANDLE, SendMessageW(hCtl, BM_GETIMAGE, imageType, 0))
#define Button_GetBitmap(hCtl) CAST(HANDLE, SendMessageW(hCtl, BM_GETIMAGE, IMAGE_BITMAP, 0))
#define Button_GetIcon(hCtl) CAST(HANDLE, SendMessageW(hCtl, BM_GETIMAGE, IMAGE_ICON, 0))
#define Button_SetBitmap(hCtl, hbmp) cast(HBITMAP, SendMessageW(hCtl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hbmp)))
#define Button_SetIcon(hCtl, hIcon) cast(..HICON, SendMessageW(hCtl, BM_SETIMAGE, IMAGE_ICON, cast(LPARAM, hIcon)))
#define Button_SetImage(hCtl, ImageType, hImage) cast(HANDLE, SendMessageW(hCtl, BM_SETIMAGE, ImageType, cast(LPARAM, hImage)))
#define Button_DeleteBitmap(hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hCtl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, NULL))))
#define Button_DeleteIcon(hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hCtl, BM_SETIMAGE, IMAGE_ICON, cast(LPARAM, NULL))))
#define Button_DeleteImage(hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hCtl, BM_SETIMAGE, ImageType, cast(LPARAM, NULL))))
#if _WIN32_WINNT = &h0602
#define Button_SetDontClick(hCtl, fState) SendMessageW(hCtl, BM_SETDONTCLICK, cast(WPARAM, fState), 0)
#endif

' ========================================================================================
' Assigns an image list to a button control. Requires Windows XP or superior.
' Parameters:
' - hButon     = The handle of the button.
' - hImageList =
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
' Note:  To use this application programming interface (API), you must provide a manifest
' specifying Comclt32.dll version 6.0.
' ========================================================================================
PRIVATE FUNCTION Button_SetImageListXY (BYVAL hButton AS HWND, BYVAL himl AS HIMAGELIST, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG, BYVAL uALign AS DWORD = 0) AS BOOLEAN
   DIM bi AS BUTTON_IMAGELIST = (himl, (nLeft, nTop, nRight, nBottom), uAlign)
   FUNCTION = SendMessageW(hButton, BCM_SETIMAGELIST, 0, cast(LPARAM, @bi))
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                    *** COMBOBOX ***
' ########################################################################################

#define ComboBox_GetComboBoxInfo(hCtl, pcbi) CBOOL(SendMessageW(hCtl, CB_GETCOMBOBOXINFO, 0, cast(LPARAM, pcbi)))
#define ComboBox_GetDroppedWidth(hCtl) SendMessageW(hCtl, CB_GETDROPPEDWIDTH, 0, 0)
#define ComboBox_GetHorizontalExtent(hCtl) SendMessageW(hCtl, CB_GETHORIZONTALEXTENT, 0, 0)
#define ComboBox_GetLocale(hCtl) SendMessageW(hCtl, CB_GETLOCALE, 0, 0)
#define ComboBox_GetTopIndex(hCtl) SendMessageW(hCtl, CB_GETTOPINDEX, 0, 0)
#define ComboBox_InitStorage(hCtl) SendMessageW(hCtl, CB_INITSTORAGE, CAST(WPARAM, items), CAST(LPARAM, cb))
#define ComboBox_SetDroppedWidth(hCtl) SendMessageW(hCtl, CB_SETDROPPEDWIDTH, nWidth, 0)
#define ComboBox_SetHorizontalExtent(hCtl) SendMessageW(hCtl, CB_SETHORIZONTALEXTENT, cast(WPARAM, nWidth), 0)
#define ComboBox_SetLocale(hCtl) SendMessageW(hCtl, CB_SETLOCALE, cast(WPARAM, lcid), 0)
#define ComboBox_SetTopIndex(hCtl, index) SendMessageW(hCtl, CB_SETTOPINDEX, cast(WPARAM, index), 0)

' ========================================================================================
' Calculates the size, in logical units, of the widest string item.
' ========================================================================================
PRIVATE FUNCTION Combobox_CalcHorizontalExtent (BYVAL hComboBox AS HWND) AS LONG
   IF IsWindow(hComboBox) = 0 THEN EXIT FUNCTION
   ' // Get an handle to the device context of the control
   DIM hDC AS HDC = GetDC(hComboBox)
   ' // Get the number of items
   DIM i AS LONG, cxScrollWidth AS LONG
   DIM cItems AS LRESULT = SendMessageW(hComboBox, CB_GETCOUNT, 0, 0)
   FOR i = 0 TO cItems - 1
      ' // Get the text of the item
      DIM cbLen AS LONG = SendMessageW(hComboBox, CB_GETLBTEXTLEN, i, 0)
      IF cbLen = 0 THEN CONTINUE FOR
      DIM pbuffer AS WSTRING PTR = CAllocate(cbLen * 2)
      SendMessageW(hComboBox, CB_GETLBTEXT, i, CAST(LPARAM, pbuffer))
      ' // Computes the width and height of the text
      DIM tsize AS SIZE
      GetTextExtentPoint32W hDC, pbuffer, cbLen, @tsize
      cxScrollWidth = MAX(cxScrollWidth, tsize.cx)
      Deallocate pbuffer
   NEXT
   ' // Release the device context
   ReleaseDC hComboBox, hDC
   ' // Return the width of the text
   FUNCTION = cxScrollWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the edit box of the combobox
' ========================================================================================
PRIVATE FUNCTION ComboBox_GetEditBoxHandle (BYVAL hComboBox AS HWND) AS HWND
   DIM tci AS COMBOBOXINFO
   tci.cbSize = sizeof(COMBOBOXINFO)
   IF SendMessageW(hComboBox, CB_GETCOMBOBOXINFO, 0, cast(LPARAM, cast(COMBOBOXINFO PTR, @tci))) = 0 THEN EXIT FUNCTION
   FUNCTION = tci.hwndItem
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the drop down of the combobox
' ========================================================================================
PRIVATE FUNCTION ComboBox_GetListBoxHandle (BYVAL hComboBox AS HWND) AS HWND
   DIM tci AS COMBOBOXINFO
   tci.cbSize = sizeof(COMBOBOXINFO)
   IF SendMessageW(hComboBox, CB_GETCOMBOBOXINFO, 0, cast(LPARAM, cast(COMBOBOXINFO PTR, @tci))) = 0 THEN EXIT FUNCTION
   FUNCTION = tci.hwndList
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                   *** COMBOBOXEX ***
' ########################################################################################

#define ComboBoxEx_DeleteItem(hCtl, index) SendMessageW(hCtl, CBEM_DELETEITEM, nIndex, 0)
#define ComboBoxEx_GetComboControl(hCtl) cast(HWND, SendMessageW(hCtl, CBEM_GETCOMBOCONTROL, 0, 0))
#define ComboBoxEx_GetEditControl(hCtl) cast(HWND, SendMessageW(hCtl, CBEM_GETEDITCONTROL, 0, 0))
#define ComboBoxEx_GetExtendedStyle(hCtl) SendMessageW(hCtl, CBEM_GETEXTENDEDSTYLE, 0, 0)
#define ComboBoxEx_GetImageList(hCtl) cast(HIMAGELIST, SendMessageW(hCtl, CBEM_GETIMAGELIST, 0, 0))
#define ComboBoxEx_GetItem(hCtl, pCBItem) SendMessageW(hCtl, CBEM_GETITEMW, 0, cast(LPARAM, pCBItem))
#define ComboBoxEx_GetUnicodeFormat(hCtl) SendMessageW(hCtl, CBEM_GETUNICODEFORMAT, 0, 0)
#define ComboBoxEx_HasEditChanged(hCtl) SendMessageW(hCtl, CBEM_HASEDITCHANGED, 0, 0)
#define ComboBoxEx_InsertItem(hCtl, pCBItem) SendMessageW(hCtl, CBEM_INSERTITEMW, 0, cast(LPARAM, pCBItem))
#define ComboBoxEx_SetExtendedStyle(hCtl, dwExMask, dwExStyle) SendMessageW(hCtl, CBEM_SETEXTENDEDSTYLE, dwExMask, dwExStyle)
#define ComboBoxEx_SetImageList(hCtl, himl) cast(HIMAGELIST, SendMessageW(hCtl, CBEM_SETIMAGELIST, 0, cast(LPARAM, himl)))
#define ComboBoxEx_SetItem(hCtl, pCBItem) SendMessageW(hCtl, CBEM_SETITEMW, 0, cast(LPARAM, pCBItem))
#define ComboBoxEx_SetUnicodeFormat(hCtl, fUnicode) SendMessageW(hCtl, CBEM_SETUNICODEFORMAT, 0, fUnicode)
#define ComboBoxEx_SetWindowTheme(hCtl, pwszStyle) SendMessageW(hCtl, CBEM_SETWINDOWTHEME, 0, cast(LPARAM, pwszStyle))


' ########################################################################################
'                                      *** EDIT ***
' ########################################################################################

#define Edit_CharFromPos(hCtl,x,y) SendMessageW(hCtl, EM_CHARFROMPOS, 0, cast(LPARAM, MAKELONG(x, y)))
#define Edit_Clear(hCtl) SendMessageW(hCtl, WM_CLEAR, 0, 0)
#define Edit_Copy(hCtl) SendMessageW(hCtl, WM_COPY, 0, 0)
#define Edit_Cut(hCtl) SendMessageW(hCtl, WM_CUT, 0, 0)
#define Edit_GetIMEStatus(hCtl) SendMessageW(hCtl, EM_GETIMESTATUS, cast(WPARAM, nStatusType), 0)
#define Edit_GetLimitText(hCtl) SendMessageW(hCtl, EM_GETLIMITTEXT, 0, 0)
#define Edit_GetMargins(hCtl) SendMessageW(hCtl, EM_GETMARGINS, 0, 0)
#define Edit_GetLeftMargin(hCtl) LOWORD(SendMessageW(hCtl, EM_GETMARGINS, 0, 0))
#define Edit_GetRightMargin(hCtl) HIWORD(SendMessageW(hCtl, EM_GETMARGINS, 0, 0))
#define Edit_GetSelStart(hCtl) LOWORD(SendMessageW(hCtl, EM_GETSEL, 0, 0))
#define Edit_GetSelEnd(hCtl) HIWORD(SendMessageW(hCtl, EM_GETSEL, 0, 0))
#define Edit_GetThumb(hCtl) SendMessageW(hCtl, EM_GETTHUMB, 0, 0)
#define Edit_LineScroll(hCtl,x,y) SendMessageW(hEdit, EM_LINESCROLL, x, y)
#define Edit_Paste(hCtl) SendMessageW(hCtl, WM_PASTE, 0, 0)
#define Edit_PosFromChar(hCtl, index) SendMessageW(hCtl, EM_POSFROMCHAR, index, 0)
#define Edit_SetCueBannerTextFocused(hCtl, pwszText, fDrawFocused) SendMessageW(hEdit, EM_SETCUEBANNER, fDrawFocused, cast(WPARAM, pwszText))
#define Edit_SetIMEStatus(hCtl, nStatusType, flags) SendMessageW(hCtl, EM_SETIMESTATUS, nStatusType, flags)
#define Edit_SetMargins(hCtl, nMargins, nWidth) SendMessageW(hCtl, EM_SETMARGINS, nMargins, nWidth)
#define Edit_SetLeftMargin(hCtl, nWidth) SendMessageW(hCtl, EM_SETMARGINS, EC_LEFTMARGIN, cast(LPARAM, MAKELONG(nWidth, 0))
#define Edit_SetRightMargin(hCtl, nWidth) SendMessageW(hCtl, EM_SETMARGINS, EC_RIGHTMARGIN, cast(LPARAM, MAKELONG(nWidth, 0))


' ########################################################################################
'                                      *** HEADER ***
' ########################################################################################

#define Header_HitTest (hCtl, phdhti) SendMessageW(hCtl, HDM_HITTEST, 0, cast(LPARAM, phdhti))

' ========================================================================================
' Gets the width of the specified item of a header control.
' Returns the index of the item at the specified position, if any, or – 1 otherwise.
' ========================================================================================
PRIVATE FUNCTION Header_GetItemWidth (BYVAL hCtl AS HWND, BYVAL nItem AS LONG) AS LONG
   DIM hdi AS HDITEMW : hdi.mask = HDI_WIDTH
   IF SendMessageW(hCtl, HDM_GETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi))) THEN FUNCTION = hdi.cxy
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width of the specified item of a header control.
' Returns nonzero upon success, or zero otherwise.
' ========================================================================================
PRIVATE FUNCTION Header_SetItemWidth (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL nWidth AS LONG) AS BOOLEAN
   DIM hdi AS HDITEMW : hdi.mask = HDI_WIDTH : hdi.cxy = nWidth
   FUNCTION = SendMessageW(hCtl, HDM_SETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the order of the specified item of a header control.
' ========================================================================================
PRIVATE FUNCTION Header_GetItemOrder (BYVAL hCtl AS HWND, BYVAL nItem AS LONG) AS LONG
   DIM hdi AS HDITEMW : hdi.mask = HDI_ORDER
   IF SendMessageW(hCtl, HDM_GETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi))) THEN FUNCTION = hdi.iOrder
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the order of the specified item of a header control.
' ========================================================================================
PRIVATE FUNCTION Header_SetItemOrder (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL iOrder AS LONG) AS LONG
   DIM hdi AS HDITEMW : hdi.mask = HDI_ORDER : hdi.iOrder = iOrder
   FUNCTION = SendMessageW(hCtl, HDM_SETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of the specified column.
' - hwndHD: Handle to the header control.
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
' Header_GetItemText(hTreeView, hItem, @wszText, 260)
' AfxMsg(wszText)
' ========================================================================================
PRIVATE FUNCTION Header_GetItemText (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DIM hdi AS HDITEMW : hdi.mask = HDI_TEXT : hdi.psztext = pwszText : hdi.cchtextmax = cchtextmax
   FUNCTION = SendMessageW(hCtl, HDM_GETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified item. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION Header_SetItemText (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM hdi AS HDITEMW : hdi.mask = HDI_TEXT : hdi.cchTextMax = LEN(*pwszText) : hdi.pszText = pwszText
   FUNCTION = SendMessageW(hCtl, HDM_SETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** HOTKEY ***
' ########################################################################################

#define HotKey_GetHotKey (hCtl) SendMessageW(hCtl, HKM_GETHOTKEY, 0, 0)
#define HotKey_SetHotKey (hCtl) SendMessageW(hCtl, HKM_SETHOTKEY, MAKELONG(vkCode, fMods), 0)
#define HotKey_SetRules (hCtl, fwCombInv, fwModInv) SendMessageW(hCtl, HKM_SETRULES, cast(WPARAM, fwCombInv), cast(LPARAM, MAKELONG(fwModInv, 0)))


' ########################################################################################
'                                   *** IP ADDRESS ***
' ########################################################################################

#define IPAddress_ClearAddress (hCtl) SendMessageW(hCtl, IPM_CLEARADDRESS, 0, 0)
#define IPAddress_GetAddress (hCtl, pdwAddr) SendMessageW(hCtl, IPM_GETADDRESS, 0, CAST(LPARAM, pdwAddr))
#define IPAddress_IsBlank (hCtl) SendMessageW(hCtl, IPM_ISBLANK, 0, 0)
#define IPAddress_SetAddress (hCtl, dwAddr) SendMessageW(hCtl, IPM_SETADDRESS, 0, cast(LPARAM, dwAddr))
#define IPAddress_SetFocus (hCtl, nField) SendMessageW(hCtl, IPM_SETFOCUS, cast(WPARAM, nField), 0)
#define IPAddress_SetRange (hCtl, nField, loRange, hiRange) SendMessageW(hCtl, IPM_SETRANGE, nField, MAKEIPRANGE(loRange, hiRange))


' ########################################################################################
'                                    *** LISTBOX ***
' ########################################################################################

#define ListBox_AddFile (hCtl, pwszFileName) SendMessageW(hCtl, LB_ADDFILE, 0, cast(LPARAM, pwszFileName))
#define ListBox_GetAnchorIndex (hCtl) SendMessageW(hCtl, LB_GETANCHORINDEX, 0, 0)
#define Listbox_GetListboxInfo (hCtl) SendMessageW(hCtl, LB_GETLISTBOXINFO, 0, 0)
#define ListBox_GetLocale (hCtl) SendMessageW(hCtl, LB_GETLOCALE, 0, 0)
#define Listbox_InitStorage (hCtl, nItems, nBytes) SendMessageW(hCtl, LB_INITSTORAGE, nItems, nBytes)
#define Listbox_ItemFromPoint (hCtl,x,y) SendMessageW(hCtl, LB_ITEMFROMPOINT, 0, MAKELONG(x, y))
#define Listbox_SelItemRangeEx (hCtl, startIndex, emdImdex) SendMessageW(hCtl, LB_SELITEMRANGEEX, startIndex, endIndex)
#define ListBox_SetAnchorIndex (hCtl, index) SendMessageW(hCtl, LB_SETANCHORINDEX, index, 0)
#define Listbox_SetCount (hCtl, nCount) SendMessageW(hCtl, LB_SETCOUNT, nCount, 0)
#define ListBox_SetLocale (hCtl, wLocaleID) SendMessageW(hCtl, LB_SETLOCALE, wLocaleID, 0)

' ========================================================================================
' Replaces a string, and his associated data, into a list box.
' The return value is the index of the position at which the string was inserted.
' If an error occurs, the return value is LB_ERR. If there is insufficient space to store
' the new string, the return value is LB_ERRSPACE.
' ========================================================================================
PRIVATE FUNCTION ListBox_ReplaceString (BYVAL hListBox AS HWND, BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
   ' // Delete the string
   DIM lRes AS LRESULT = SendMessageW(hListBox, LB_DELETESTRING, index, 0)
   IF lRes = LB_ERR THEN RETURN lRes
   ' // Insert the new string
   index = SendMessageW(hListBox, LB_INSERTSTRING, index, CAST(LPARAM, pwszNewText))
   IF index = LB_ERR OR index = LB_ERRSPACE THEN Return index
   lRes = SendMessageW(hListBox, LB_SETITEMDATA, index, CAST(LPARAM, pNewData))
   IF lRes = LB_ERR THEN Return lRes
   FUNCTION = SendMessageW(hListBox, IIF((GetWindowLongPtrW(hListBox, GWL_STYLE) AND LBS_MULTIPLESEL), LB_SETSEL, LB_SETCURSEL), index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the size, in logical units, of the widest string item.
' ========================================================================================
PRIVATE FUNCTION Listbox_CalcHorizontalExtent (BYVAL hListBox AS HWND) AS LONG
   IF IsWindow(hListbox) = 0 THEN EXIT FUNCTION
   ' // Get an handle to the device context of the control
   DIM hDC AS HDC = GetDC(hListBox)
   ' // Get the number of items
   DIM i AS LONG, cxScrollWidth AS LONG
   DIM cItems AS LRESULT = SendMessageW(hListBox, LB_GETCOUNT, 0, 0)
   FOR i = 0 TO cItems - 1
      ' // Get the text of the item
      DIM cbLen AS LONG = SendMessageW(hListBox, LB_GETTEXTLEN, i, 0)
      IF cbLen = 0 THEN CONTINUE FOR
      DIM pbuffer AS WSTRING PTR = CAllocate(cbLen * 2)
      SendMessageW(hListBox, LB_GETTEXT, i, CAST(LPARAM, pbuffer))
      ' // Computes the width and height of the text
      DIM tsize AS SIZE
      GetTextExtentPoint32W hDC, pbuffer, cbLen, @tsize
      cxScrollWidth = MAX(cxScrollWidth, tsize.cx)
      Deallocate pbuffer
   NEXT
   ' // Release the device context
   ReleaseDC hListBox, hDC
   ' // Return the width of the text
   FUNCTION = cxScrollWidth
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                    *** LISTVIEW ***
' ########################################################################################

#define ListView_GetFocusedItem (hCtl) SendMessageW(hCtl, LVM_GETNEXTITEM, -1, LVNI_ALL OR LVNI_FOCUSED)
#define ListView_GetISearchStringLen (hCtl) SendMessageW(hCtl, LVM_GETISEARCHSTRINGW, 0, NULL)
#define ListView_GetSelection (hCtl) SendMessageW(hCtl, LVM_HITTEST, cast(WPARAM, -1), cast(LPARAM, pinfo))
#define ListView_HitTestEx (hCtl) SendMessageW(hCtl, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
#define ListView_SelectItem (hCtl) ListView_SetItemState(hCtl, nIndex, LVIS_FOCUSED OR LVIS_SELECTED, &H000F)
#define ListView_UnselectItem (hCtl, index) ListView_SetItemState(hCtl, nIndex, 0, LVIS_SELECTED)
#define ListView_UnselectAllItems (hCtl) ListView_SetItemState(hCtl, -1, 0, LVIS_SELECTED OR LVIS_FOCUSED)
#define ListView_FitContent (hCtl) SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, nCol, LVSCW_AUTOSIZE)
#define ListView_FitHeader (hCtl) SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, nCol, LVSCW_AUTOSIZE_USEHEADER)

' ========================================================================================
' Searches the listview for the specified string.
' Returns the index of the item if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION ListView_FindString (BYVAL hCtl AS HWND, BYVAL pwszText AS WSTRING PTR) AS LONG
   DIM lvf AS LVFINDINFOW : lvf.flags = LVFI_STRING : lvf.psz = pwszText
   RETURN SendMessageW(hCtl, LVM_FINDITEMW, -1, cast(LPARAM, cast(LVFINDINFOW PTR, @lvf)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Selects all the items.
' ========================================================================================
PRIVATE SUB ListView_SelectAllItems (BYVAL hCtl AS HWND)
   DIM lvi AS LVITEMW : lvi.mask = LVIF_STATE : lvi.state = -1 : lvi.stateMask = LVIS_SELECTED
   SendMessageW(hCtl, LVM_SETITEMSTATE, cast(WPARAM, -1), cast(LPARAM, cast(LVITEMW PTR, @lvi)))
END SUB
' ========================================================================================
' ========================================================================================
' Removes the HDS_BUTTONS style from the header control to give it a flat appearance.
' ========================================================================================
PRIVATE SUB ListView_MakeHeaderFlat (BYVAL hCtl AS HWND)
   DIM hLvHeader AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT SUB
   SetWindowLongPtrW hLvHeader, GWL_STYLE, GetWindowLongPtrW(hLvHeader, GWL_STYLE) XOR HDS_BUTTONS
   SetWindowPos GetParent(hCtl), NULL, 0, 0, 0, 0, SWP_NOZORDER OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_DRAWFRAME
END SUB
' ========================================================================================
' ========================================================================================
' Automatically sizes all the columns of a listview. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION ListView_AutoSizeColumns (BYVAL hCtl AS HWND) AS BOOLEAN
   DIM i AS LONG, nCount AS LONG, hLvHeader AS HWND
   hLvHeader = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT FUNCTION
   nCount = SendMessageW(hLvHeader, HDM_GETITEMCOUNT, 0, 0)
   FOR i = 0 TO nCount - 1
      FUNCTION = SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, cast(WPARAM, i), cast(LPARAM, LVSCW_AUTOSIZE))
   NEXT
END FUNCTION
' ========================================================================================
' ========================================================================================
' Automatically sizes all the columns of the header of a listview
' ========================================================================================
PRIVATE FUNCTION ListView_AutoSizeHeader (BYVAL hCtl AS HWND) AS BOOLEAN
   DIM i AS LONG, nCount AS LONG, hLvHeader AS HWND
   hLvHeader = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT FUNCTION
   nCount = SendMessageW(hLvHeader, HDM_GETITEMCOUNT, 0, 0)
   FOR i = 0 TO nCount - 1
      FUNCTION = SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, cast(WPARAM, i), cast(LPARAM, LVSCW_AUTOSIZE_USEHEADER))
   NEXT
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the number of columns of a listview.
' ========================================================================================
PRIVATE FUNCTION ListView_GetColumnCount (BYVAL hCtl AS HWND) AS LONG
   DIM hLvHeader AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT FUNCTION
   RETURN SendMessageW(hLvHeader, HDM_GETITEMCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the order of the specified column of the listview's header control.
' ========================================================================================
PRIVATE FUNCTION ListView_GetColumnOrder (BYVAL hCtl AS HWND, BYVAL nCol AS LONG) AS LONG
   DIM hLvHeader AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT FUNCTION
   DIM hdi AS HDITEMW
   hdi.mask = HDI_ORDER
   IF SendMessageW(hLvHeader, HDM_GETITEM, cast(WPARAM, nCol), cast(LPARAM, cast(HDITEMW PTR, @hdi))) THEN RETURN hdi.iOrder
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the current left-to-right order of items in the header of a listview control.
' ========================================================================================
PRIVATE FUNCTION ListView_GetOrderArray (BYVAL hCtl AS HWND, BYVAL nSize AS DWORD, BYVAL lpiArray AS LONG PTR) AS LONG
   DIM hLvHeader AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT FUNCTION
   RETURN SendMessageW(hLvHeader, HDM_GETORDERARRAY, cast(WPARAM, nSize), cast(LPARAM, lpiArray))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the font used by the ToolTip control that the list-view control uses to display ToolTips.
' The return value is a handle to the font used by the control, or NULL if the control is
' using the system font.
' ========================================================================================
PRIVATE FUNCTION ListView_GetTooltipsFont (BYVAL hCtl AS HWND) AS HFONT
   DIM hwndTooltips AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETTOOLTIPS, 0, 0))
   RETURN cast(HFONT, SendMessageW(hwndTooltips, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the font used by the ToolTip control that the list-view control uses to display ToolTips.
' Parameters:
' - hwndLV  = Handle to the ListView control.
' - hFont   = Handle to the font. If this parameter is NULL, the control uses the default system font to draw text.
' - fRedraw = Specifies whether the control should be redrawn immediately upon setting the font.
'             If this parameter is TRUE, the control redraws itself.
' ========================================================================================
PRIVATE SUB ListView_SetTooltipsFont (BYVAL hCtl AS HWND, BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
   DIM hwndTooltips AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETTOOLTIPS, 0, 0))
   IF hwndTooltips = NULL THEN EXIT SUB
   SendMessageW(hwndTooltips, WM_SETFONT, cast(WPARAM, hFont), cast(LPARAM, fRedraw))
END SUB
' ========================================================================================
' ========================================================================================
' Retrieves the text of the specified column. Returns TRUE or FALSE.
' - hwndLV: Handle to the listview control.
' - nCol: Zero-based index of the column
' - pwszText: A pointer to a buffer to receive the retrieved text.
' - cchTextMax : Maximum number of characters to return. Both this value and the size of the
'   buffer pointed by pwszText must be one character bigger that the wanted length of the text
'   to return to make room for the null character terminator.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION ListView_GetColumnText (BYVAL hCtl AS HWND, BYVAL nCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DIM lvc AS LVCOLUMNW : lvc.mask = LVCF_TEXT : lvc.cchTextMax = cchTextMax : lvc.pszText = pwszText
   RETURN SendMessageW(hCtl, LVM_GETCOLUMNW, cast(WPARAM, nCol), cast(LPARAM, cast(LVCOLUMNW PTR, @lvc)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text of a list-view item or subitem.
' - hwndLV: Handle to the ListView control.
' - nCol: Zero-based index of the column
' - pwszText: A pointer to a buffer that receives the item or subitem text.
' - cchTextMax : Maximum number of characters to return. Both this value and the size of the
'   buffer pointed by pwszText must be one character bigger that the wanted length of the text
'   to return to make room for the null character terminator.
' Returns the number of characters in the pszText member of the LVITEM structure.
' ========================================================================================
PRIVATE FUNCTION ListView_GetHeaderText (BYVAL hCtl AS HWND, BYVAL nCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS LONG
   DIM hLvHeader AS HWND = cast(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = NULL THEN EXIT FUNCTION
   DIM hdi AS HDITEMW : hdi.mask = HDI_TEXT : hdi.cchTextMax = cchTextMax : hdi.pszText = pwszText
   RETURN SendMessageW(hLvHeader, HDM_GETITEMW, cast(WPARAM, nCol), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new column in a list-view control.
' - hwndLV: Handle to the ListView control.
' - nCol: Index of the new column.
' - pwszText: The column header text.
' - nWidth: The width of the column, in pixels.
' - nFormat: Alignment of the column header and the subitem text in the column.
'   The alignment of the leftmost column is always left-justified; it cannot be changed.
'   This member can be one of the following values:
'   see https://msdn.microsoft.com/en-us/library/windows/desktop/bb774743%28v=vs.85%29.aspx
' Returns the index of the new column if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION ListView_AddColumn (BYVAL hCtl AS HWND, BYVAL nCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL nWidth AS LONG, BYVAL nFormat AS LONG = 0) AS LONG
   DIM lvc AS LVCOLUMNW
   lvc.mask = LVCF_FMT OR LVCF_WIDTH OR LVCF_TEXT OR LVCF_SUBITEM
   lvc.fmt = nFormat : lvc.pszText = pwszText : lvc.cx = nWidth : lvc.iSubItem = nCol
   RETURN SendMessageW(hCtl, LVM_INSERTCOLUMNW, cast(WPARAM, nCol), cast(LPARAM, cast(LVCOLUMNW PTR, @lvc)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new item in a list-view control.
' - hwndLV: Handle to the ListView control.
' - nItem: Zero-based index at which the new item should be inserted. If this value is
'   greater than the number of items currently contained by the listview control, the new
'   item will be appended to the end of the list and assigned the correct index. Examine the
'   return value to determine the actual index assigned to the item.
' - nImage: Zero-based index of the item's icon in the control's image list.
' - pwszText: The item text.
' Returns the index of the new item if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION ListView_AddItem (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL nImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DIM lvi AS LVITEMW
   lvi.mask = LVIF_TEXT : lvi.pszText = pwszText : lvi.iItem = nItem : lvi.iImage = nImage
   RETURN SendMessageW(hCtl, LVM_INSERTITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(LVITEMW PTR, @lvi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a new item in a list-view control.
' - hwndLV: Handle to the ListView control.
' - nImage: Zero-based index of the item's icon in the control's image list.
' - pwszText: The item text.
' Returns the index of the new item if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION ListView_AppendItem (BYVAL hCtl AS HWND, BYVAL nImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DIM lvi AS LVITEMW
   DIM nItem AS LONG = SendMessageW(hCtl, LVM_GETITEMCOUNT, 0, 0)
   lvi.mask = LVIF_TEXT : lvi.pszText = pwszText : lvi.iItem = nItem : lvi.iImage = nImage
   RETURN SendMessageW(hCtl, LVM_INSERTITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(LVITEMW PTR, @lvi)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                      *** REBAR ***
' ########################################################################################

#define Rebar_BeginDrag (hCtl) SendMessageW(hCtl, RB_BEGINDRAG, uBand, dwPos)
#define Rebar_DeleteBand (hCtl) SendMessageW(hCtl, RB_DELETEBAND, uBand, 0)
#define Rebar_DragMove (hCtl, dwPos) SendMessageW(hCtl, RB_DRAGMOVE, 0, dwPos)
#define Rebar_EndDrag (hCtl) SendMessageW(hCtl, RB_ENDDRAG, 0, 0)
#define Rebar_GetBandBorders (hCtl) SendMessageW(hCtl, RB_GETBANDBORDERS, uBand, cast(LPARAM, prc))
#define Rebar_GetBandCount (hCtl) SendMessageW(hCtl, RB_GETBANDCOUNT, 0, 0)
#define Rebar_GetBandMargins (hCtl, pMargins) SendMessageW(hCtl, RB_GETBANDMARGINS, 0, pMargins)
#define Rebar_GetBarHeight (hCtl) SendMessageW(hCtl, RB_GETBARHEIGHT, 0, 0)
#define Rebar_GetBarInfo (hCtl, prbi) SendMessageW(hCtl, RB_GETBARINFO, 0, cast(LPARAM, lprbi))
#define Rebar_GetBkColor (hCtl) SendMessageW(hCtl, RB_GETBKCOLOR, 0, 0)
#define Rebar_GetColorScheme (hCtl, pcs) SendMessageW(hCtl, RB_GETCOLORSCHEME, 0, cast(LPARAM, pcs))
#define Rebar_GetExtendedStyle (hCtl) SendMessageW(hCtl, RB_GETEXTENDEDSTYLE, 0, 0)
#define Rebar_GetPalette (hCtl) SendMessageW(hCtl, RB_GETPALETTE, 0, 0)
#define Rebar_GetRect (hCtl, uBand, prc) SendMessageW(hRebar, RB_GETRECT, uBand, cast(LPARAM, prc))
#define Rebar_GetRowCount (hCtl) SendMessageW(hCtl, RB_GETROWCOUNT, 0, 0)
#define Rebar_GetRowHeight (hCtl, uRow) SendMessageW(hCtl, RB_GETROWHEIGHT, uRow, 0)
#define Rebar_GetTextColor (hCtl) SendMessageW(hCtl, RB_GETTEXTCOLOR, 0, 0)
#define Rebar_GetTooltips (hCtl) SendMessageW(hCtl, RB_GETTOOLTIPS, 0, 0)
#define Rebar_GetUnicodeFormat (hCtl) SendMessageW(hCtl, RB_GETUNICODEFORMAT, 0, 0)
#define Rebar_HitTest (hCtl, prhbt) SendMessageW(hCtl, RB_HITTEST, 0, cast(LPARAM, prhbt))
#define Rebar_IdToIndex (hCtl, uBandID) SendMessageW(hCtl, RB_IDTOINDEX, uBandID, 0)
#define Rebar_InsertBand (hCtl, index, prbbi) SendMessageW(hCtl, RB_INSERTBAND, index, cast(LPARAM, prbbi))
#define Rebar_MaximizeBand (hCtl, uBand, fIdeal) SendMessageW(hCtl, RB_MAXIMIZEBAND, uBand, fIdeal)
#define Rebar_MinimizeBand (hCtl, uBand) SendMessageW(hCtl, RB_MAXIMIZEBAND, uBand, 0)
#define Rebar_MoveBand (hCtl, nFrom, nTo) SendMessageW(hCtl, RB_MOVEBAND, nFrom, nTo)
#define Rebar_PushChevron (hCtl, uBand, iAppValue) SendMessageW(hCtl, RB_PUSHCHEVRON, uBand, iAppValue)
#define Rebar_SetBandInfo (hCtl, uBand, prbbi) SendMessageW(hCtl, RB_SETBANDINFO, uBand, cast(LPARAM, prbbi))
#define Rebar_SetBarInfo (hCtl, prbbi) SendMessageW(hCtl, RB_SETBARINFO, 0, cast(LPARAM, prbbi))
#define Rebar_SetBkColor (hCtl, clrBk) SendMessageW(hCtl, RB_SETBKCOLOR, 0, clrBk)
#define Rebar_SetColorScheme (hCtl, pcs) SendMessageW(hCtl, RB_SETCOLORSCHEME, 0, cast(LPARAM, pcs))
#define Rebar_SetPalette (hCtl, hpal) SendMessageW(hCtl, RB_SETPALETTE, 0, cast(LPARAM, hpal))
#define Rebar_SetParent (hCtl, hwndParent) SendMessageW(hCtl, RB_SETPARENT, cast(WPARAM, hwndParent), 0))
#define Rebar_SetTextColor (hCtl, clrText) SendMessageW(hCtl, RB_SETTEXTCOLOR, 0, clrText)
#define Rebar_SetTooltips (hCtl, hwndToolTip) SendMessageW(hCtl, RB_SETTOOLTIPS, cast(WPARAM, hwndToolTip), 0))
#define Rebar_SetUnicodeFormat (hCtl, fUnicode) SendMessageW(hCtl, RB_SETUNICODEFORMAT, fUnicode, 0))
#define Rebar_SetWindowTheme (hCtl, pwszTheme) SendMessageW(hCtl, RB_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
#define Rebar_ShowBand (hCtl, uBand, fShow) SendMessageW(hCtl, RB_SHOWBAND, uBand, fShow)
#define Rebar_SizeToRect (hCtl, prc) SendMessageW(hCtl, RB_SIZETORECT, 0, cast(LPARAM, prc))
#if _WIN32_WINNT = &h0602
#define Rebar_SetBandWidth (hCtl, uBand, nWidth) SendMessageW(hCtl, RB_SETBANDWIDTH, uBand, nWidth)
#endif
' ========================================================================================

' #define Rebar_GetBandInfo (hCtl, prbbi) SendMessageW(hCtl, RB_GETBANDINFOW, uBand, prbbi)
' ========================================================================================
' Retrieves information about the specified band of the Rebar control
' ========================================================================================
PRIVATE FUNCTION Rebar_GetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS LONG) AS REBARBANDINFOW
   DIM rbbi AS REBARBANDINFOW
   IF AfxWindowsVersion >= 600 AND AfxComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   rbbi.fMask = RBBIM_BACKGROUND OR RBBIM_CHILD OR RBBIM_CHILDSIZE OR RBBIM_COLORS OR RBBIM_ID OR _
                RBBIM_HEADERSIZE OR RBBIM_IDEALSIZE OR RBBIM_IMAGE OR RBBIM_SIZE OR RBBIM_STYLE OR RBBIM_TEXT
   SendMessageW(hRebar, RB_GETBANDINFOW, uBand, CAST(LPARAM, @rbbi))
   RETURN rbbi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a rebar control's IDropTarget interface pointer.
' ========================================================================================
PRIVATE FUNCTION Rebar_GetDropTarget (BYVAL hRebar AS HWND) AS IDropTarget PTR
   DIM ppIDropTarget AS IDropTarget PTR
   SendMessageW(hRebar, RB_GETDROPTARGET, 0, CAST(LPARAM, @ppIDropTarget))
   FUNCTION = ppIDropTarget
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                     *** STATIC ***
' ########################################################################################

#define Static_DeleteBitmap (hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hCtl, STM_SETIMAGE, IMAGE_BITMAP, 0)))
#define Static_DeleteCursor (hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETIMAGE, IMAGE_CURSOR, 0)))
#define Static_DeleteEnhancedMetafile (hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETIMAGE, IMAGE_ENHMETAFILE, 0)))
#define Static_DeleteIcon (hCtl) DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETICON, 0, 0)))
#define Static_DeleteImage (hCtl, nType) DeleteObject(cast(HGDIOBJ, SendMessageW(hCtl, STM_SETIMAGE, nType, 0)))
#define Static_GetImage (hCtl, nType) cast(HANDLE, SendMessageW(hCtl, STM_GETIMAGE, nType, 0))
#define Static_SetBitmap (hCtl, hBitmap) cast(HBITMAP, SendMessageW(hCtl, STM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap)))
#define Static_SetCursor (hCtl, hCursor) cast(HCURSOR, SendMessageW(hStatic, STM_SETIMAGE, IMAGE_CURSOR, cast(LPARAM, hCursor)))
#define Static_SetEnhancedMetafile (hCtl, hEnhMetafile) cast(HENHMETAFILE, SendMessageW(hStatic, STM_SETIMAGE, IMAGE_ENHMETAFILE, cast(LPARAM, hEnhMetafile)))
#define Static_SetImage (hCtl, nType, hImage) cast(HANDLE, SendMessageW(hCtl, STM_SETIMAGE, cast(WPARAM, nType), cast(LPARAM, hImage)))


' ########################################################################################
'                                  *** PROGRESS BAR ***
' ########################################################################################

#define ProgressBar_DeltaPos (hCtl, nIncrement) SendMessageW(hCtl, PBM_DELTAPOS, nIncrement, 0)
#define ProgressBar_GetHighLimit (hCtl) SendMessageW(hCtl, PBM_GETRANGE, FALSE, 0)
#define ProgressBar_GetLowLimit (hCtl) SendMessageW(hCtl, PBM_GETRANGE, CTRUE, 0)
#define ProgressBar_GetMaxRange (hCtl) SendMessageW(hCtl, PBM_GETRANGE, FALSE, 0)
#define ProgressBar_GetMinRange (hCtl) SendMessageW(hCtl, PBM_GETRANGE, CTRUE, 0)
#define ProgressBar_GetPos (hCtl) SendMessageW(hCtl, PBM_GETPOS, 0, 0)
#define ProgressBar_GetRange (hCtl, fWhichLimit, ppBRange) SendMessageW(hCtl, PBM_GETRANGE, fWhichLimit, cast(LPARAM, ppBRange))

' ========================================================================================
' Advances the current position for a progress bar by the specified increment and redraws
' the bar to reflect the new position. nIncVaL: Signed integer that becomes the new position.
' If nIncVal is outside the range of the control, the position is set to the closest boundary.
' Do not send this message to a control that has the PBS_MARQUEE style.
' ========================================================================================
PRIVATE FUNCTION ProgressBar_IncPos (BYVAL hProgress AS HWND, BYVAL nIncVal AS LONG) AS LONG
   DIM nPos AS LONG = SendMessageW(hProgress, PBM_GETPOS, 0, 0)
   FUNCTION = SendMessageW(hProgress, PBM_SETPOS, cast(WPARAM, nPos + nIncVal), 0)
END FUNCTION
' ========================================================================================

#define ProgressBar_SetBarColor (hCtl, clrBar) SendMessageW(hCtl, PBM_SETBARCOLOR, 0, clrBar)
#define ProgressBar_SetBkColor (hCtl, clrBk) SendMessageW(hCtl, PBM_SETBKCOLOR, 0, clrBk)
#define ProgressBar_SetMarquee (hCtl, fOn, dwAnimationSpeed) SendMessageW(hCtl, PBM_SETMARQUEE, fOn, dwAnimationSpeed)
#define ProgressBar_SetMarqueeOff (hCtl) SendMessageW(hCtl, PBM_SETMARQUEE, FALSE, 0)
#define ProgressBar_SetMarqueeOn (hCtl, dwAnimationSpeed) SendMessageW(hCtl, PBM_SETMARQUEE, CTRUE, dwAnimationSpeed)
#define ProgressBar_SetPos (hCtl, nNewPos) SendMessageW(hCtl, PBM_SETPOS, CTRUE, nNewPos)
#define ProgressBar_SetRange (hCtl, nMinRange, nMaxRange) SendMessageW(hCtl, PBM_SETRANGE, 0, MAKELONG(nMinRange, nMaxRange))
#define ProgressBar_SetRange32 (hCtl, nMinRange, nMaxRange) SendMessageW(hCtl, PBM_SETRANGE32, nMinRange, nMaxRange)
#define ProgressBar_SetStep (hCtl, nStepInc) SendMessageW(hCtl, PBM_SETSTEP, nStepInc, 0)
#define ProgressBar_StepIt (hCtl) SendMessageW(hCtl, PBM_STEPIT, 0, 0)
#if _WIN32_WINNT = &h0602
#define ProgressBar_GetBarColor (hCtl) SendMessageW(hCtl, PBM_GETBARCOLOR, 0, 0)
#define ProgressBar_GetBkColor (hCtl) SendMessageW(hCtl, PBM_GETBKCOLOR, 0, 0)
#define ProgressBar_GetState (hCtl) SendMessageW(hCtl, PBM_GETSTATE, 0, 0)
#define ProgressBar_SetState (hCtl, iState) SendMessageW(hCtl, PBM_SETSTATE, iState, 0)
#endif


' ########################################################################################
'                                  *** STATUS BAR ***
' ########################################################################################

#define StatusBar_GetBorders (hCtl, rgBorders) SendMessageW(hCtl, SB_GETBORDERS, 0, cast(LPARAM, rgBorders))
#define StatusBar_GetIcon (hCtl, nPart) SendMessageW(hCtl, SB_GETICON, nPart, 0)
#define StatusBar_GetParts (hCtl, nParts, rgRightCoords) SendMessageW(hCtl, SB_GETPARTS, nParts, cast(LPARAM, rgRightCoords))
#define StatusBar_GetPartsCount (hCtl) SendMessageW(hCtl, SB_GETPARTS, 0, 0)
#define StatusBar_GetRect (hCtl, nPart, prc) SendMessageW(hCtl, SB_GETRECT, nPart, cast(LPARAM, prc))
#define StatusBar_GetTextDrawingMode (hCtl, nPart) SendMessageW(hCtl, SB_GETTEXTLENGTH, nPart, 0)
#define StatusBar_GetTextLength (hCtl, nPart) SendMessageW(hCtl, SB_GETTEXTLENGTH, nPart, 0)
#define StatusBar_GetTipText (hCtl, nPart, pwszText, wMaxLen) SendMessageW(hCtl, SB_GETTIPTEXTW, MAKELONG(nPart, wMaxLen), cast(LPARAM, pwszText))
#define StatusBar_GetUnicodeFormat (hCtl) SendMessageW(hCtl, SB_GETUNICODEFORMAT, 0, 0)
#define StatusBar_IsSimple (hCtl) SendMessageW(hCtl, SB_ISSIMPLE, 0, 0)
#define StatusBar_SetBkColor (hCtl, clrBk) SendMessageW(hCtl, SB_SETBKCOLOR, 0, clrBk)
#define StatusBar_SetFont (hCtl, hFont, fRedraw) SendMessageW(hCtl, WM_SETFONT, cast(WPARAM, hFont), fRedraw)
#define StatusBar_SetIcon (hCtl, nPart, hIcon) SendMessageW(hCtl, SB_SETICON, nPart, cast(LPARAM, hIcon))
#define StatusBar_SetMinHeight (hCtl, minHeight) SendMessageW(hCtl, SB_SETMINHEIGHT, minHeight, 0)
#define StatusBar_SetParts (hCtl, dwParts, rgWidths) SendMessageW(hCtl, SB_SETPARTS, dwParts, cast(LPARAM, rgWidths))
#define StatusBar_SetText (hCtl, nPart, pwszText, uType) SendMessageW(hCtl, SB_SETTEXTW, nPart OR uType, cast(LPARAM, pwszText))
#define StatusBar_SetTipText (hCtl, nPart, pwszTipText, uType) SendMessageW(hCtl, SB_SETTIPTEXTW, nPart, cast(LPARAM, pwszTipText))
#define StatusBar_SetUnicodeFormat (hCtl, fUnicode) SendMessageW(hCtl, SB_SETUNICODEFORMAT, fUnicode, 0)
#define StatusBar_Simple (hCtl, fSimple) SendMessageW(hCtl, SB_SIMPLE, 0, 0)
#define StatusBar_GetText (hCtl, nPart, pwszText) SendMessageW(hCtl, SB_GETTEXTW, nPart, cast(LPARAM, pwszText))

' ========================================================================================
' Returns the height of the status bar.
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetHeight (BYVAL hStatusBar AS HWND) AS LONG
   DIM rc AS RECT
   GetWindowRect(hStatusbar, @rc)
   FUNCTION = rc.Bottom - rc.Top
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                      *** TAB ***
' ########################################################################################

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
PRIVATE FUNCTION TabCtrl_InsertTab (BYVAL hwndTab AS HWND, BYVAL nTabIndex AS DWORD, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0) AS LONG
   IF IsWindow(hwndTab) = FALSE THEN RETURN -1
   IF pwszText = NULL THEN RETURN -1
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT OR TCIF_IMAGE OR TCIF_RTLREADING OR TCIF_PARAM OR TCIF_STATE OR TCIS_BUTTONPRESSED OR TCIS_HIGHLIGHTED
   ttc_item.pszText = pwszText
   ttc_item.cchTextMax = LEN(*pwszText)
   ttc_item.iImage = iImage
   ttc_item.lParam = lParam
   FUNCTION = SendMessageW(hwndTab, TCM_INSERTITEMW, cast(WPARAM, nTabIndex), cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of tabs and calls TabCtrl_InsertTab to add a new tab.
' ========================================================================================
PRIVATE FUNCTION TabCtrl_AddTab (BYVAL hwndTab AS HWND, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0) AS LONG
   IF IsWindow(hwndTab) = FALSE THEN RETURN -1
   IF pwszText = NULL THEN RETURN -1
   DIM nTabIndex AS LONG = SendMessageW(hwndTab, TCM_GETITEMCOUNT, 0, 0)
   FUNCTION = TabCtrl_InsertTab(hwndTab, nTabIndex, iImage, pwszText, lParam)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the name of a tab in a Tab control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TabCtrl_GetText (BYVAL hwndTab AS HWND, BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT
   ttc_item.pszText = pwszText
   ttc_item.cchTextMax = cchTextMax
   FUNCTION = SendMessageW(hwndTab, TCM_GETITEMW, cast(WPARAM, nTabIndex), cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the name of a tab in a Tab control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TabCtrl_SetText (BYVAL hwndTab AS HWND, BYVAL nTabIndex AS DWORD, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_TEXT
   ttc_item.pszText = pwszText
   ttc_item.cchTextMax = LEN(*pwszText)
   FUNCTION = SendMessageW(hwndTab, TCM_SETITEMW, cast(WPARAM, nTabIndex), cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the 0-based index in the tab control's image list.
' Returns the index of the image or -1 if not found.
' ========================================================================================
PRIVATE FUNCTION TabCtrl_GetImageIndex (BYVAL hwndTab AS HWND, BYVAL iItem AS DWORD) AS LONG
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_IMAGE
   FUNCTION = -1
   IF SendMessageW(hwndTab, TCM_GETITEMW, cast(WPARAM, iItem), cast(LPARAM, @ttc_item)) THEN
      FUNCTION = ttc_item.iImage
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the zero-based index in the tab control's image list.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TabCtrl_SetImageIndex (BYVAL hwndTab AS HWND, BYVAL iItem AS DWORD, BYVAL iImage AS LONG) AS BOOLEAN
   DIM ttc_item AS TCITEMW
   ttc_item.mask = TCIF_IMAGE
   ttc_item.iImage = iImage
   FUNCTION = SendMessageW(hwndTab, TCM_SETITEMW, cast(WPARAM, iItem), cast(LPARAM, @ttc_item))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                   *** TASK DIALOG ***
' ########################################################################################
#if _WIN32_WINNT = &h0602
#define TaskDialog_ClickButton (hCtl, nButton) SendMessageW(hCtl, TDM_CLICK_BUTTON, nButton, 0)
#define TaskDialog_ClickRadioButton (hCtl, nRadioButton) SendMessageW(hCtl, TDM_CLICK_RADIO_BUTTON, nRadioButton, 0)
#define TaskDialog_ClickVerification (hCtl, fChecked, fFocused) SendMessageW(hCtl, TDM_CLICK_VERIFICATION, fChecked, fFocused)
#define TaskDialog_EnableButton (hCtl, nButton, fEnable) SendMessageW(hCtl, TDM_ENABLE_BUTTON, nButton, fEnable)
#define TaskDialog_NavigatePage (hCtl, pTaskConfig) SendMessageW(hCtl, TDM_NAVIGATE_PAGE, 0, CAST(LPARAM, pTaskConfig))
#define TaskDialog_SetButtonElevationRequiredState (hCtl, nButton, fElevation) SendMessageW(hCtl, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, nButton, fElevation)
#define TaskDialog_SetElementText (hCtl, eElement, pszNewText) SendMessageW(hCtl, TDM_SET_ELEMENT_TEXT, eElement, cast(LPARAM, pszNewText))
#define TaskDialog_SetMarqueeProgressBar (hCtl, fMarquee) SendMessageW(hCtl, TDM_SET_MARQUEE_PROGRESS_BAR, fMarquee, 0)
#define TaskDialog_SetProgressBarMarquee (hCtl, fMarquee, nSpeed) SendMessageW(hCtl, TDM_SET_PROGRESS_BAR_MARQUEE, fMarquee, nSpeed)
#define TaskDialog_SetProgressBarPos (hCtl, nNewPos) SendMessageW(hCtl, TDM_SET_PROGRESS_BAR_POS, nNewPos, 0)
#define TaskDialog_SetProgressBarRange (hCtl, nMinRange, nMaxRange) SendMessageW(hCtl, TDM_SET_PROGRESS_BAR_RANGE, 0, MAKELONG(nMinRange, nMaxRange))
#define TaskDialog_SetProgressBarState (hCtl, nNewState) SendMessageW(hCtl, TDM_SET_PROGRESS_BAR_STATE, 0, nNewState)
#define TaskDialog_UpdateElementText (hCtl, eElement, pszNewText) SendMessageW(hCtl, TDM_UPDATE_ELEMENT_TEXT, eElement, cast(LPARAM, pszNewText))
#define TaskDialog_UpdateIcon (hCtl, nIconElement, idResource) SendMessageW(hCtl, TDM_UPDATE_ELEMENT_TEXT, nIconElement, idResource)
#endif


' ########################################################################################
'                                    *** TOOLBAR ***
' ########################################################################################

#define Toolbar_AddBitmap (hCtl, nButtons, tabm) SendMessageW(hCtl, TB_ADDBITMAP, nButtons, CAST(LPARAM, tabm))
#define Toolbar_AddButtons (hCtl, nButtons, pButtons) SendMessageW(hCtl, TB_ADDBUTTONSW, nButtons, CAST(LPARAM, pButtons))
#define Toolbar_AddString (hCtl, hInst, idString) SendMessageW(hCtl, TB_ADDSTRINGW, CAST(WPARAM, hInst), IIF(hInst = NULL, idString, MAKELONG(idString, 0)))
#define Toolbar_AutoSize (hCtl) SendMessageW(hCtl, TB_AUTOSIZE, 0, 0)
#define Toolbar_ButtonCount (hCtl) SendMessageW(hCtl, TB_BUTTONCOUNT, 0, 0)
#define Toolbar_ButtonStructSize (hCtl, cb) SendMessageW(hCtl, TB_BUTTONSTRUCTSIZE, cb, 0)
#define Toolbar_ChangeBitmap (hCtl, idButton, iBitmap) SendMessageW(hCtl, TB_CHANGEBITMAP, idButton, cast(LPARAM, MAKELONG(iBitmap, 0)))
#define Toolbar_CheckButton (hCtl, idButton, fCheck) SendMessageW(hCtl, TB_CHECKBUTTON, idButton, cast(LPARAM, MAKELONG(fCheck, 0)))
#define Toolbar_CommandToIndex (hCtl, idButton) SendMessageW(hCtl, TB_COMMANDTOINDEX, idButton, 0)
#define Toolbar_Customize (hCtl) SendMessageW(hCtl, TB_CUSTOMIZE, 0, 0)
#define Toolbar_DeleteButton (hCtl) SendMessageW(hCtl, TB_DELETEBUTTON, nButton, 0)
#define Toolbar_EnableButton (hCtl, idButton) SendMessageW(hCtl, TB_ENABLEBUTTON, idButton, MAKELONG(CTRUE, 0))
#define Toolbar_DisableButton (hCtl, idButton) SendMessageW(hCtl, TB_ENABLEBUTTON, idButton, MAKELONG(FALSE, 0))
#define Toolbar_GetAnchorHighlight (hCtl) SendMessageW(hCtl, TB_GETANCHORHIGHLIGHT, 0, 0)
#define Toolbar_GetBitmap (hCtl, idButton) SendMessageW(hCtl, TB_GETBITMAP, idButton, 0)
#define Toolbar_GetBitmapFlags (hCtl) SendMessageW(hCtl, TB_GETBITMAPFLAGS, 0, 0)
#define Toolbar_GetButton (hCtl, nButton, ptbb) SendMessageW(hCtl, TB_GETBUTTON, nButton, CAST(LPARAM, ptbb))
#define Toolbar_GetButtonSize (hCtl) SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0)
#define Toolbar_GetButtonHeight (hCtl) HIWORD(SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0))
#define Toolbar_GetButtonWidth (hCtl) LOWORD(SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0))
#define Toolbar_GetButtonInfo (hCtl, iID, ptbbi) SendMessageW(hCtl, TB_GETBUTTONINFOW, iID, cast(LPARAM, lptbbi))
#define Toolbar_GetColorScheme (hCtl, pcs) SendMessageW(hCtl, TB_GETCOLORSCHEME, 0, CAST(LPARAM, pcs))
#define Toolbar_GetImageList (hCtl) cast(HIMAGELIST, SendMessageW(hCtl, TB_GETIMAGELIST, 0, 0))
#define Toolbar_GetDisabledImageList (hCtl) cast(HIMAGELIST, SendMessageW(hCtl, TB_GETDISABLEDIMAGELIST, 0, 0))
#define Toolbar_GetHotImageList (hCtl) cast(HIMAGELIST, SendMessageW(hCtl, TB_GETHOTIMAGELIST, 0, 0))
#define Toolbar_GetExtendedStyle (hCtl) SendMessageW(hCtl, TB_GETEXTENDEDSTYLE, 0, 0)
#define Toolbar_GetHorizontalPadding (hCtl) LOWORD(SendMessageW(hCtl, TB_GETPADDING, 0, 0))
#define Toolbar_GetVerticalPadding (hCtl) HIWORD(SendMessageW(hCtl, TB_GETPADDING, 0, 0))
#define Toolbar_GetHotItem (hCtl) SendMessageW(hCtl, TB_GETHOTITEM, 0, 0)
#define Toolbar_GetInsertMark (hCtl, ptbim) TB_GETINSERTMARK(hCtl, TB_GETINSERTMARK, 0, cast(LPARAM, ptbim))
#define Toolbar_GetInsertMarkColor (hCtl) SendMessageW(hCtl, TB_GETINSERTMARKCOLOR, 0, 0)
#define Toolbar_GetItemRect (hCtl, nButton, prc) SendMessageW(hCtl, TB_GETITEMRECT, nButton, CAST(LPARAM, prc))
#define Toolbar_GetMaxSize (hCtl, psize) SendMessageW(hCtl, TB_GETMAXSIZE, 0, CAST(LPARAM, psize))
#define Toolbar_GetMetrics (hCtl, ptbMetrics) SendMessageW(hCtl, TB_GETMETRICS, 0, CAST(LPARAM, ptbMetrics))
#define Toolbar_GetPadding (hCtl) SendMessageW(hCtl, TB_GETPADDING, 0, 0)
#define Toolbar_GetRect (hCtl, iID, prc) SendMessageW(hCtl, TB_GETRECT, iID, CAST(LPARAM, prc))
#define Toolbar_GetRows (hCtl) SendMessageW(hCtl, TB_GETROWS, 0, 0)
#define Toolbar_GetState (hCtl, idButton) SendMessageW(hCtl, TB_GETSTATE, idButton, 0)
#define Toolbar_GetString (hCtl, iString, pBuffer) SendMessageW(hCtl, TB_GETSTRINGW, MAKELONG(cchMax, iString), CAST(LPARAM, pbuffer))
#define Toolbar_GetStyle (hCtl) SendMessageW(hCtl, TB_GETSTYLE, 0, 0)
#define Toolbar_GetTextRows (hCtl) SendMessageW(hCtl, TB_GETTEXTROWS, 0, 0)
#define Toolbar_GetTooltips (hCtl) SendMessageW(hCtl, TB_GETTOOLTIPS, 0, 0)
#define Toolbar_GetUnicodeFormat (hCtl) SendMessageW(hCtl, TB_GETUNICODEFORMAT, 0, 0)
#define Toolbar_HideButton (hCtl, idButton, fShow) SendMessageW(hCtl, TB_HIDEBUTTON, idButton, MAKELONG(fShow, 0))
#define Toolbar_HitTest (hCtl, ptHitTest) SendMessageW(hCtl, SendMessageW(hCtl, TB_HITTEST, 0, CAST(LPARAM, ptHitTest))
#define Toolbar_Indeterminate (hCtl, idButton, fIndeterminate) SendMessageW(hCtl, TB_INDETERMINATE, idButton, MAKELONG(fIndeterminate, 0))
#define Toolbar_InsertMarkHitTest (hCtl, ppt, ptbim) SendMessageW(hCtl, TB_INSERTMARKHITTEST, CAST(WPARAM, lppt), CAST(LPARAM, ptbim))
#define Toolbar_IsButtonChecked (hCtl, idButton) SendMessageW(hCtl, TB_ISBUTTONCHECKED, idButton, 0)
#define Toolbar_IsButtonEnabled (hCtl, idButton) SendMessageW(hCtl, TB_ISBUTTONENABLED, idButton, 0)
#define Toolbar_IsButtonHidden (hCtl, idButton) SendMessageW(hCtl, TB_ISBUTTONHIDDEN, idButton, 0)
#define Toolbar_IsButtonHighlighted (hCtl, idButton) SendMessageW(hCtl, TB_ISBUTTONHIGHLIGHTED, idButton, 0)
#define Toolbar_IsButtonIndeterminate (hCtl, idButton) SendMessageW(hCtl, TB_ISBUTTONINDETERMINATE, idButton, 0)
#define Toolbar_IsButtonPressed (hCtl, idButton) SendMessageW(hCtl, TB_ISBUTTONPRESSED, idButton, 0)
#define Toolbar_LoadImages (hCtl, iBitmapID, hInst) SendMessageW(hCtl, TB_LOADIMAGES, cast(WPARAM, iBitmapID), CAST(LPARAM, hinst))
#define Toolbar_MapAccelerator (hCtl, chAccel, pIDBtn) SendMessageW(hCtl, TB_MAPACCELERATOR, chAccel, CAST(LPARAM, pIDBtn))
#define Toolbar_MarkButton (hCtl, idButton, fHighlight) SendMessageW(hCtl, TB_MAPACCELERATOR, idButton, MAKELONG(fHighlight, 0))
#define Toolbar_MoveButton (hCtl, uOldPos, uNewPos) SendMessageW(hCtl, TB_MOVEBUTTON, uOldPos, uNewPos)
#define Toolbar_PressButton (hCtl, idButton, fPress) SendMessageW(hCtl, TB_PRESSBUTTON, idButton, MAKELONG(fPress, 0))
#define Toolbar_ReplaceBitmap (hCtl, ptbrb) SendMessageW(hCtl, TB_REPLACEBITMAP, 0, CAST(LPARAM, ptbrb))
#define Toolbar_SaveRestore (hCtl, fSave, ptbsp) SendMessageW(hCtl, TB_SAVERESTOREW, fSave, CAST(LPARAM, ptbsp))
#define Toolbar_SetAnchorHighlight (hCtl, fAnchor) SendMessageW(hCtl, TB_SETANCHORHIGHLIGHT, fAnchor, 0)
#define Toolbar_SetBitmapSize (hCtl, dxBitmap, dyBitmap) SendMessageW(hCtl, TB_SETBITMAPSIZE, 0, MAKELONG(dxBitmap, dyBitmap))
#define Toolbar_SetButtonInfo (hCtl, iID, ptbbi) SendMessageW(hCtl, TB_SETBUTTONINFOW, iID, cast(LPARAM, ptbbi))
#define Toolbar_SetButtonSize (hCtl, dxButton, dyButton) SendMessageW(hCtl, TB_SETBUTTONSIZE, 0, MAKELONG(dxButton, dyButton))
#define Toolbar_SetButtonWidth (hCtl, cxMin, cxMax) SendMessageW(hCtl, TB_SETBUTTONWIDTH, 0, MAKELONG(cxMin, cxMax))
#define Toolbar_SetCmdId (hCtl, index, cmdid) SendMessageW(hCtl, TB_SETCMDID, 0, MAKELONG(index, cmdid))
#define Toolbar_SetColorScheme (hCtl, pcs) SendMessageW(hCtl, TB_SETCOLORSCHEME, 0, CAST(LPARAM, lpcs))
#define Toolbar_SetImageList (hCtl, himlNew) SendMessageW(hCtl, TB_SETIMAGELIST, 0, CAST(LPARAM, himlNew))
#define Toolbar_SetHotImageList (hCtl, himlNewHot) cast(HIMAGELIST, SendMessageW(hCtl, TB_SETHOTIMAGELIST, 0, CAST(LPARAM, himlNewHot)))
#define Toolbar_SetDisabledImageList (hCtl, himlNewDisabled) cast(HIMAGELIST, SendMessageW(hCtl, TB_SETDISABLEDIMAGELIST, 0, CAST(LPARAM, himlNewDisabled)))
#define Toolbar_SetPressedImageList (hCtl, iImageID, himlNew) cast(HIMAGELIST, SendMessagew(hCtl, TB_SETPRESSEDIMAGELIST, iImageID, CAST(LPARAM, himlNew)))
#define Toolbar_SetDrawTextFlags (hCtl, dwMask, dwDTFlags) SendMessageW(hCtl, TB_SETDRAWTEXTFLAGS, dwMask, dwDTFlags)
#define Toolbar_SetExtendedStyle (hCtl, dwMask, dwDTFlags) SendMessageW(hCtl, TB_SETEXTENDEDSTYLE, 0, dwExStyle)
#define Toolbar_SetHotItem (hCtl, iHot) SendMessageW(hCtl, TB_SETHOTITEM, iHot, 0)
#define Toolbar_SetIndent (hCtl, iIndent) SendMessageW(hCtl, TB_SETINDENT, iIndent, 0)
#define Toolbar_SetInsertMark (hCtl, ptbim) SendMessageW(hCtl, TB_SETINSERTMARK, 0, cast(LPARAM, ptbim))
#define Toolbar_SetInsertMarkColor (hCtl, ptbim) SendMessageW(hCtl, TB_SETINSERTMARKCOLOR, clrInsertMark, 0)
#define Toolbar_SetListGap (hCtl, gap) SendMessageW(hCtl, TB_SETLISTGAP, gap, 0)
#define Toolbar_SetMaxTextRows (hCtl, iMaxRows) SendMessageW(hCtl, TB_SETMAXTEXTROWS, iMaxRows, 0)
#define Toolbar_SetMetrics (hCtl, ptbMetrics) SendMessageW(hCtl, TB_SETMETRICS, CAST(LPARAM, ptbmetrics))
#define Toolbar_SetPadding (hCtl, cx, cy) SendMessageW(hCtl, TB_SETPADDING, 0, MAKELONG(cx, cy))
#define Toolbar_SetParent (hCtl, hwndParent) SendMessageW(hCtl, TB_SETPARENT, cast(WPARAM, hwndParent), 0))
#define Toolbar_SetRows (hCtl, cRows, fLarger, prc) SendMessageW(hCtl, TB_SETROWS, MAKELONG(cRows, fLarger), CAST(LPARAM, prc))
#define Toolbar_SetState (hCtl, idButton, fState) SendMessageW(hCtl, TB_SETSTATE, idButton, MAKELONG(fState, 0))
#define Toolbar_SetStyle (hCtl, dwStyle) SendMessageW(hCtl, TB_SETSTYLE, 0, dwStyle)
#define Toolbar_SetToolTips (hCtl, hwndToolTip) SendMessageW(hCtl, TB_SETTOOLTIPS, CAST(WPARAM, hwndToolTip), 0)
#define Toolbar_SetUnicodeFormat (hCtl, fUnicode) SendMessageW(hCtl, TB_SETUNICODEFORMAT, fUnicode, 0)
#define Toolbar_SetWindowTheme (hCtl, pwszTheme) SendMessageW(hCtl, TB_SETWINDOWTHEME, CAST(WPARAM, pwszTheme), 0)
#define Toolbar_MakeOpaque (hCtl) SetWindowLongPtrW(hCtl, GWL_STYLE, GetWindowLongPtrW(hToolBar, GWL_STYLE) AND NOT TBSTYLE_TRANSPARENT)
#define Toolbar_GetButtonText (hCtl, idButton, pBuffer) SendMessageW(hCtl, TB_GETBUTTONTEXTW, idButton, CAST(LPARAM, pbuffer))
#define Toolbar_ButtonGetTextLength (hCtl, idButton) SendMessageW(hCtl, TB_GETBUTTONTEXTW, idButton, NULL)

' ========================================================================================
' Adds a button to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION Toolbar_AddButton (BYVAL hToolBar AS HWND, BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
   IF fsState = 0 THEN fsState = TBSTATE_ENABLED
   DIM idxString AS INT_PTR
   IF pwszText <> NULL THEN idxString = IIF(LEN(*pwszText) = 0, -1, CAST(INT_PTR, pwszText))
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0, 0, 0, 0, 0}, dwData, idxString)
#else
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0}, dwData, idxString)
#endif
   FUNCTION = SendMessageW(hToolBar, TB_ADDBUTTONSW, 1, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a separator to a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION Toolbar_AddSeparator (BYVAL hToolBar AS HWND, BYVAL nWidth AS LONG = 0) AS BOOLEAN
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0, 0, 0, 0, 0}, 0, -1)
#else
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0}, 0, -1)
#endif
   FUNCTION = SendMessageW(hToolBar, TB_ADDBUTTONSW, 1, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION Toolbar_InsertButton OVERLOAD (BYVAL hToolBar AS HWND, BYVAL idxButton AS LONG, BYVAL lpButton AS TBBUTTON PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_INSERTBUTTONW, cast(WPARAM, idxButton), CAST(LPARAM, lpButton))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_InsertButton OVERLOAD (BYVAL hToolBar AS HWND, BYVAL idxButton AS LONG, BYVAL idxBitmap AS LONG, BYVAL idCommand AS LONG, _
BYVAL fsState AS UBYTE = 0, BYVAL fsStyle AS UBYTE = 0, BYVAL dwData AS DWORD_PTR = 0, BYVAL pwszText AS WSTRING PTR = NULL) AS BOOLEAN
   IF fsState = 0 THEN fsState = TBSTATE_ENABLED
   DIM idxString AS INT_PTR
   IF pwszText <> NULL THEN idxString = IIF(LEN(*pwszText) = 0, -1, CAST(INT_PTR, pwszText))
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0, 0, 0, 0, 0}, dwData, idxString)
#else
   DIM tbb AS TBBUTTON = (idxBitmap, idCommand, fsState, fsStyle, {0, 0}, dwData, idxString)
#endif
   FUNCTION = SendMessageW(hToolBar, TB_INSERTBUTTONW, idxButton, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a separator in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION Toolbar_InsertSeparator (BYVAL hToolBar AS HWND, BYVAL idxButton AS LONG, BYVAL nWidth AS LONG = 0) AS BOOLEAN
#ifdef __FB_64BIT__
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0, 0, 0, 0, 0}, 0, -1)
#else
   DIM tbb AS TBBUTTON = (nWidth, 0, TBSTATE_ENABLED, TBSTYLE_SEP, {0, 0}, 0, -1)
#endif
   FUNCTION = SendMessageW(hToolBar, TB_INSERTBUTTONW, idxButton, CAST(LPARAM, @tbb))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the IDropTarget for a toolbar control.
' Minimum operating systems Windows 2000, Windows NT 4.0 with Internet Explorer 4.0,
' Windows 98, Windows 95 with Internet Explorer 4.0
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetObject (BYVAL hToolBar AS HWND) AS IDropTarget PTR
   DIM ppvObject AS IDropTarget PTR
   DIM IID_IDropTarget AS GUID = (&h00000122, &h0000, &h0000, {&hC0, &h00, &h00, &h00, &h00, &h00, &h00, &h46})
   DIM hr AS HRESULT = SendMessageW(hToolBar, TB_GETOBJECT, CAST(WPARAM, @IID_IDropTarget), CAST(LPARAM, ppvObject))
   IF hr <> S_OK THEN FUNCTION = ppvObject
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                     *** TRACKBAR ***
' ########################################################################################

#define Trackbar_ClearSel (hCtl, fRedraw) SendMessageW(hCtl, TBM_CLEARSEL, fRedraw, 0)
#define Trackbar_ClearTics (hCtl, fRedraw) SendMessageW(hCtl, TBM_CLEARTICS, fRedraw, 0)
#define Trackbar_GetBuddy (hCtl, fLocation) SendMessageW(hCtl, TBM_GETBUDDY, fLocation, 0)
#define Trackbar_GetChannelRect (hCtl, prc) SendMessageW(hCtl, TBM_GETCHANNELRECT, 0, cast(LPARAM, prc))
#define Trackbar_GetLineSize (hCtl) SendMessageW(hCtl, TBM_GETLINESIZE, 0, 0)
#define Trackbar_GetNumTics (hCtl) SendMessageW(hCtl, TBM_GETNUMTICS, 0, 0)
#define Trackbar_GetPageSize (hCtl) SendMessageW(hCtl, TBM_GETPAGESIZE, 0, 0)
#define Trackbar_GetPos (hCtl) SendMessageW(hCtl, TBM_GETPOS, 0, 0)
#define Trackbar_GetPTics (hCtl) SendMessageW(hCtl, TBM_GETPTICS, 0, 0)
#define Trackbar_GetRangeMax (hCtl) SendMessageW(hCtl, TBM_GETRANGEMAX, 0, 0)
#define Trackbar_GetRangeMin (hCtl) SendMessageW(hCtl, TBM_GETRANGEMIN, 0, 0)
#define Trackbar_GetSelEnd (hCtl) SendMessageW(hCtl, TBM_GETSELEND, 0, 0)
#define Trackbar_GetSelStart (hCtl) SendMessageW(hCtl, TBM_GETSELSTART, 0, 0)
#define Trackbar_GetThumbLength (hCtl) SendMessageW(hCtl, TBM_GETTHUMBLENGTH, 0, 0)
#define Trackbar_GetThumbRect (hCtl, prc) SendMessageW(hCtl, TBM_GETTHUMBRECT, 0, cast(LPARAM, prc))
#define Trackbar_GetTic (hCtl, iTic) SendMessageW(hCtl, TBM_GETTIC, iTic, 0)
#define Trackbar_GetTicPos (hCtl, iTic) SendMessageW(hCtl, TBM_GETTICPOS, iTic, 0)
#define Trackbar_GetTooltips (hCtl) SendMessageW(hCtl, TBM_GETTOOLTIPS, 0, 0)
#define Trackbar_GetUnicodeFormat (hCtl) SendMessageW(hCtl, TBM_GETUNICODEFORMAT, 0, 0)
#define Trackbar_SetBuddy (hCtl, fLocation, hwndBuddy) SendMessageW(hCtl, TBM_SETBUDDY, fLocation, cast(LPARAM, hwndBuddy))
#define Trackbar_SetLineSize (hCtl, lLineSize) SendMessageW(hCtl, TBM_SETLINESIZE, 0, lLineSize)
#define Trackbar_SetPageSize (hCtl, lPageSize) SendMessageW(hCtl, TBM_SETPAGESIZE, 0, lPageSize)
#define Trackbar_SetPos (hCtl, fPosition, lPosition) SendMessageW(hCtl, TBM_SETPOS, fPosition, lPosition)
#define Trackbar_SetRange (hCtl, lMinimum, lMaximum, fRedraw) SendMessageW(hCtl, TBM_SETRANGE, fRedraw, MAKELONG(lMinimum, lMaximum))
#define Trackbar_SetRangeMax (hCtl, lMaximum, fRedraw) SendMessageW(hCtl, TBM_SETRANGEMAX, fRedraw, lMaximum)
#define Trackbar_SetRangeMin (hCtl, lMinimum, fRedraw) SendMessageW(hCtl, TBM_SETRANGEMIN, fRedraw, lMinimum)
#define Trackbar_SetSel (hCtl, lMinimum, lMaximum, fRedraw) SendMessageW(hCtl, TBM_SETSEL, fRedraw, MAKELONG(lMinimum, lMaximum))
#define Trackbar_SetSelEnd (hCtl, lEnd, fRedraw) SendMessageW(hCtl, TBM_SETSELEND, fRedraw, lEnd)
#define Trackbar_SetSelStart (hCtl, lStart, fRedraw) SendMessageW(hCtl, TBM_SETSELSTART, fRedraw, lStart)
#define Trackbar_SetThumbLength (hCtl, iLength) SendMessageW(hCtl, TBM_SETTHUMBLENGTH, iLength, 0)
#define Trackbar_SetTic (hCtl, lPosition) SendMessageW(hCtl, TBM_SETTIC, 0, lPosition)
#define Trackbar_SetTicFreq (hCtl, wFreq) SendMessageW(hCtl, TBM_SETTICFREQ, wFreq, 0)
#define Trackbar_SetTipSide (hCtl, fLocation) SendMessageW(hCtl, TBM_SETTIPSIDE, fLocation, 0)
#define Trackbar_SetToolTips (hCtl, hwndTT) SendMessageW(hCtl, TBM_SETTOOLTIPS, cast(WPARAM, hwndTT), 0)
#define Trackbar_SetUnicodeFormat (hCtl, fUnicode) SendMessageW(hCtl, TBM_SETUNICODEFORMAT, fUnicode, 0)


' ########################################################################################
'                                   *** PROPSHEET ***
' ########################################################################################

#if _WIN32_WINNT = &h0602

'PSM_SETNEXTTEXTW         = (WM_USER + 137)
'PSM_SETNEXTTEXT          = PSM_SETNEXTTEXTW
#define PropSheet_SetNextText(hDlg, pwszText) SendMessage(hDlg, WM_USER + 137, 0, cast(LPARAM, pwszText))

'PSWIZB_SHOW                    = 0
'PSWIZB_RESTORE                 = 1
'PSM_SHOWWIZBUTTONS             = (WM_USER + 138)
#define PropSheet_ShowWizButtons(hDlg, dwFlag, dwButton) SendMessage(hDlg, WM_USER + 138, dwFlag, dwButton)

'PSM_ENABLEWIZBUTTONS           = (WM_USER + 139)
#define PropSheet_EnableWizButtons(hDlg, dwState, dwMask) SendMessage(hDlg, WM_USER + 139, dwState, dwMask)

'PSM_SETBUTTONTEXTW         = (WM_USER + 140)
'PSM_SETBUTTONTEXT          = PSM_SETBUTTONTEXTW
#define PropSheet_SetButtonText(hDlg, dwButton, pwszText) SendMessage(hDlg, WM_USER + 140, dwButton, cast(LPARAM, pwszText))

#endif


' ########################################################################################
'                                    *** SCROLLBAR ***
' ########################################################################################

#define ScrollBar_SetRangeRedraw (hCtl) SendMessageW(hCtl, SBM_SETRANGEREDRAW, posMin, posMax)


' ########################################################################################
'                                    *** SYSLINK ***
' ########################################################################################

#define SysLink_GetIdealHeight (hCtl) SendMessageW(hCtl, LM_GETIDEALHEIGHT, 0, 0)
#define Syslink_GetItem (hCtl, pItem) SendMessageW(hCtl, LM_GETITEM, 0, cast(LPARAM, pItem))
#define SysLink_HitTest (hCtl, phti) SendMessageW(hCtl, LM_HITTEST, 0, cast(LPARAM, phti))
#define SysLink_SetItem (hCtl, pItem) SendMessageW(hCtl, LM_SETITEM, 0, cast(LPARAM, pItem))


' ########################################################################################
'                                   *** TOOLTIPS ***
' ########################################################################################

#define Tooltip_Activate (hCtl, fActivate) SendMessageW(hCtl, TTM_ACTIVATE, fActivate, 0)
#define Tooltip_AddTool (hCtl, pti) SendMessageW(hCtl, TTM_ADDTOOLW, 0, CAST(LPARAM, pti))
#define Tooltip_AdjustRect (hCtl, fLarger, prc) SendMessageW(hCtl, TTM_ADJUSTRECT, fLarger, cast(LPARAM, prc))
#define Tooltip_DelTool (hCtl, pti) SendMessageW(hCtl, TTM_DELTOOLW, 0, cast(LPARAM, lpti))
#define Tooltip_EnumTools (hCtl, iTool, pti) SendMessageW(hCtl, TTM_ENUMTOOLSW, iTool, cast(LPARAM, pti))
#define Tooltip_GetBubbleSize (hCtl, pTtm) SendMessageW(hCtl, TTM_GETBUBBLESIZE, 0, cast(LPARAM, pTtm))
#define Tooltip_GetCurrentTool (hCtl, pti) SendMessageW(hCtl, TTM_GETCURRENTTOOLW, 0, cast(LPARAM, pti))
#define Tooltip_GetDelayTime (hCtl, dwDuration) SendMessageW(hCtl, TTM_GETDELAYTIME, dwDuration, 0)
#define Tooltip_GetFont (hCtl) SendMessageW(hCtl, WM_GETFONT, 0, 0)
#define Tooltip_GetMargin (hCtl, prc) SendMessageW(hCtl, TTM_GETMARGIN, 0, cast(LPARAM, prc))
#define Tooltip_GetMaxTipWidth (hCtl) SendMessageW(hCtl, TTM_GETMAXTIPWIDTH, 0, 0)
#define Tooltip_GetText (hCtl, numChars, pti) SendMessageW(hCtl, TTM_GETTEXTW, numChars, cast(LPARAM, pti))
#define Tooltip_GetTipBkColor (hCtl) SendMessageW(hCtl, TTM_GETTIPBKCOLOR, 0, 0)
#define Tooltip_GetTipTextColor (hCtl) SendMessageW(hCtl, TTM_GETTIPTEXTCOLOR, 0, 0)
#define Tooltip_GetTitle (hCtl, pTTGetTitle) SendMessageW(hCtl, TTM_GETTITLE, 0, cast(LPARAM, pTTGetTitle))
#define Tooltip_GetToolCount (hCtl) SendMessageW(hCtl, TTM_GETTOOLCOUNT, 0, 0)
#define Tooltip_GetToolInfo (hCtl, pti) SendMessageW(hCtl, TTM_GETTOOLINFOW, 0, cast(LPARAM, pti))
#define Tooltip_HitTest (hCtl, phti) SendMessageW(hCtl, TTM_HITTEST, 0, cast(LPARAM, phti))
#define Tooltip_NewToolRect (hCtl, pti) SendMessageW(hCtl, TTM_NEWTOOLRECTW, 0, cast(LPARAM, pti))
#define Tooltip_Pop (hCtl) SendMessageW(hCtl, TTM_POP, 0, 0)
#define Tooltip_Popup (hCtl) SendMessageW(hCtl, TTM_POPUP, 0, 0)
#define Tooltip_RelayEvent (hCtl, pmsg) SendMessageW(hCtl, TTM_RELAYEVENT, 0, cast(LPARAM, pmsg))
#define Tooltip_SetDelayTime (hCtl, dwDuration, iTime) SendMessageW(hCtl, TTM_SETDELAYTIME, dwDuration, MAKELONG(iTime, 0))
#define Tooltip_SetFont (hCtl, hFont, fRedraw) SendMessageW(hCtl, WM_SETFONT, cast(WPARAM, hFont), cast(LPARAM, fRedraw))
#define Tooltip_SetMargin (hCtl, prc) SendMessageW(hCtl, TTM_SETMARGIN, 0, cast(LPARAM, prc))
#define Tooltip_SetMaxTipWidth (hCtl, iWidth) SendMessageW(hCtl, TTM_SETMAXTIPWIDTH, 0, iWidth)
#define Tooltip_SetTipBkColor (hCtl, clr) SendMessageW(hCtl, TTM_SETTIPBKCOLOR, clr, 0)
#define Tooltip_SetTipTextColor (hCtl, clr) SendMessageW(hCtl, TTM_SETTIPTEXTCOLOR, clr, 0)
#define Tooltip_SetTitle (hCtl, nIcon, pwszTitle) SendMessageW(hCtl, TTM_SETTITLEW, nIcon, cast(LPARAM, pwszTitle))
#define Tooltip_SetToolInfo (hCtl, pti) SendMessageW(hCtl, TTM_SETTOOLINFOW, 0, cast(LPARAM, pti))
#define Tooltip_SetWindowTheme (hCtl, pwszTheme) SendMessageW(hCtl, TTM_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
#define Tooltip_TrackActivate (hCtl, fActivate, pti) SendMessageW(hCtl, TTM_TRACKACTIVATE, fActivate, cast(LPARAM, pti))
#define Tooltip_TrackPosition (hCtl, xPos, yPos) SendMessageW(hCtl, TTM_TRACKPOSITION, 0, MAKELONG(xPos, yPos))
#define Tooltip_Update (hCtl) SendMessageW(hCtl, TTM_UPDATE, 0, 0)
#define Tooltip_UpdateTipText (hCtl, pti) SendMessageW(hCtl, TTM_UPDATETIPTEXTW, 0, cast(LPARAM, pti))
#define Tooltip_WindowFromPoint (hCtl, ppt) SendMessageW(hCtl, TTM_WINDOWFROMPOINT, 0, cast(LPARAM, ppt))

' ========================================================================================
' Creates a tooltip for a control.
' Parameters:
' - hwnd      = Handle of the window or control
' - wszText   = Tooltip text
' - bBalloon  = Ballon tip (TRUE or FALSE)
' - bCentered = Centered (TRUE or FALSE)
' Return Value:
'   The handle of the tooltip control
' ========================================================================================
PRIVATE FUNCTION Tooltip_Add (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING = "", BYVAL bBalloon AS BOOLEAN = FALSE, BYVAL bCentered AS BOOLEAN = FALSE) AS HWND
   IF IsWindow(hwnd) = 0 THEN EXIT FUNCTION
   ' // Creates the tooltip control
   DIM dwStyle AS DWORD = WS_POPUP OR TTS_ALWAYSTIP
   IF bBalloon THEN dwStyle = dwStyle OR TTS_BALLOON
   DIM hTooltip AS HWND
   hTooltip = CreateWindowExW(0, "tooltips_class32", "", dwStyle, 0, 0, 0, 0, NULL, NULL, NULL, NULL)
   IF hTooltip = NULL THEN EXIT FUNCTION
   ' // You must explicitly define a tooltip control as topmost. Otherwise, it might be covered by the parent window.
   SetWindowPos(hTooltip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)
   ' // Registers the window with the tooltip control
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF AfxComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF AfxComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.uFlags = TTF_IDISHWND OR TTF_SUBCLASS
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.uFlags = TTF_SUBCLASS
      tti.hwnd = hwnd
      tti.uId = 0
      GetClientRect(hwnd, @tti.rect)
   END IF
   IF bCentered THEN tti.uFlags = tti.uFlags OR TTF_CENTERTIP
   tti.hinst = GetModuleHandleW(NULL)
   tti.lpszText = CAST(LPWSTR, @wszText)
   SendMessageW(hTooltip, TTM_ADDTOOLW, 0, CAST(LPARAM, @tti))
   FUNCTION = hTooltip
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets/replaces the text of a tooltip control
' Parameters:
' - hTooltip = Handle of the tooltip control
' - hwnd     = Handle of the window or control
' - wszText  = Tooltip text
' ========================================================================================
PRIVATE SUB Tooltip_SetText (BYVAL hTooltip AS HWND, BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING)
   IF hTooltip = NULL OR hwnd = NULL THEN EXIT SUB
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF AfxComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF AfxComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.hwnd = hwnd
      tti.uId = 0
   END IF
   ' // Retrieve the tooltip information
   SendMessageW(hTooltip, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti))
   IF SendMessageW(hTooltip, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti)) THEN
      tti.lpszText = CAST(LPWSTR, @wszText)
      SendMessageW(hTooltip, TTM_SETTOOLINFOW, 0, CAST(LPARAM, @tti))
   END IF
   ' // Note: Windows provides the TTM_UPDATETIPTEXT message, but needs Vista or superior.
END SUB
' ========================================================================================

' ========================================================================================
' Removes a tool from a tooltip control.
' Parameters:
' - hTooltip = Handle of the tooltip control
' - hwnd     = Handle of the window or control
' ========================================================================================
PRIVATE SUB Tooltip_Delete (BYVAL hTooltip AS HWND, BYVAL hwnd AS HWND)
   IF hTooltip = NULL OR hwnd = NULL THEN EXIT SUB
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF AfxComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF AfxComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.hwnd = hwnd
      tti.uId = 0
   END IF
   ' // Remove the tooltip
   SendMessageW(hTooltip, TTM_DELTOOL, 0, CAST(LPARAM, @tti))
END SUB
' ========================================================================================


' ########################################################################################
'                                   *** TREEVIEW ***
' ########################################################################################

' ========================================================================================
' Retrieves some or all of a tree-view item's attributes.
' Returns TRUE if successful, or FALSE otherwise.
' Usage example:
' DIM item AS TVITEMEXW
' item.mask = TVIF_TEXT
' item.hitem = hitem
' DIM wszText AS WSTRING * MAX_PATH
' item.pszText = @wszText
' item.cchTextMax = MAX_PATH
' TreeView_GetItemEx(hTreeView, @item)
' ========================================================================================
PRIVATE FUNCTION TreeView_GetItemEx (BYVAL hwndTV AS HWND, BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hwndTV, TVM_GETITEMW, 0, cast(LPARAM, pitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new item in a tree-view control.
' Note: It should have been named TreeView_InsertItem, but this name is already been used by a macro.
' ========================================================================================
PRIVATE FUNCTION TreeView_AddItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, BYVAL hInsertAfter AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   DIM tvinsert AS TVINSERTSTRUCTW
   tvinsert.hParent             = hParent
   tvinsert.hInsertAfter        = hInsertAfter
   tvinsert.Item.iImage         = iImage
   tvinsert.Item.iSelectedImage = iSelectedImage
   tvinsert.Item.lParam         = lParam
   tvinsert.Item.mask           = TVIF_TEXT OR TVIF_IMAGE OR TVIF_SELECTEDIMAGE OR TVIF_PARAM
   tvinsert.Item.pszText        = pwszText
   IF pwszText <> LPSTR_TEXTCALLBACK THEN
      tvinsert.Item.cchTextMax  = LEN(*pwszText)
   END IF
   FUNCTION = cast(HTREEITEM, SendMessageW(hwndTV, TVM_INSERTITEMW, 0, cast(LPARAM, cast(TVINSERTSTRUCTW PTR, @tvinsert))))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts the item at the end of the list in a tree-view control.
' ========================================================================================
PRIVATE FUNCTION TreeView_AppendItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   FUNCTION = TreeView_AddItem(hwndTV, hParent, TVI_LAST, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts the item at the beginning of the list in a tree-view control.
' ========================================================================================
PRIVATE FUNCTION TreeView_InsertFirstItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   FUNCTION = TreeView_AddItem(hwndTV, hParent, TVI_FIRST, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Add the item as the root item in a tree-view control.
' ========================================================================================
PRIVATE FUNCTION TreeView_AddRootItem (BYVAL hwndTV AS HWND, BYVAL pwszText AS WSTRING PTR, _
BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   FUNCTION = TreeView_AddItem(hwndTV, NULL, TVI_ROOT, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts the item into the list in alphabetical order.
' ========================================================================================
PRIVATE FUNCTION TreeView_InsertSortedItem (BYVAL hwndTV AS HWND, BYVAL hParent AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   FUNCTION = TreeView_AddItem(hwndTV, hParent, TVI_SORT, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the bold attribute is set, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TreeView_GetBold (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM) AS BOOLEAN
   FUNCTION = ((TreeView_GetItemState(hwndTV, hItem, TVIS_BOLD) AND TVIS_BOLD) = TVIS_BOLD)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the item is checked, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TreeView_IsItemChecked (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM) AS BOOLEAN
   FUNCTION = (TreeView_GetCheckState(hwndTV, hItem) = 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of characters in the incremental search string.
' If the tree-view control is not in incremental search mode, the return value is zero.
' ========================================================================================
PRIVATE FUNCTION TreeView_GetISearchStringLen (BYVAL hwndTV AS HWND) AS LONG
   FUNCTION = SendMessage(hwndTV, TVM_GETISEARCHSTRING, 0, CAST(LPARAM, NULL))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieve the lParam value from a Treeview
' ========================================================================================
PRIVATE FUNCTION TreeView_GetlParam (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM) AS LPARAM
   DIM tvi AS TVITEMW
   tvi.hItem = hItem
   tvi.mask  = TVIF_HANDLE OR TVIF_PARAM
   TreeView_GetItem(hwndTV, @tvi)
   FUNCTION = tvi.lParam
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets/unsets the bold attribute of the specified item.
' ========================================================================================
PRIVATE SUB TreeView_SetBold (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL fBold AS BOOLEAN)
   TreeView_SetItemState(hwndTV, hItem, IIF(fBold = 0, 0, TVIS_BOLD), TVIS_BOLD)
END SUB
' ========================================================================================

' ========================================================================================
' Sets some or all of a tree-view item's attributes.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TreeView_SetItemEx (BYVAL hwndTV AS HWND, BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hwndTV, TVM_SETITEMW, 0, cast(WPARAM, pitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified item.
' ========================================================================================
PRIVATE FUNCTION TreeView_SetItemText (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM tvi AS TVITEMW
   tvi.hitem = hItem
   tvi.mask = TVIF_TEXT
   tvi.psztext = pwszText
   tvi.cchtextmax = LEN(*pwszText)
   FUNCTION = SendMessageW(hwndTV, TVM_SETITEMW, 0, cast(LPARAM, @tvi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Expands the list of child items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB TreeView_ExpandChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Expand the node
      TreeView_Expand(hwndTV, hNode, TVM_EXPAND)
      ' // Get the first child node
      hChildNode = TreeView_GetChild(hwndTV, hNode)
      IF hChildNode THEN TreeView_ExpandChildItems(hwndTV, hChildNode)
      hNode = NULL
   LOOP
END SUB
' ========================================================================================

' ========================================================================================
' Expands the list of sibling items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB TreeView_ExpandSiblings (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Expand the node
      TreeView_Expand(hwndTV, hNode, TVM_EXPAND)
      ' // Get the next sibling item
      hNode = TreeView_GetNextSibling(hwndTV, hNode)
   LOOP
END SUB
' ========================================================================================

' ========================================================================================
' Expands the list of child items, if any, associated with the specified parent node and
' the sibling nodes below it.
' ========================================================================================
PRIVATE SUB TreeView_ExpandAllChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Expand the node
      TreeView_Expand(hwndTV, hNode, TVM_EXPAND)
      ' // Get the first child node
      hChildNode = TreeView_GetChild(hwndTV, hNode)
      IF hChildNode THEN TreeView_ExpandAllChildItems(hwndTV, hChildNode)
      ' // Get the next sibling item
      hNode = TreeView_GetNextSibling(hwndTV, hNode)
   LOOP
END SUB
' ========================================================================================

' ========================================================================================
' Expands all the list of items.
' Note: This is a very slow process if the control has many items because of the continuous
' redrawing; therefore, disable redrawing before calling it with
'   SendMessageW(hwndTV, WM_SETREDRAW, FALSE, 0)
' and reenable drawing once it has ended with
'   SendMessageW(hwndTV, WM_SETREDRAW, TRUE, 0)
' followed by a redrawing of the control
'   InvalidateRect hwndTV, NULL, CTRUE
'   UpdateWindow hwndTV
' ========================================================================================
PRIVATE SUB TreeView_ExpandAllItems (BYVAL hwndTV AS HWND)
   DIM hNode AS HTREEITEM = TreeView_GetRoot(hwndTV)
   IF hNode THEN TreeView_ExpandAllChildItems(hwndTV, hNode)
END SUB
' ========================================================================================

' ========================================================================================
' Collapses the list of child items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB TreeView_CollapseChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Collapse the node
      TreeView_Expand(hwndTV, hNode, TVE_COLLAPSE)
      ' // Get the first child node
      hChildNode = TreeView_GetChild(hwndTV, hNode)
      IF hChildNode THEN TreeView_CollapseChildItems(hwndTV, hChildNode)
      hNode = NULL
   LOOP
END SUB
' ========================================================================================

' ========================================================================================
' Collapses the list of sibling items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB TreeView_CollapseSiblings (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Collapse the node
      TreeView_Expand(hwndTV, hNode, TVE_COLLAPSE)
      ' // Get the next sibling item
      hNode = TreeView_GetNextSibling(hwndTV, hNode)
   LOOP
END SUB
' ========================================================================================

' ========================================================================================
' Collapses the list of child items, if any, associated with the specified parent node and
' the sibling nodes below it.
' ========================================================================================
PRIVATE SUB TreeView_CollapseAllChildItems (BYVAL hwndTV AS HWND, BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Collapse the node
      TreeView_Expand(hwndTV, hNode, TVE_COLLAPSE)
      ' // Get the first child node
      hChildNode = TreeView_GetChild(hwndTV, hNode)
      IF hChildNode THEN TreeView_CollapseAllChildItems(hwndTV, hChildNode)
      ' // Get the next sibling item
      hNode = TreeView_GetNextSibling(hwndTV, hNode)
   LOOP
END SUB
' ========================================================================================

' ========================================================================================
' Collapses all the list of items.
' ========================================================================================
PRIVATE SUB TreeView_CollapseAllItems (BYVAL hwndTV AS HWND)
   DIM hNode AS HTREEITEM = TreeView_GetRoot(hwndTV)
   IF hNode THEN TreeView_CollapseAllChildItems(hwndTV, hNode)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the text of the specified item.
' Returns TRUE if successful, or FALSE otherwise.
' Usage example:
' DIM wszText AS WSTRING * 260
' TreeView_GetItemText(hTreeView, hItem, @wszText, 260)
' AfxMsg(wszText)
' ========================================================================================
PRIVATE FUNCTION TreeView_GetItemText (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchtextmax AS LONG) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM tvi AS TVITEMW
   tvi.hitem = hItem
   tvi.mask = TVIF_TEXT
   tvi.psztext = pwszText
   tvi.cchtextmax = cchtextmax
   FUNCTION = SendMessageW(hwndTV, TVM_GETITEMW, 0, cast(LPARAM, cast(TVITEMW PTR, @tvi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text of the parent item of the specified tree-view item.
' Usage example:
' DIM wszText AS WSTRING * 260
' TreeView_GetParentText(hTreeView, hItem, @wszText, 260)
' AfxMsg(wszText)
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION TreeView_GetParentText (BYVAL hwndTV AS HWND, BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM hNode AS HTREEITEM = TreeView_GetNextItem(hwndTV, hItem, TVGN_PARENT)
   IF hNode THEN FUNCTION = TreeView_GetItemText(hwndTV, hNode, pwszText, cchTextMax)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for an exact string in a tree view node. Search is case sensitive.
' Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION TreeView_ItemExists (BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM cbLen AS LONG = LEN(*pwszText)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM dwsItemText AS DWSTRING = WSPACE(cbLen + 1)
   DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, hNode)
   DO
      IF hItem = NULL THEN EXIT FUNCTION
      IF TreeView_GetItemText(hTreeView, hItem, *dwsItemText, LEN(dwsItemText)) = FALSE THEN EXIT DO
      IF dwsItemText = *pwszText THEN
         FUNCTION = TRUE
         EXIT DO
      END IF
      hItem = TreeView_GetNextSibling(hTreeView, hItem)
   LOOP
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for an exact string in a tree view node. Search is case insensitive.
' Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION TreeView_ItemExistsIgnoreCase (BYVAL hTreeView AS HWND, BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM cbLen AS LONG = LEN(*pwszText)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM dwsItemText AS DWSTRING = WSPACE(cbLen + 1)
   DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, hNode)
   DO
      IF hItem = NULL THEN EXIT FUNCTION
      IF TreeView_GetItemText(hTreeView, hItem, *dwsItemText, LEN(dwsItemText)) = FALSE THEN EXIT DO
      IF UCASE(dwsItemText) = UCASE(*pwszText) THEN
         FUNCTION = TRUE
         EXIT DO
      END IF
      hItem = TreeView_GetNextSibling(hTreeView, hItem)
   LOOP
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                     *** UPDOWN ***
' ########################################################################################

#define UpDown_GetAccel (hCtl, cAccels, paAccels) SendMessageW(hCtl, UDM_GETACCEL, cAccels, CAST(LPARAM, paAccels))
#define UpDown_GetBase (hCtl) SendMessageW(hCtl, UDM_GETBASE, 0, 0)
#define UpDown_GetBuddy (hCtl) SendMessageW(hCtl, UDM_GETBUDDY, 0, 0)
#define UpDown_GetPos (hCtl) SendMessageW(hCtl, UDM_GETPOS, 0, 0)
#define UpDown_GetPos32 (hCtl, pfError) SendMessageW(hCtl, UDM_GETPOS32, 0, cast(LPARAM, pfError))
#define UpDown_GetRange (hCtl) SendMessageW(hCtl, UDM_GETRANGE, 0, 0)
#define UpDown_GetRange32 (hCtl, pLow, pHigh) SendMessageW(hCtl, UDM_GETRANGE32, pLow, pHigh)
#define UpDown_GetUnicodeFormat (hCtl) SendMessageW(hCtl, UDM_GETUNICODEFORMAT, 0, 0)
#define UpDown_SetAccel (hCtl, cAccels, paAccels) SendMessageW(hCtl, UDM_SETACCEL, cAccels, cast(LPARAM, paAccels))
#define UpDown_SetBase (hCtl, nBase) SendMessageW(hCtl, UDM_SETBASE, nBase, 0)
#define UpDown_SetBuddy (hCtl, hwndBuddy) SendMessageW(hCtl, UDM_SETBUDDY, cast(WPARAM, hwndBuddy), 0))
#define UpDown_SetPos (hCtl, hwndBuddy) SendMessageW(hCtl, UDM_SETPOS, 0, cast(LPARAM, MAKELONG(nPos, 0)))
#define UpDown_SetPos32 (hCtl, nPos) SendMessageW(hCtl, UDM_SETPOS32, 0, nPos)
#define UpDown_SetRange (hCtl, nPos) SendMessageW(hCtl, UDM_SETRANGE, 0, MAKELONG(nUpper, nLower))
#define UpDown_SetRange32 (hCtl, iLow, iHigh) SendMessageW(hCtl, UDM_SETRANGE32, iLow, iHigh)
#define UpDown_SetUnicodeFormat (hCtl, fUnicode) SendMessageW(hCtl, UDM_SETUNICODEFORMAT, fUnicode, 0)
