' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxCtl2.inc
' Contents: Common controls macros and functions
' Compiler: FreeBASIC 32 & 64 bit
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

' ########################################################################################
'                                     *** BUTTON ***
' ########################################################################################

' ========================================================================================
PRIVATE SUB Button_Click (BYVAL hButton AS HWND)
   SendMessageW(hButton, BM_CLICK, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_GetImage (BYVAL hButton AS HWND) AS HANDLE
   FUNCTION = cast(HANDLE, SendMessageW(hButton, BM_GETIMAGE, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_SetBitmap (BYVAL hButton AS HWND, BYVAL hbmp AS HBITMAP) AS HBITMAP
   FUNCTION = cast(HBITMAP, SendMessageW(hButton, BM_SETIMAGE, cast(WPARAM, IMAGE_BITMAP), cast(LPARAM, hbmp)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_SetIcon (BYVAL hButton AS HWND, BYVAL hIcon AS HICON) AS HICON
   FUNCTION = cast(HICON, SendMessageW(hButton, BM_SETIMAGE, cast(WPARAM, IMAGE_ICON), cast(LPARAM, hIcon)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_SetImage (BYVAL hButton AS HWND, BYVAL ImageType AS DWORD, BYVAL hImage AS HANDLE) AS HANDLE
   FUNCTION = cast(HANDLE, SendMessageW(hButton, BM_SETIMAGE, cast(WPARAM, ImageType), cast(LPARAM, hImage)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_DeleteBitmap (BYVAL hButton AS HWND) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hButton, BM_SETIMAGE, cast(WPARAM, IMAGE_BITMAP), cast(LPARAM, NULL))))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_DeleteIcon (BYVAL hButton AS HWND) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hButton, BM_SETIMAGE, cast(WPARAM, IMAGE_ICON), cast(LPARAM, NULL))))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_DeleteImage (BYVAL hButton AS HWND, BYVAL ImageType AS DWORD) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hButton, BM_SETIMAGE, cast(WPARAM, ImageType), cast(LPARAM, NULL))))
END FUNCTION
' ========================================================================================
' ========================================================================================
#if _WIN32_WINNT = &h0602
PRIVATE SUB Button_SetDontClick (BYVAL hButton AS HWND, BYVAL bState AS BOOL)
   SendMessageW(hButton, BM_SETDONTCLICK, cast(WPARAM, bState), 0)
END SUB
#endif
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Button_SetImageListXY (BYVAL hButton AS HWND, BYVAL himl AS HIMAGELIST, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG, BYVAL uALign AS DWORD = 0) AS BOOLEAN
   DIM bi AS BUTTON_IMAGELIST = (himl, (nLeft, nTop, nRight, nBottom), uAlign)
   FUNCTION = SendMessageW(hButton, BCM_SETIMAGELIST, 0, cast(LPARAM, @bi))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** COMBOBOX ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION ComboBox_GetComboBoxInfo (BYVAL hComboBox AS HWND, BYVAL pcbi AS COMBOBOXINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hComboBox, CB_GETCOMBOBOXINFO, 0, cast(LPARAM, pcbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_GetDroppedWidth (BYVAL hComboBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_GETDROPPEDWIDTH, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_GetHorizontalExtent (BYVAL hComboBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_GETHORIZONTALEXTENT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_GetLocale (BYVAL hComboBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_GETLOCALE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_GetTopIndex (BYVAL hComboBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_GETTOPINDEX, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_InitStorage (BYVAL hComboBox AS HWND, BYVAL items AS LONG, BYVAL cb AS LONG) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_INITSTORAGE, CAST(WPARAM, items), CAST(LPARAM, cb))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_SetDroppedWidth (BYVAL hComboBox AS HWND, BYVAL nWidth AS LONG) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_SETDROPPEDWIDTH, cast(WPARAM, nWidth), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB ComboBox_SetHorizontalExtent (BYVAL hComboBox AS HWND, BYVAL nWidth AS LONG)
   SendMessageW(hComboBox, CB_SETHORIZONTALEXTENT, cast(WPARAM, nWidth), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_SetLocale (BYVAL hComboBox AS HWND, BYVAL lcid AS LONG) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_SETLOCALE, cast(WPARAM, lcid), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBox_SetTopIndex (BYVAL hComboBox AS HWND, BYVAL index AS DWORD) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_SETTOPINDEX, cast(WPARAM, index), 0)
END FUNCTION
' ========================================================================================

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

' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_DeleteItem (BYVAL hComboboxEx AS HWND, BYVAL nIndex AS DWORD) AS INT_
   FUNCTION = SendMessageW(hComboboxEx, CBEM_DELETEITEM, cast(WPARAM, nIndex), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_GetComboControl (BYVAL hComboboxEx AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hComboBoxEx, CBEM_GETCOMBOCONTROL, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_GetEditControl (BYVAL hComboboxEx AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hComboBoxEx, CBEM_GETEDITCONTROL, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_GetExtendedStyle (BYVAL hComboboxEx AS HWND) AS DWORD
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_GetImageList (BYVAL hComboboxEx AS HWND) AS HIMAGELIST
   FUNCTION = cast(HIMAGELIST, SendMessageW(hComboBoxEx, CBEM_GETIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_GetItem (BYVAL hComboboxEx AS HWND, BYVAL pCBItem AS COMBOBOXEXITEMW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_GETITEMW, 0, cast(LPARAM, pCBItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_GetUnicodeFormat (BYVAL hComboboxEx AS HWND) AS DWORD
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_HasEditChanged (BYVAL hComboboxEx AS HWND) AS BOOLEAN
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_HASEDITCHANGED, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_InsertItem (BYVAL hComboboxEx AS HWND, BYVAL lpCBItem AS COMBOBOXEXITEMW PTR) AS LONG
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_INSERTITEMW, 0, cast(LPARAM, lpCBItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_SetExtendedStyle (BYVAL hComboboxEx AS HWND, BYVAL dwExMask AS DWORD, BYVAL dwExStyle AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_SETEXTENDEDSTYLE, cast(WPARAM, dwExMask), cast(LPARAM, dwExStyle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_SetImageList (BYVAL hComboboxEx AS HWND, BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = cast(HIMAGELIST, SendMessageW(hComboBoxEx, CBEM_SETIMAGELIST, 0, cast(LPARAM, himl)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_SetItem (BYVAL hComboboxEx AS HWND, BYVAL lpCBItem AS COMBOBOXEXITEMW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_SETITEMW, 0, cast(LPARAM, lpCBItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ComboBoxEx_SetUnicodeFormat (BYVAL hComboboxEx AS HWND, BYVAL fUnicode AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hComboBoxEx, CBEM_SETUNICODEFORMAT, cast(WPARAM, fUnicode), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB ComboBoxEx_SetWindowTheme (BYVAL hComboboxEx AS HWND, BYVAL pwszStyle AS WSTRING PTR)
   SendMessageW(hComboBoxEx, CBEM_SETWINDOWTHEME, 0, cast(LPARAM, pwszStyle))
END SUB
' ========================================================================================


' ########################################################################################
'                                      *** EDIT ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION Edit_CharFromPos (BYVAL hEdit AS HWND, BYVAL x AS SHORT, BYVAL y AS SHORT) AS LONG
   FUNCTION = SendMessageW(hEdit, EM_CHARFROMPOS, 0, cast(LPARAM, MAKELONG(x, y)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_Clear (BYVAL hEdit AS HWND)
   SendMessageW(hEdit, WM_CLEAR, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_Copy (BYVAL hEdit AS HWND)
   SendMessageW(hEdit, WM_COPY, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_Cut (BYVAL hEdit AS HWND)
   SendMessageW(hEdit, WM_CUT, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetIMEStatus (BYVAL hEdit AS HWND, BYVAL nStatusType AS LONG) AS LONG
   FUNCTION = SendMessageW(hEdit, EM_GETIMESTATUS, cast(WPARAM, nStatusType), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetLimitText (BYVAL hEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hEdit, EM_GETLIMITTEXT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetMargins (BYVAL hEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hEdit, EM_GETMARGINS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetLeftMargin (BYVAL hEdit AS HWND) AS WORD
   FUNCTION = LOWORD(SendMessageW(hEdit, EM_GETMARGINS, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the width of the right margin for an edit control.
' ========================================================================================
PRIVATE FUNCTION Edit_GetRightMargin (BYVAL hEdit AS HWND) AS WORD
   FUNCTION = HIWORD(SendMessageW(hEdit, EM_GETMARGINS, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetSelStart (BYVAL hEdit AS HWND) AS WORD
   FUNCTION = LOWORD(SendMessageW(hEdit, EM_GETSEL, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetSelEnd (BYVAL hEdit AS HWND) AS WORD
   FUNCTION = HIWORD(SendMessageW(hEdit, EM_GETSEL, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_GetThumb (BYVAL hEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hEdit, EM_GETTHUMB, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_LineScroll (BYVAL hEdit AS HWND, BYVAL x AS LONG, BYVAL y AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hEdit, EM_LINESCROLL, cast(WPARAM, x), cast(LPARAM, y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_Paste (BYVAL hEdit AS HWND)
   SendMessageW(hEdit, WM_PASTE, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_PosFromChar (BYVAL hEdit AS HWND, BYVAL index AS DWORD) AS LONG
   FUNCTION = SendMessageW(hEdit, EM_POSFROMCHAR, cast(WPARAM, index), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_SetCueBannerTextFocused (BYVAL hEdit AS HWND, BYVAL pwszText AS WSTRING PTR, BYVAL fDrawFocused AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hEdit, EM_SETCUEBANNER, cast(WPARAM, fDrawFocused), cast(WPARAM, pwszText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Edit_SetIMEStatus (BYVAL hEdit AS HWND, BYVAL nStatusType AS LONG, BYVAL flags AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hEdit, EM_SETIMESTATUS, cast(WPARAM, nStatusType), cast(LPARAM, flags))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_SetMargins (BYVAL hEdit AS HWND, BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
   SendMessageW(hEdit, EM_SETMARGINS, cast(WPARAM, nMargins), cast(LPARAM, nWidth))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_SetLeftMargin (BYVAL hEdit AS HWND, BYVAL nWidth AS WORD)
   SendMessageW(hEdit, EM_SETMARGINS, cast(WPARAM, EC_LEFTMARGIN), cast(LPARAM, MAKELONG(nWidth, 0)))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Edit_SetRightMargin (BYVAL hEdit AS HWND, BYVAL nWidth AS WORD)
   SendMessageW(hEdit, EM_SETMARGINS, cast(WPARAM, EC_RIGHTMARGIN), cast(LPARAM, MAKELONG(0, nWidth)))
END SUB
' ========================================================================================


' ########################################################################################
'                                      *** HEADER ***
' ########################################################################################

' ========================================================================================
' Tests a point to determine which header item, if any, is at the specified point.
' ========================================================================================
PRIVATE FUNCTION Header_HitTest (BYVAL hCtl AS HWND, BYVAL phdhti AS HDHITTESTINFO PTR) AS LONG
   RETURN SendMessageW(hCtl, HDM_HITTEST, 0, cast(LPARAM, phdhti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the specified item of a header control.
' Returns the index of the item at the specified position, if any, or – 1 otherwise.
' ========================================================================================
PRIVATE FUNCTION Header_GetItemWidth (BYVAL hCtl AS HWND, BYVAL nItem AS LONG) AS LONG
   DIM hdi AS HDITEMW : hdi.mask = HDI_WIDTH
   IF SendMessageW(hCtl, HDM_GETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi))) THEN RETURN hdi.cxy
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width of the specified item of a header control.
' Returns nonzero upon success, or zero otherwise.
' ========================================================================================
PRIVATE FUNCTION Header_SetItemWidth (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL nWidth AS LONG) AS BOOLEAN
   DIM hdi AS HDITEMW : hdi.mask = HDI_WIDTH : hdi.cxy = nWidth
   RETURN SendMessageW(hCtl, HDM_SETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the order of the specified item of a header control.
' ========================================================================================
PRIVATE FUNCTION Header_GetItemOrder (BYVAL hCtl AS HWND, BYVAL nItem AS LONG) AS LONG
   DIM hdi AS HDITEMW : hdi.mask = HDI_ORDER
   IF SendMessageW(hCtl, HDM_GETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi))) THEN RETURN hdi.iOrder
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the order of the specified item of a header control.
' ========================================================================================
PRIVATE FUNCTION Header_SetItemOrder (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL iOrder AS LONG) AS LONG
   DIM hdi AS HDITEMW : hdi.mask = HDI_ORDER : hdi.iOrder = iOrder
   RETURN SendMessageW(hCtl, HDM_SETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
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
   RETURN SendMessageW(hCtl, HDM_GETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified item. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION Header_SetItemText (BYVAL hCtl AS HWND, BYVAL nItem AS LONG, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM hdi AS HDITEMW : hdi.mask = HDI_TEXT : hdi.cchTextMax = LEN(*pwszText) : hdi.pszText = pwszText
   RETURN SendMessageW(hCtl, HDM_SETITEMW, cast(WPARAM, nItem), cast(LPARAM, cast(HDITEMW PTR, @hdi)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** HOTKEY ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION HotKey_GetHotKey (BYVAL hCtl AS HWND) AS DWORD
   RETURN SendMessageW(hCtl, HKM_GETHOTKEY, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB HotKey_SetHotKey (BYVAL hCtl AS HWND, BYVAL vkCode AS SHORT, BYVAL fMods AS SHORT)
   SendMessageW(hCtl, HKM_SETHOTKEY, MAKELONG(vkCode, fMods), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB HotKey_SetRules (BYVAL hCtl AS HWND, BYVAL fwCombInv AS WORD, BYVAL fwModInv AS WORD)
   SendMessageW(hCtl, HKM_SETRULES, cast(WPARAM, fwCombInv), cast(LPARAM, MAKELONG(fwModInv, 0)))
END SUB
' ========================================================================================

' ########################################################################################
'                                   *** IP ADDRESS ***
' ########################################################################################

' ========================================================================================
PRIVATE SUB IPAddress_ClearAddress (BYVAL hIPM AS HWND)
   SendMessageW(hIPM, IPM_CLEARADDRESS, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION IPAddress_GetAddress (BYVAL hIPM AS HWND, BYVAL pdwAddr AS DWORD PTR) AS LONG
   FUNCTION = SendMessageW(hIPM, IPM_GETADDRESS, 0, CAST(LPARAM, pdwAddr))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB IPAddress_IsBlank (BYVAL hIPM AS HWND)
   SendMessageW(hIPM, IPM_ISBLANK, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB IPAddress_SetAddress (BYVAL hIPM AS HWND, BYVAL dwAddr AS DWORD)
   SendMessageW(hIPM, IPM_SETADDRESS, 0, cast(LPARAM, dwAddr))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB IPAddress_SetFocus (BYVAL hIPM AS HWND, BYVAL nField AS LONG)
   SendMessageW(hIPM, IPM_SETFOCUS, cast(WPARAM, nField), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION IPAddress_SetRange (BYVAL hIPM AS HWND, BYVAL nField AS LONG, BYVAL loRange AS BYTE, BYVAL hiRange AS BYTE) AS BOOLEAN
   FUNCTION = SendMessageW(hIPM, IPM_SETRANGE, cast(WPARAM, nField), cast(LPARAM, MAKEIPRANGE(loRange, hiRange)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** LISTBOX ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION ListBox_AddFile (BYVAL hListBox AS HWND, BYVAL pwszFileName AS WSTRING PTR) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_ADDFILE, 0, cast(LPARAM, pwszFileName))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBox_GetAnchorIndex (BYVAL hListBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_GETANCHORINDEX, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Listbox_GetListboxInfo (BYVAL hListBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_GETLISTBOXINFO, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBox_GetLocale (BYVAL hListBox AS HWND) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_GETLOCALE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Listbox_InitStorage (BYVAL hListBox AS HWND, BYVAL nItems AS DWORD, BYVAL nBytes AS DWORD) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_INITSTORAGE, cast(WPARAM, nItems), cast(LPARAM, nBytes))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Listbox_ItemFromPoint (BYVAL hListBox AS HWND, BYVAL x AS SHORT, BYVAL y AS SHORT) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_ITEMFROMPOINT, 0, cast(LPARAM, MAKELONG(x, y)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Listbox_SelItemRangeEx (BYVAL hListBox AS HWND, BYVAL startIndex AS LONG, BYVAL endIndex AS LONG) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_SELITEMRANGEEX, cast(WPARAM, startIndex), cast(LPARAM, endIndex))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBox_SetAnchorIndex (BYVAL hListBox AS HWND, BYVAL index AS LONG) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_SETANCHORINDEX, cast(WPARAM, index), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Listbox_SetCount (BYVAL hListBox AS HWND, BYVAL nCount AS DWORD) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_SETCOUNT, cast(WPARAM, nCount), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListBox_SetLocale (BYVAL hListBox AS HWND, BYVAL wLocaleID AS DWORD) AS LONG
   FUNCTION = SendMessageW(hListBox, LB_SETLOCALE, cast(WPARAM, wLocaleID), 0)
END FUNCTION
' ========================================================================================

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

' ========================================================================================
PRIVATE FUNCTION ListView_GetFocusedItem (BYVAL hCtl AS HWND) AS LONG
   FUNCTION = SendMessageW(hCtl, LVM_GETNEXTITEM, -1, LVNI_ALL OR LVNI_FOCUSED)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListView_GetISearchStringLen (BYVAL hCtl AS HWND) AS LONG
   FUNCTION = SendMessageW(hCtl, LVM_GETISEARCHSTRINGW, 0, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListView_GetSelection (BYVAL hCtl AS HWND) AS LONG
   FUNCTION = SendMessageW(hCtl, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListView_HitTestEx (BYVAL hCtl AS HWND, BYVAL pinfo AS LVHITTESTINFO PTR) AS LONG
   FUNCTION = SendMessageW(hCtl, LVM_HITTEST, cast(WPARAM, -1), cast(LPARAM, pinfo))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB ListView_SelectItem (BYVAL hCtl AS HWND, BYVAL nIndex AS LONG)
   ListView_SetItemState(hCtl, nIndex, LVIS_FOCUSED OR LVIS_SELECTED, &H000F)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB ListView_UnselectItem (BYVAL hCtl AS HWND, BYVAL nIndex AS LONG)
   ListView_SetItemState(hCtl, nIndex, 0, LVIS_SELECTED)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB ListView_UnselectAllItems (BYVAL hCtl AS HWND)
   ListView_SetItemState(hCtl, -1, 0, LVIS_SELECTED OR LVIS_FOCUSED)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListView_FitContent (BYVAL hCtl AS HWND, BYVAL nCol AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, cast(WPARAM, nCol), cast(LPARAM, LVSCW_AUTOSIZE))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ListView_FitHeader (BYVAL hCtl AS HWND, BYVAL nCol AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, cast(WPARAM, nCol), cast(LPARAM, LVSCW_AUTOSIZE_USEHEADER))
END FUNCTION
' ========================================================================================

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

' ========================================================================================
PRIVATE SUB Rebar_BeginDrag (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL dwPos AS DWORD)
   SendMessageW(hRebar, RB_BEGINDRAG, cast(WPARAM, uBand), cast(LPARAM, dwPos))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_DeleteBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_DELETEBAND, cast(WPARAM, uBand), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_DragMove (BYVAL hRebar AS HWND, BYVAL dwPos AS DWORD)
   SendMessageW(hRebar, RB_DRAGMOVE, 0, cast(LPARAM, dwPos))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_EndDrag (BYVAL hRebar AS HWND)
   SendMessageW(hRebar, RB_ENDDRAG, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_GetBandBorders (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL lprc AS RECT PTR)
   SendMessageW(hRebar, RB_GETBANDBORDERS, cast(WPARAM, uBand), cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetBandCount (BYVAL hRebar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_GETBANDCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_GetBandMargins (BYVAL hRebar AS HWND, BYVAL pMargins AS MARGINS PTR)
   SendMessageW(hRebar, RB_GETBANDMARGINS, 0, cast(LPARAM, pMargins))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetBarHeight (BYVAL hRebar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_GETBARHEIGHT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetBarInfo (BYVAL hRebar AS HWND, BYVAL lprbi AS REBARINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_GETBARINFO, 0, cast(LPARAM, lprbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetBkColor (BYVAL hRebar AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hRebar, RB_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetColorScheme (BYVAL hRebar AS HWND, BYVAL lpcs AS COLORSCHEME PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_GETCOLORSCHEME, 0, cast(LPARAM, lpcs))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetExtendedStyle (BYVAL hRebar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetPalette (BYVAL hRebar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_GETPALETTE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetRect (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_GETRECT, cast(WPARAM, uBand), cast(LPARAM, lprc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetRowCount (BYVAL hRebar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_GETROWCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetRowHeight (BYVAL hRebar AS HWND, BYVAL uRow AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_GETROWHEIGHT, cast(WPARAM, uRow), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetTextColor (BYVAL hRebar AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hRebar, RB_GETTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetTooltips (BYVAL hRebar AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hRebar, RB_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_GetUnicodeFormat (BYVAL hRebar AS HWND) AS LONG
   FUNCTION = SendMessageW(hRebar, RB_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_HitTest (BYVAL hRebar AS HWND, BYVAL lprhbt AS RBHITTESTINFO PTR) AS LONG
   FUNCTION = SendMessageW(hRebar, RB_HITTEST, 0, cast(LPARAM, lprhbt))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_IdToIndex (BYVAL hRebar AS HWND, BYVAL uBandID AS DWORD) AS LONG
   FUNCTION = SendMessageW(hRebar, RB_IDTOINDEX, cast(WPARAM, uBandID), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_InsertBand (BYVAL hRebar AS HWND, BYVAL nIndex AS LONG, BYVAL lprbbi AS REBARBANDINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_INSERTBAND, cast(WPARAM, nIndex), cast(LPARAM, lprbbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_MaximizeBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL fIdeal AS DWORD)
   SendMessageW(hRebar, RB_MAXIMIZEBAND, cast(WPARAM, uBand), cast(LPARAM, fIdeal))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_MinimizeBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD)
  SendMessageW(hRebar, RB_MINIMIZEBAND, cast(WPARAM, uBand), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_MoveBand (BYVAL hRebar AS HWND, BYVAL nFrom AS DWORD, BYVAL nTo AS DWORD) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_MOVEBAND, cast(WPARAM, nFrom), cast(LPARAM, nTo))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_PushChevron (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL iAppValue AS DWORD)
   SendMessageW(hRebar, RB_PUSHCHEVRON, cast(WPARAM, uBand), cast(LPARAM, iAppValue))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL lprbbi AS REBARBANDINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_SETBANDINFO, cast(WPARAM, uBand), cast(LPARAM, lprbbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetBarInfo (BYVAL hRebar AS HWND, BYVAL lprbi AS REBARINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_SETBARINFO, 0, cast(LPARAM, lprbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetBkColor (BYVAL hRebar AS HWND, BYVAL clrBk AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_SETBKCOLOR, 0, cast(LPARAM, clrBk))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_SetColorScheme (BYVAL hRebar AS HWND, BYVAL lpcs AS COLORSCHEME PTR)
   SendMessageW(hRebar, RB_SETCOLORSCHEME, 0, cast(LPARAM, lpcs))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetPalette (BYVAL hRebar AS HWND, BYVAL hpal AS HPALETTE) AS HPALETTE
   FUNCTION = cast(HPALETTE, SendMessageW(hRebar, RB_SETPALETTE, 0, cast(LPARAM, hpal)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetParent (BYVAL hRebar AS HWND, BYVAL hwndParent AS hWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hRebar, RB_SETPARENT, cast(WPARAM, hwndParent), 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetTextColor (BYVAL hRebar AS HWND, BYVAL clrText AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(hRebar, RB_SETTEXTCOLOR, 0, cast(LPARAM, clrText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_SetTooltips (BYVAL hRebar AS HWND, BYVAL hwndToolTip AS HWND)
   SendMessageW(hRebar, RB_SETTOOLTIPS, cast(WPARAM, hwndToolTip), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SetUnicodeFormat (BYVAL hRebar AS HWND, BYVAL fUnicode AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRebar, RB_SETUNICODEFORMAT, cast(WPARAM, fUnicode), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Rebar_SetWindowTheme (BYVAL hRebar AS HWND, BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(hRebar, RB_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_ShowBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL fShow AS DWORD) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_SHOWBAND, cast(WPARAM, uBand), cast(LPARAM, fShow))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Rebar_SizeToRect (BYVAL hRebar AS HWND, BYVAL prc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_SIZETORECT, 0, cast(LPARAM, prc))
END FUNCTION
' ========================================================================================
' ========================================================================================
#if _WIN32_WINNT = &h0602
PRIVATE FUNCTION Rebar_SetBandWidth (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL nWidth AS DWORD) AS BOOLEAN
   FUNCTION = SendMessageW(hRebar, RB_SETBANDWIDTH, cast(WPARAM, uBand), cast(LPARAM, nWidth))
END FUNCTION
#endif
' ========================================================================================
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

' ========================================================================================
PRIVATE FUNCTION Static_DeleteBitmap (BYVAL hStatic AS HWND) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, IMAGE_BITMAP), 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_DeleteCursor (BYVAL hStatic AS HWND) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, IMAGE_CURSOR), 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_DeleteEnhancedMetafile (BYVAL hStatic AS HWND) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, IMAGE_ENHMETAFILE), 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_DeleteIcon (BYVAL hStatic AS HWND) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETICON, 0, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_DeleteImage (BYVAL hStatic AS HWND, BYVAL nType AS LONG) AS BOOLEAN
   FUNCTION = DeleteObject(cast(HGDIOBJ, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, nType), 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_GetImage (BYVAL hStatic AS HWND, BYVAL nType AS LONG) AS HANDLE
   FUNCTION = cast(HANDLE, SendMessageW(hStatic, STM_GETIMAGE, nType, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_SetBitmap (BYVAL hStatic AS HWND, BYVAL hBitmap AS HBITMAP) AS HBITMAP
   FUNCTION = cast(HBITMAP, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, IMAGE_BITMAP), cast(LPARAM, hBitmap)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_SetCursor (BYVAL hStatic AS HWND, BYVAL hCursor AS HCURSOR) AS HCURSOR
   FUNCTION = cast(HCURSOR, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, IMAGE_CURSOR), cast(LPARAM, hCursor)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_SetEnhancedMetafile (BYVAL hStatic AS HWND, BYVAL hEnhMetafile AS HENHMETAFILE) AS HENHMETAFILE
   FUNCTION = cast(HENHMETAFILE, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, IMAGE_ENHMETAFILE), cast(LPARAM, hEnhMetafile)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Static_SetImage (BYVAL hStatic AS HWND, BYVAL nType AS LONG, BYVAL hImage AS HANDLE) AS HANDLE
   FUNCTION = cast(HANDLE, SendMessageW(hStatic, STM_SETIMAGE, cast(WPARAM, nType), cast(LPARAM, hImage)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                  *** PROGRESS BAR ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION ProgressBar_DeltaPos (BYVAL hProgress AS HWND, BYVAL nIncrement AS LONG) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_DELTAPOS, cast(WPARAM, nIncrement), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetHighLimit (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_GETRANGE, FALSE, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetLowLimit (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_GETRANGE, TRUE, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetMaxRange (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_GETRANGE, FALSE, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetMinRange (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_GETRANGE, TRUE, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetPos (BYVAL hProgress AS HWND) AS DWORD
   FUNCTION = SendMessageW(hProgress, PBM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetRange (BYVAL hProgress AS HWND, BYVAL fWhichLimit AS DWORD, BYVAL ppBRange AS PBRANGE PTR) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_GETRANGE, cast(WPARAM, fWhichLimit), cast(LPARAM, ppBRange))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_IncPos (BYVAL hProgress AS HWND, BYVAL nIncVal AS LONG) AS LONG
   DIM nPos AS LONG = SendMessageW(hProgress, PBM_GETPOS, 0, 0)
   FUNCTION = SendMessageW(hProgress, PBM_SETPOS, cast(WPARAM, nPos + nIncVal), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetBarColor (BYVAL hProgress AS HWND, BYVAL clrBar AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(hProgress, PBM_SETBARCOLOR, 0, cast(LPARAM, clrBar))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetBkColor (BYVAL hProgress AS HWND, BYVAL clrBk AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(hProgress, PBM_SETBKCOLOR, 0, cast(LPARAM, clrBk))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetMarquee (BYVAL hProgress AS HWND, BYVAL fOn AS LONG, BYVAL dwAnimationSpeed AS DWORD) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_SETMARQUEE, cast(WPARAM, fOn), cast(LPARAM, dwAnimationSpeed))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetMarqueeOff (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_SETMARQUEE, cast(WPARAM, FALSE), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetMarqueeOn (BYVAL hProgress AS HWND, BYVAL dwAnimationSpeed AS DWORD) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_SETMARQUEE, cast(WPARAM, CTRUE), cast(LPARAM, dwAnimationSpeed))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetPos (BYVAL hProgress AS HWND, BYVAL nNewPos AS LONG) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_SETPOS, cast(WPARAM, nNewPos), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetRange (BYVAL hProgress AS HWND, BYVAL nMinRange AS WORD, BYVAL nMaxRange AS WORD) AS DWORD
   FUNCTION = SendMessageW(hProgress, PBM_SETRANGE, 0, cast(LPARAM, MAKELONG(nMinRange, nMaxRange)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetRange32 (BYVAL hProgress AS HWND, BYVAL nMinRange AS LONG, BYVAL nMaxRange AS LONG) AS DWORD
   FUNCTION = SendMessageW(hProgress, PBM_SETRANGE32, cast(WPARAM, nMinRange), cast(LPARAM, nMaxRange))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetStep (BYVAL hProgress AS HWND, BYVAL nStepInc AS LONG) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_SETSTEP, cast(WPARAM, nStepInc), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_StepIt (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_STEPIT, 0, 0)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetBarColor (BYVAL hProgress AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hProgress, PBM_GETBARCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetBkColor (BYVAL hProgress AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hProgress, PBM_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_GetState (BYVAL hProgress AS HWND) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_GETSTATE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION ProgressBar_SetState (BYVAL hProgress AS HWND, BYVAL iState AS LONG) AS LONG
   FUNCTION = SendMessageW(hProgress, PBM_SETSTATE, cast(LPARAM, iState), 0)
END FUNCTION
' ========================================================================================
#endif


' ########################################################################################
'                                  *** STATUS BAR ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION StatusBar_GetBorders (BYVAL hStatusBar AS HWND, BYVAL rgBorders AS LONG PTR) AS LONG
   FUNCTION = SendMessageW(hStatusBar, SB_GETBORDERS, 0, cast(LPARAM, rgBorders))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the height of the status bar.
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetHeight (BYVAL hStatusBar AS HWND) AS LONG
   DIM rc AS RECT
   GetWindowRect(hStatusbar, @rc)
   FUNCTION = rc.Bottom - rc.Top
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetIcon (BYVAL hStatusBar AS HWND, BYVAL nPart AS DWORD) AS HICON
  FUNCTION = cast(HICON, SendMessageW(hStatusBar, SB_GETICON, nPart, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetParts (BYVAL hStatusBar AS HWND, BYVAL nParts AS DWORD, BYVAL rgRightCoords AS LONG PTR) AS LONG
   FUNCTION = SendMessageW(hStatusBar, SB_GETPARTS, nParts, cast(LPARAM, rgRightCoords))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetPartsCount (BYVAL hStatusBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hStatusBar, SB_GETPARTS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetRect (BYVAL hStatusBar AS HWND, BYVAL nPart AS DWORD, BYVAL lprc AS RECT PTR) AS LONG
   FUNCTION = SendMessageW(hStatusBar, SB_GETRECT, cast(WPARAM, nPart), cast(LPARAM, lprc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetTextDrawingMode (BYVAL hStatusBar AS HWND, BYVAL nPart AS DWORD) AS LONG
   DIM dwResult AS DWORD = SendMessageW(hStatusBar, SB_GETTEXTLENGTH, cast(WPARAM, nPart), 0)
   FUNCTION = HIWORD(dwResult)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetTextLength (BYVAL hStatusBar AS HWND, BYVAL nPart AS DWORD) AS LONG
   DIM dwResult AS DWORD = SendMessageW(hStatusBar, SB_GETTEXTLENGTH, cast(WPARAM, nPart), 0)
   FUNCTION = LOWORD(dwResult)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB StatusBar_GetTipText (BYVAL hStatusBar AS HWND, BYVAL nPart AS WORD, BYVAL pwszText AS WSTRING PTR, BYVAL wMaxLen AS WORD)
   IF pwszText = NULL THEN EXIT SUB
   SendMessageW(hStatusBar, SB_GETTIPTEXTW, cast(WPARAM, MAKELONG(nPart, wMaxLen)), cast(LPARAM, pwszText))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetUnicodeFormat (BYVAL hStatusBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hStatusBar, SB_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_IsSimple (BYVAL hStatusBar AS HWND) AS BOOLEAN
   FUNCTION = SendMessageW(hStatusBar, SB_ISSIMPLE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_SetBkColor (BYVAL hStatusBar AS HWND, BYVAL clrBk AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(hStatusBar, SB_SETBKCOLOR, 0, cast(LPARAM, clrBk))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB StatusBar_SetFont (BYVAL hStatusBar AS HWND, BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = FALSE)
   SendMessageW(hStatusBar, WM_SETFONT, cast(WPARAM, hFont), cast(LPARAM, fRedraw))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_SetIcon (BYVAL hStatusBar AS HWND, BYVAL nPart AS LONG, BYVAL hIcon AS HICON) AS BOOLEAN
   FUNCTION = SendMessageW(hStatusBar, SB_SETICON, cast(WPARAM, nPart), cast(LPARAM, hIcon))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB StatusBar_SetMinHeight (BYVAL hStatusBar AS HWND, BYVAL minHeight AS DWORD)
   SendMessageW(hStatusBar, SB_SETMINHEIGHT, cast(WPARAM, minHeight), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_SetParts (BYVAL hStatusBar AS HWND, BYVAL dwParts AS DWORD, BYVAL rgWidths AS LONG PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hStatusBar, SB_SETPARTS, cast(WPARAM, dwParts), cast(LPARAM, rgWidths))
END FUNCTION
' ========================================================================================
PRIVATE FUNCTION StatusBar_SetText (BYVAL hStatusBar AS HWND, BYVAL nPart AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL uType AS DWORD = 0) AS BOOLEAN
   FUNCTION = SendMessageW(hStatusBar, SB_SETTEXTW, cast(WPARAM, nPart OR uType), cast(LPARAM, pwszText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB StatusBar_SetTipText (BYVAL hStatusBar AS HWND, BYVAL nPart AS LONG, BYVAL pwszTipText AS WSTRING PTR)
   SendMessageW(hStatusBar, SB_SETTIPTEXTW, cast(WPARAM, nPart), cast(LPARAM, pwszTipText))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_SetUnicodeFormat (BYVAL hStatusBar AS HWND, BYVAL fUnicode AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hStatusBar, SB_SETUNICODEFORMAT, cast(WPARAM, fUnicode), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB StatusBar_Simple (BYVAL hStatusBar AS HWND, BYVAL fSimple AS BOOLEAN)
   SendMessageW(hStatusBar, SB_SIMPLE, fSimple, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION StatusBar_GetText (BYVAL hStatusBar AS HWND, BYVAL nPart AS DWORD, BYVAL pwszText AS WSTRING PTR) AS DWORD
   IF pwszText = NULL THEN EXIT FUNCTION
   FUNCTION = SendMessageW(hStatusBar, SB_GETTEXTW, cast(WPARAM, nPart), cast(LPARAM, pwszText))
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
' ========================================================================================
PRIVATE SUB TaskDialog_ClickButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS DWORD)
   SendMessageW(hTaskDlg, TDM_CLICK_BUTTON, cast(WPARAM, nButton), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_ClickRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS DWORD)
   SendMessageW(hTaskDlg, TDM_CLICK_RADIO_BUTTON, cast(WPARAM, nRadioButton), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_ClickVerification (BYVAL hTaskDlg AS HWND, BYVAL fChecked AS LONG, BYVAL fFocused AS LONG)
   SendMessageW(hTaskDlg, TDM_CLICK_VERIFICATION, cast(WPARAM, fChecked), cast(LPARAM, fFocused))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_EnableButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG, BYVAL fEnable AS LONG)
   SendMessageW(hTaskDlg, TDM_ENABLE_BUTTON, cast(WPARAM, nButton), cast(LPARAM, fEnable))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_EnableRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS LONG, BYVAL fEnable AS LONG)
   SendMessageW(hTaskDlg, TDM_ENABLE_RADIO_BUTTON, cast(WPARAM, nRadioButton), cast(LPARAM, fEnable))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_NavigatePage (BYVAL hTaskDlg AS HWND, BYVAL pTaskConfig AS TASKDIALOGCONFIG PTR)
   SendMessageW(hTaskDlg, TDM_NAVIGATE_PAGE, 0, CAST(LPARAM, pTaskConfig))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_SetButtonElevationRequiredState (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG, BYVAL fElevation AS LONG)
   SendMessageW(hTaskDlg, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, cast(WPARAM, nButton), cast(LPARAM, fElevation))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_SetElementText (BYVAL hTaskDlg AS HWND, BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   SendMessageW(hTaskDlg, TDM_SET_ELEMENT_TEXT, cast(WPARAM, eElement), cast(LPARAM, pszNewText))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_SetMarqueeProgressBar (BYVAL hTaskDlg AS HWND, BYVAL fMarquee AS BOOLEAN)
   SendMessageW(hTaskDlg, TDM_SET_MARQUEE_PROGRESS_BAR, cast(WPARAM, fMarquee), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_SetProgressBarMarquee (BYVAL hTaskDlg AS HWND, BYVAL fMarquee AS LONG, BYVAL nSpeed AS LONG)
   SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_MARQUEE, cast(WPARAM, fMarquee), cast(LPARAM, nSpeed))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION TaskDialog_SetProgressBarPos (BYVAL hTaskDlg AS HWND, BYVAL nNewPos AS LONG) AS LONG
   FUNCTION = SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_POS, cast(WPARAM, nNewPos), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION TaskDialog_SetProgressBarRange (BYVAL hTaskDlg AS HWND, BYVAL nMinRange AS SHORT, BYVAL nMaxRange AS SHORT) AS DWORD
   FUNCTION = SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_RANGE, 0, MAKELONG(nMinRange, nMaxRange))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION TaskDialog_SetProgressBarState (BYVAL hTaskDlg AS HWND, BYVAL nNewState AS LONG) AS LONG
   FUNCTION = SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_STATE, cast(WPARAM, nNewState), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_UpdateElementText (BYVAL hTaskDlg AS HWND, BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   SendMessageW(hTaskDlg, TDM_UPDATE_ELEMENT_TEXT, cast(WPARAM, eElement), CAST(LPARAM, pszNewText))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB TaskDialog_UpdateIcon (BYVAL hTaskDlg AS HWND, BYVAL nIconElement AS LONG, BYVAL idResource AS DWORD)
   SendMessageW(hTaskDlg, TDM_UPDATE_ICON, cast(WPARAM, nIconElement), CAST(LPARAM, idResource))
END SUB
' ========================================================================================
#endif


' ########################################################################################
'                                    *** TOOLBAR ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION Toolbar_AddBitmap (BYVAL hToolBar AS HWND, BYVAL nButtons AS DWORD, BYVAL tabm AS TBADDBITMAP PTR) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_ADDBITMAP, cast(WPARAM, nButtons), CAST(LPARAM, tabm))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_AddButtons (BYVAL hToolBar AS HWND, BYVAL nButtons AS DWORD, BYVAL lpButtons AS TBBUTTON PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ADDBUTTONSW, cast(WPARAM, nButtons), CAST(LPARAM, lpButtons))
END FUNCTION
' ========================================================================================
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
PRIVATE FUNCTION Toolbar_AddString (BYVAL hToolBar AS HWND, BYVAL hInst AS HINSTANCE, BYVAL idString AS INT_PTR) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_ADDSTRINGW, CAST(WPARAM, hInst), IIF(hInst = NULL, idString, MAKELONG(idString, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_AutoSize (BYVAL hToolBar AS HWND)
   SendMessageW(hToolBar, TB_AUTOSIZE, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_ButtonCount (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_BUTTONCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_ButtonStructSize (BYVAL hToolBar AS HWND, BYVAL cb AS LONG)
   SendMessageW(hToolBar, TB_BUTTONSTRUCTSIZE, cb, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_ChangeBitmap (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL iBitmap AS WORD) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_CHANGEBITMAP, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(iBitmap, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_CheckButton (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL fCheck AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_CHECKBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(fCheck, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_CommandToIndex (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_COMMANDTOINDEX, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_Customize (BYVAL hToolBar AS HWND)
   SendMessageW(hToolBar, TB_CUSTOMIZE, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_DeleteButton (BYVAL hToolBar AS HWND, BYVAL nButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_DELETEBUTTON, cast(WPARAM, nButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_EnableButton (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ENABLEBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(CTRUE, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_DisableButton (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ENABLEBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(FALSE, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetAnchorHighlight (BYVAL hToolBar AS HWND) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_GETANCHORHIGHLIGHT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetBitmap (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_GETBITMAP, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetBitmapFlags (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETBITMAPFLAGS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetButton (BYVAL hToolBar AS HWND, BYVAL nButton AS LONG, BYVAL ptbb AS TBBUTTON PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_GETBUTTON, CAST(WPARAM, nButton), CAST(LPARAM, ptbb))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetButtonSize (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETBUTTONSIZE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetButtonHeight (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = HIWORD(SendMessageW(hToolBar, TB_GETBUTTONSIZE, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetButtonWidth (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = LOWORD(SendMessageW(hToolBar, TB_GETBUTTONSIZE, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetButtonInfo (BYVAL hToolBar AS HWND, BYVAL iID AS LONG, BYVAL lptbbi AS TBBUTTONINFOW PTR) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_GETBUTTONINFOW, cast(WPARAM, iID), cast(LPARAM, lptbbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetColorScheme (BYVAL hToolBar AS HWND, BYVAL lpcs AS COLORSCHEME PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_GETCOLORSCHEME, 0, CAST(LPARAM, lpcs))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetImageList (BYVAL hToolBar AS HWND) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(hToolBar, TB_GETIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetDisabledImageList (BYVAL hToolBar AS HWND) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(hToolBar, TB_GETDISABLEDIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetHotImageList (BYVAL hToolBar AS HWND) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(hToolBar, TB_GETHOTIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetExtendedStyle (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetHorizontalPadding (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = LOWORD(SendMessageW(hToolBar, TB_GETPADDING, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetVerticalPadding (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = HIWORD(SendMessageW(hToolBar, TB_GETPADDING, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetHotItem (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETHOTITEM, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetInsertMark (BYVAL hToolBar AS HWND, BYVAL ptbim AS TBINSERTMARK PTR) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_GETINSERTMARK, 0, CAST(LPARAM, ptbim))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetInsertMarkColor (BYVAL hToolBar AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hToolBar, TB_GETINSERTMARKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetItemRect (BYVAL hToolBar AS HWND, BYVAL nButton AS LONG, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_GETITEMRECT, cast(WPARAM, nButton), CAST(LPARAM, lprc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetMaxSize (BYVAL hToolBar AS HWND, BYVAL psize AS SIZE PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_GETMAXSIZE, 0, CAST(LPARAM, psize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_GetMetrics (BYVAL hToolBar AS HWND, BYVAL ptbMetrics AS TBMETRICS PTR)
   SendMessageW(hToolBar, TB_GETMETRICS, 0, CAST(LPARAM, ptbMetrics))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetPadding (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETPADDING, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetRect (BYVAL hToolBar AS HWND, BYVAL iID AS LONG, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_GETRECT, cast(WPARAM, iID), CAST(LPARAM, lprc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetRows (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETROWS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetState (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETSTATE, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetString (BYVAL hToolBar AS HWND, BYVAL iString AS SHORT, BYVAL pBuffer AS WSTRING PTR, BYVAL cchMax AS SHORT) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_GETSTRINGW, MAKELONG(cchMax, iString), CAST(LPARAM, pbuffer))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetStyle (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETSTYLE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetTextRows (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETTEXTROWS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetTooltips (BYVAL hToolBar AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hToolBar, TB_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetUnicodeFormat (BYVAL hToolBar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_HideButton (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL fShow AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_HIDEBUTTON, idButton, MAKELONG(fShow, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_HitTest (BYVAL hToolBar AS HWND, BYVAL pptHitTest AS POINT PTR) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_HITTEST, 0, CAST(LPARAM, pptHitTest))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_Indeterminate (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL fIndeterminate AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_INDETERMINATE, cast(WPARAM, idButton), MAKELONG(fIndeterminate, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_InsertMarkHitTest (BYVAL hToolBar AS HWND, BYVAL lppt AS POINT PTR, BYVAL lptbim AS TBINSERTMARK PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_INSERTMARKHITTEST, CAST(WPARAM, lppt), CAST(LPARAM, lptbim))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_IsButtonChecked (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ISBUTTONCHECKED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_IsButtonEnabled (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ISBUTTONENABLED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_IsButtonHidden (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ISBUTTONHIDDEN, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_IsButtonHighlighted (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ISBUTTONHIGHLIGHTED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_IsButtonIndeterminate (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ISBUTTONINDETERMINATE, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_IsButtonPressed (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_ISBUTTONPRESSED, cast(WPARAM, idButton), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_LoadImages (BYVAL hToolBar AS HWND, BYVAL iBitmapID AS DWORD, BYVAL hinst AS HINSTANCE) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_LOADIMAGES, cast(WPARAM, iBitmapID), CAST(LPARAM, hinst))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_MapAccelerator (BYVAL hToolBar AS HWND, BYVAL chAccel AS DWORD, BYVAL pIDBtn AS DWORD PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_MAPACCELERATOR, cast(WPARAM, chAccel), CAST(LPARAM, pIDBtn))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_MarkButton (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL fHighlight AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_MARKBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(fHighlight, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_MoveButton (BYVAL hToolBar AS HWND, BYVAL uOldPos AS DWORD, BYVAL uNewPos AS DWORD) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_MOVEBUTTON, cast(WPARAM, uOldPos), cast(LPARAM, uNewPos))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_PressButton (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL fPress AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_PRESSBUTTON, cast(WPARAM, idButton), cast(LPARAM, MAKELONG(fPress, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_ReplaceBitmap (BYVAL hToolBar AS HWND, BYVAL ptbrb AS TBREPLACEBITMAP PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_REPLACEBITMAP, 0, CAST(LPARAM, ptbrb))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SaveRestore (BYVAL hToolBar AS HWND, BYVAL fSave AS LONG, BYVAL ptbsp AS TBSAVEPARAMSW PTR)
   SendMessageW(hToolBar, TB_SAVERESTOREW, cast(WPARAM, fSave), CAST(LPARAM, ptbsp))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetAnchorHighlight (BYVAL hToolBar AS HWND, BYVAL fAnchor AS BOOLEAN) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETANCHORHIGHLIGHT, fAnchor, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetBitmapSize (BYVAL hToolBar AS HWND, BYVAL dxBitmap AS SHORT, BYVAL dyBitmap AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETBITMAPSIZE, 0, MAKELONG(dxBitmap, dyBitmap))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetButtonInfo (BYVAL hToolBar AS HWND, BYVAL iID AS LONG, BYVAL lptbbi AS TBBUTTONINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETBUTTONINFOW, cast(WPARAM, iID), CAST(LPARAM, lptbbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetButtonSize (BYVAL hToolBar AS HWND, BYVAL dxButton AS SHORT, BYVAL dyButton AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETBUTTONSIZE, 0, cast(LPARAM, MAKELONG(dxButton, dyButton)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetButtonWidth (BYVAL hToolBar AS HWND, BYVAL cxMin AS SHORT, BYVAL cxMax AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETBUTTONWIDTH, 0, cast(LPARAM, MAKELONG(cxMin, cxMax)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetCmdId (BYVAL hToolBar AS HWND, BYVAL index AS SHORT, BYVAL cmdid AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETCMDID, cast(WPARAM, index), cast(LPARAM, cmdid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetColorScheme (BYVAL hToolBar AS HWND, BYVAL lpcs AS COLORSCHEME PTR)
   SendMessageW(hToolBar, TB_SETCOLORSCHEME, 0, CAST(LPARAM, lpcs))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetImageList (BYVAL hToolBar AS HWND, BYVAL himlNew AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(hToolBar, TB_SETIMAGELIST, 0, CAST(LPARAM, himlNew)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetHotImageList (BYVAL hToolBar AS HWND, BYVAL himlNewHot AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(hToolBar, TB_SETHOTIMAGELIST, 0, CAST(LPARAM, himlNewHot)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetDisabledImageList (BYVAL hToolBar AS HWND, BYVAL himlNewDisabled AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessageW(hToolBar, TB_SETDISABLEDIMAGELIST, 0, CAST(LPARAM, himlNewDisabled)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetPressedImageList (BYVAL hToolBar AS HWND, BYVAL iImageID AS DWORD, BYVAL himlNew AS HIMAGELIST) AS HIMAGELIST
   FUNCTION = CAST(HIMAGELIST, SendMessagew(hToolBar, TB_SETPRESSEDIMAGELIST, cast(WPARAM, iImageID), CAST(LPARAM, himlNew)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetDrawTextFlags (BYVAL hToolBar AS HWND, BYVAL dwMask AS DWORD, BYVAL dwDTFlags AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_SETDRAWTEXTFLAGS, cast(WPARAM, dwMask), cast(LPARAM, dwDTFlags))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetExtendedStyle (BYVAL hToolBar AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_SETEXTENDEDSTYLE, 0, cast(LPARAM, dwExStyle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetHotItem (BYVAL hToolBar AS HWND, BYVAL iHot AS LONG) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_SETHOTITEM, cast(WPARAM, iHot), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetIndent (BYVAL hToolBar AS HWND, BYVAL iIndent AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETINDENT, cast(WPARAM, iIndent), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetInsertMark (BYVAL hToolBar AS HWND, BYVAL lptbim AS TBINSERTMARK PTR)
   SendMessageW(hToolBar, TB_SETINSERTMARK, 0, CAST(LPARAM, lptbim))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetInsertMarkColor (BYVAL hToolBar AS HWND, BYVAL clrInsertMark AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(hToolBar, TB_SETINSERTMARKCOLOR, cast(WPARAM, clrInsertMark), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetListGap (BYVAL hToolBar AS HWND, BYVAL gap AS DWORD)
   SendMessageW(hToolBar, TB_SETLISTGAP, gap, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetMaxTextRows (BYVAL hToolBar AS HWND, BYVAL iMaxRows AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETMAXTEXTROWS, cast(WPARAM, iMaxRows), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetMetrics (BYVAL hToolBar AS HWND, BYVAL ptbMetrics AS TBMETRICS PTR)
   SendMessageW(hToolBar, TB_SETMETRICS, 0, CAST(LPARAM, ptbmetrics))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetPadding (BYVAL hToolBar AS HWND, BYVAL cx AS SHORT, BYVAL cy AS SHORT) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_SETPADDING, 0, MAKELONG(cx, cy))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetParent (BYVAL hToolBar AS HWND, BYVAL hwndParent AS HWND) AS HWND
   FUNCTION = CAST(HWND, SendMessageW(hToolBar, TB_SETPARENT, CAST(WPARAM, hwndParent), 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetRows (BYVAL hToolBar AS HWND, BYVAL cRows AS WORD, BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR)
   SendMessageW(hToolBar, TB_SETROWS, MAKELONG(cRows, fLarger), CAST(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetState (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL fState AS SHORT) AS BOOLEAN
   FUNCTION = SendMessageW(hToolBar, TB_SETSTATE, cast(WPARAM, idButton), MAKELONG(fState, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetStyle (BYVAL hToolBar AS HWND, BYVAL dwStyle AS DWORD)
   SendMessageW(hToolBar, TB_SETSTYLE, 0, dwStyle)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetToolTips (BYVAL hToolBar AS HWND, BYVAL hwndToolTip AS HWND)
   SendMessageW(hToolBar, TB_SETTOOLTIPS, CAST(WPARAM, hwndToolTip), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_SetUnicodeFormat (BYVAL hToolBar AS HWND, BYVAL fUnicode AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hToolBar, TB_SETUNICODEFORMAT, cast(WPARAM, fUnicode), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_SetWindowTheme (BYVAL hToolBar AS HWND, BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(hToolBar, TB_SETWINDOWTHEME, CAST(WPARAM, pwszTheme), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Toolbar_MakeOpaque (BYVAL hToolBar AS HWND)
   SetWindowLongPtrW(hToolBar, GWL_STYLE, GetWindowLongPtrW(hToolBar, GWL_STYLE) AND NOT TBSTYLE_TRANSPARENT)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_GetButtonText (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG, BYVAL pBuffer AS WSTRING PTR) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_GETBUTTONTEXTW, cast(WPARAM, idButton), CAST(LPARAM, pbuffer))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Toolbar_ButtonGetTextLength (BYVAL hToolBar AS HWND, BYVAL idButton AS LONG) AS LONG
   FUNCTION = SendMessageW(hToolBar, TB_GETBUTTONTEXTW, cast(WPARAM, idButton), 0)
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

' ========================================================================================
PRIVATE SUB Trackbar_ClearSel (BYVAL hTrackbar AS HWND, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_CLEARSEL, cast(WPARAM, fRedraw), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_ClearTics (BYVAL hTrackbar AS HWND, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_CLEARTICS, cast(WPARAM, fRedraw), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetBuddy (BYVAL hTrackbar AS HWND, BYVAL fLocation AS LONG) AS HWND
   FUNCTION = cast(HWND, SendMessage(hTrackbar, TBM_GETBUDDY, cast(WPARAM, fLocation), 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_GetChannelRect (BYVAL hTrackbar AS HWND, BYVAL lprc AS RECT PTR)
   SendMessageW(hTrackbar, TBM_GETCHANNELRECT, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetLineSize (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETLINESIZE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetNumTics (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETNUMTICS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetPageSize (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETPAGESIZE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetPos (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetPTics (BYVAL hTrackbar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hTrackbar, TBM_GETPTICS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetRangeMax (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETRANGEMAX, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetRangeMin (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETRANGEMIN, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetSelEnd (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETSELEND, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetSelStart (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETSELSTART, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetThumbLength (BYVAL hTrackbar AS HWND) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETTHUMBLENGTH, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_GetThumbRect (BYVAL hTrackbar AS HWND, BYVAL lprc AS RECT PTR)
   SendMessageW(hTrackbar, TBM_GETTHUMBRECT, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetTic (BYVAL hTrackbar AS HWND, BYVAL iTic AS DWORD) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETTIC, cast(WPARAM, iTic), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetTicPos (BYVAL hTrackbar AS HWND, BYVAL iTic AS DWORD) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_GETTICPOS, cast(WPARAM, iTic), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetTooltips (BYVAL hTrackbar AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hTrackbar, TBM_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_GetUnicodeFormat (BYVAL hTrackbar AS HWND) AS DWORD
   FUNCTION = SendMessageW(hTrackbar, TBM_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_SetBuddy (BYVAL hTrackbar AS HWND, BYVAL fLocation AS LONG, BYVAL hwndBuddy AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hTrackbar, TBM_SETBUDDY, cast(WPARAM, fLocation), cast(LPARAM, hwndBuddy)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_SetLineSize (BYVAL hTrackbar AS HWND, BYVAL lLineSize AS LONG) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_SETLINESIZE, 0, lLineSize)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_SetPageSize (BYVAL hTrackbar AS HWND, BYVAL lPageSize AS LONG) AS LONG
   FUNCTION = SendMessage(hTrackbar, TBM_SETPAGESIZE, 0, lPageSize)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetPos (BYVAL hTrackbar AS HWND, BYVAL fPosition AS LONG, BYVAL lPosition AS LONG)
   SendMessageW(hTrackbar, TBM_SETPOS, cast(WPARAM, fPosition), cast(LPARAM, lPosition))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetRange (BYVAL hTrackbar AS HWND, BYVAL lMinimum AS LONG, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_SETRANGE, cast(WPARAM, fRedraw), cast(LPARAM, MAKELONG(lMinimum, lMaximum)))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetRangeMax (BYVAL hTrackbar AS HWND, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_SETRANGEMAX, cast(WPARAM, fRedraw), cast(LPARAM, lMaximum))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetRangeMin (BYVAL hTrackbar AS HWND, BYVAL lMinimum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_SETRANGEMIN, cast(WPARAM, fRedraw), cast(LPARAM, lMinimum))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetSel (BYVAL hTrackbar AS HWND, BYVAL lMinimum AS LONG, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_SETSEL, cast(WPARAM, fRedraw), cast(LPARAM, MAKELONG(lMinimum, lMaximum)))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetSelEnd (BYVAL hTrackbar AS HWND, BYVAL lEnd AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_SETSELEND, cast(WPARAM, fRedraw), cast(LPARAM, lEnd))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetSelStart (BYVAL hTrackbar AS HWND, BYVAL lStart AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTrackbar, TBM_SETSELSTART, cast(WPARAM, fRedraw), cast(LPARAM, lStart))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetThumbLength (BYVAL hTrackbar AS HWND, BYVAL iLength AS LONG)
   SendMessageW(hTrackbar, TBM_SETTHUMBLENGTH, cast(WPARAM, iLength), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_SetTic (BYVAL hTrackbar AS HWND, BYVAL lPosition AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(hTrackbar, TBM_SETTIC, 0, cast(WPARAM, lPosition))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetTicFreq (BYVAL hTrackbar AS HWND, BYVAL wFreq AS WORD)
   SendMessageW(hTrackbar, TBM_SETTICFREQ, cast(WPARAM, wFreq), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_SetTipSide (BYVAL hTrackbar AS HWND, BYVAL fLocation AS LONG) AS LONG
   FUNCTION = SendMessageW(hTrackbar, TBM_SETTIPSIDE, cast(WPARAM, fLocation), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Trackbar_SetToolTips (BYVAL hTrackbar AS HWND, BYVAL hwndTT AS HWND)
   SendMessageW(hTrackbar, TBM_SETTOOLTIPS, cast(WPARAM, hwndTT), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Trackbar_SetUnicodeFormat (BYVAL hTrackbar AS HWND, BYVAL fUnicode AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hTrackbar, TBM_SETUNICODEFORMAT, fUnicode, 0)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** SCROLLBAR ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION ScrollBar_SetRangeRedraw (BYVAL hScrollBar AS HWND, BYVAL posMin AS LONG, BYVAL posMax AS LONG) AS LONG
   FUNCTION = SendMessageW(hScrollBar, SBM_SETRANGEREDRAW, posMin, posMax)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** SYSLINK ***
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION SysLink_GetIdealHeight (BYVAL hLM AS HWND) AS LONG
   FUNCTION = SendMessageW(hLM, LM_GETIDEALHEIGHT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Syslink_GetItem (BYVAL hLM AS HWND, BYVAL lpItem AS LITEM PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_GETITEM, 0, cast(LPARAM, lpItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION SysLink_HitTest (BYVAL hLM AS HWND, BYVAL lphti AS LHITTESTINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_HITTEST, 0, cast(LPARAM, lphti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION SysLink_SetItem (BYVAL hLM AS HWND, BYVAL lpItem AS LITEM PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_SETITEM, 0, cast(LPARAM, lpItem))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                   *** TOOLTIPS ***
' ########################################################################################

' ========================================================================================
PRIVATE SUB Tooltip_Activate (BYVAL hTooltip AS HWND, BYVAL fActivate AS BOOLEAN)
   SendMessageW(hTooltip, TTM_ACTIVATE, fActivate, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_AddTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_ADDTOOLW, 0, CAST(LPARAM, lpti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_AdjustRect (BYVAL hTooltip AS HWND, BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_ADJUSTRECT, cast(WPARAM, fLarger), cast(LPARAM, lprc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_DelTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(hTooltip, TTM_DELTOOLW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_EnumTools (BYVAL hTooltip AS HWND, BYVAL iTool AS DWORD, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_ENUMTOOLSW, cast(WPARAM, iTool), cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetBubbleSize (BYVAL hTooltip AS HWND, BYVAL lpTtm AS TOOLINFOW PTR) AS DWORD
   FUNCTION = SendMessageW(hTooltip, TTM_GETBUBBLESIZE, 0, cast(LPARAM, lpTtm))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetCurrentTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_GETCURRENTTOOLW, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetDelayTime (BYVAL hTooltip AS HWND, BYVAL dwDuration AS DWORD) AS LONG
   FUNCTION = SendMessageW(hTooltip, TTM_GETDELAYTIME, cast(WPARAM, dwDuration), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetFont (BYVAL hTooltip AS HWND) AS HFONT
   FUNCTION = cast(HFONT, SendMessageW(hTooltip, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_GetMargin (BYVAL hTooltip AS HWND, BYVAL lprc AS RECT PTR)
   SendMessageW(hTooltip, TTM_GETMARGIN, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetMaxTipWidth (BYVAL hTooltip AS HWND) AS LONG
   FUNCTION = SendMessageW(hTooltip, TTM_GETMAXTIPWIDTH, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_GetText (BYVAL hTooltip AS HWND, BYVAL numChars AS DWORD, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(hTooltip, TTM_GETTEXTW, numChars, cast(LPARAM, lpti))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetTipBkColor (BYVAL hTooltip AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hTooltip, TTM_GETTIPBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetTipTextColor (BYVAL hTooltip AS HWND) AS COLORREF
   FUNCTION = SendMessageW(hTooltip, TTM_GETTIPTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_GetTitle (BYVAL hTooltip AS HWND, BYVAL pTTGetTitle AS TTGETTITLE PTR)
   SendMessageW(hTooltip, TTM_GETTITLE, 0, cast(LPARAM, pTTGetTitle))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetToolCount (BYVAL hTooltip AS HWND) AS LONG
   FUNCTION = SendMessageW(hTooltip, TTM_GETTOOLCOUNT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_GetToolInfo (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_GETTOOLINFOW, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_HitTest (BYVAL hTooltip AS HWND, BYVAL lphti AS TTHITTESTINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_HITTEST, 0, cast(LPARAM, lphti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_NewToolRect (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(hTooltip, TTM_NEWTOOLRECTW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_Pop (BYVAL hTooltip AS HWND)
   SendMessageW hTooltip, TTM_POP, 0, 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_Popup (BYVAL hTooltip AS HWND)
   SendMessageW hTooltip, TTM_POPUP, 0, 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_RelayEvent (BYVAL hTooltip AS HWND, BYVAL lpmsg AS tagMSG PTR) AS LONG
   FUNCTION = SendMessageW(hTooltip, TTM_RELAYEVENT, 0, cast(LPARAM, lpmsg))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetDelayTime (BYVAL hTooltip AS HWND, BYVAL dwDuration AS DWORD, BYVAL iTime AS SHORT)
   SendMessageW(hTooltip, TTM_SETDELAYTIME, dwDuration, MAKELONG(iTime, 0))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetFont (BYVAL hTooltip AS HWND, BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(hTooltip, WM_SETFONT, cast(WPARAM, hFont), cast(LPARAM, fRedraw))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetMargin (BYVAL hTooltip AS HWND, BYVAL lprc AS RECT PTR)
   SendMessageW(hTooltip, TTM_SETMARGIN, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_SetMaxTipWidth (BYVAL hTooltip AS HWND, BYVAL iWidth AS LONG) AS LONG
   FUNCTION = SendMessageW(hTooltip, TTM_SETMAXTIPWIDTH, 0, cast(LPARAM, iWidth))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetTipBkColor (BYVAL hTooltip AS HWND, BYVAL clr AS COLORREF)
   SendMessageW(hTooltip, TTM_SETTIPBKCOLOR, cast(WPARAM, clr), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetTipTextColor (BYVAL hTooltip AS HWND, BYVAL clr AS COLORREF)
   SendMessageW(hTooltip, TTM_SETTIPTEXTCOLOR, cast(WPARAM, clr), 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_SetTitle (BYVAL hTooltip AS HWND, BYVAL nIcon AS LONG, BYVAL pwszTitle AS WSTRING PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hTooltip, TTM_SETTITLEW, nIcon, cast(LPARAM, pwszTitle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetToolInfo (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(hTooltip, TTM_SETTOOLINFOW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_SetWindowTheme (BYVAL hTooltip AS HWND, BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(hTooltip, TTM_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_TrackActivate (BYVAL hTooltip AS HWND, BYVAL bActivate AS BOOLEAN, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(hTooltip, TTM_TRACKACTIVATE, cast(WPARAM, bActivate), cast(LPARAM, lpti))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_TrackPosition (BYVAL hTooltip AS HWND, BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   SendMessageW(hTooltip, TTM_TRACKPOSITION, 0, cast(LPARAM, MAKELONG(xPos, yPos)))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_Update (BYVAL hTooltip AS HWND)
   SendMessageW(hTooltip, TTM_UPDATE, 0, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB Tooltip_UpdateTipText (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(hTooltip, TTM_UPDATETIPTEXTW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Tooltip_WindowFromPoint (BYVAL hTooltip AS HWND, BYVAL lppt AS POINT PTR) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hTooltip, TTM_WINDOWFROMPOINT, 0, cast(LPARAM, lppt)))
END FUNCTION
' ========================================================================================

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

' ========================================================================================
PRIVATE FUNCTION UpDown_GetAccel (BYVAL hUDM AS HWND, BYVAL cAccels AS DWORD, BYVAL paAccels AS UDACCEL PTR) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_GETACCEL, cAccels, CAST(LPARAM, paAccels))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_GetBase (BYVAL hUDM AS HWND) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_GETBASE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_GetBuddy (BYVAL hUDM AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hUDM, UDM_GETBUDDY, 0, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_GetPos (BYVAL hUDM AS HWND) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_GetPos32 (BYVAL hUDM AS HWND, BYVAL pfError AS LONG PTR = NULL) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_GETPOS32, 0, cast(LPARAM, pfError))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_GetRange (BYVAL hUDM AS HWND) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_GETRANGE, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB UpDown_GetRange32 (BYVAL hUDM AS HWND, BYVAL pLow AS LONG PTR, BYVAL pHigh AS LONG PTR)
   SendMessageW(hUDM, UDM_GETRANGE32, cast(WPARAM, pLow), cast(LPARAM, pHigh))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_GetUnicodeFormat (BYVAL hUDM AS HWND) AS DWORD
   FUNCTION = SendMessageW(hUDM, UDM_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_SetAccel (BYVAL hUDM AS HWND, BYVAL cAccels AS DWORD, BYREF paAccels AS UDACCEL PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hUDM, UDM_SETACCEL, cast(WPARAM, cAccels), cast(LPARAM, paAccels))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_SetBase (BYVAL hUDM AS HWND, BYVAL nBase AS LONG) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_SETBASE, cast(WPARAM, nBase), 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_SetBuddy (BYVAL hUDM AS HWND, BYVAL hwndBuddy AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(hUDM, UDM_SETBUDDY, cast(WPARAM, hwndBuddy), 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_SetPos (BYVAL hUDM AS HWND, BYVAL nPos AS SHORT) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_SETPOS, 0, cast(LPARAM, MAKELONG(nPos, 0)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_SetPos32 (BYVAL hUDM AS HWND, BYVAL nPos AS LONG) AS LONG
   FUNCTION = SendMessageW(hUDM, UDM_SETPOS32, 0, cast(LPARAM, nPos))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB UpDown_SetRange (BYVAL hUDM AS HWND, BYVAL nUpper AS SHORT, BYVAL nLower AS SHORT)
   SendMessageW(hUDM, UDM_SETRANGE, 0, MAKELONG(nUpper, nLower))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB UpDown_SetRange32 (BYVAL hUDM AS HWND, BYVAL iLow AS LONG, BYVAL iHigh AS LONG)
   SendMessageW(hUDM, UDM_SETRANGE32, cast(WPARAM, iLow), cast(LPARAM, iHigh))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION UpDown_SetUnicodeFormat (BYVAL hUDM AS HWND, BYVAL fUnicode AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hUDM, UDM_SETUNICODEFORMAT, cast(WPARAM, fUnicode), 0)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                   *** PROPSHEET ***
' ########################################################################################

#if _WIN32_WINNT = &h0602

'PSM_SETNEXTTEXTW         = (WM_USER + 137)
'PSM_SETNEXTTEXT          = PSM_SETNEXTTEXTW
PRIVATE SUB PropSheet_SetNextText (BYVAL hDlg AS HWND, BYVAL pwszText AS WSTRING PTR)
   SendMessage(hDlg, WM_USER + 137, 0, cast(LPARAM, pwszText))
END SUB

'PSWIZB_SHOW                    = 0
'PSWIZB_RESTORE                 = 1
'PSM_SHOWWIZBUTTONS             = (WM_USER + 138)
PRIVATE SUB PropSheet_ShowWizButtons (BYVAL hDlg AS HWND, BYVAL dwFlag AS DWORD, BYVAL dwButton AS DWORD)
   SendMessage(hDlg, WM_USER + 138, cast(WPARAM, dwFlag), cast(LPARAM, dwButton))
END SUB

'PSM_ENABLEWIZBUTTONS           = (WM_USER + 139)
PRIVATE SUB PropSheet_EnableWizButtons (BYVAL hDlg AS HWND, BYVAL dwState AS DWORD, BYVAL dwMask AS DWORD)
   SendMessage(hDlg, WM_USER + 139, cast(WPARAM, dwState), cast(LPARAM, dwMask))
END SUB

'PSM_SETBUTTONTEXTW         = (WM_USER + 140)
'PSM_SETBUTTONTEXT          = PSM_SETBUTTONTEXTW
PRIVATE SUB PropSheet_SetButtonText (BYVAL hDlg AS HWND, BYVAL dwButton AS DWORD, BYVAL pwszText AS WSTRING PTR)
   SendMessage(hDlg, WM_USER + 140, cast(WPARAM, dwButton), cast(LPARAM, pwszText))
END SUB

#endif
