' ########################################################################################
' Platform: Microsoft Windows
' Filename: CComboBox.inc
' Purpose:  Windows common controls - ComboBox / ComboBoxEx
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
' CComboBox class
' ########################################################################################

TYPE CComboBox EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION Enable () AS BOOLEAN
   DECLARE FUNCTION Disable () AS BOOLEAN
   DECLARE FUNCTION AddString (BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION DeleteString (BYVAL Index AS LONG) AS LONG
   DECLARE FUNCTION Dir (BYVAL Attr AS UINT, BYVAL pwszFileSpec AS WSTRING PTR) AS LONG
   DECLARE FUNCTION FindString (BYVAL start AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION FindStringExact (BYVAL start AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION GetComboBoxInfo (BYREF cbi AS COMBOBOXINFO) AS BOOLEAN
   DECLARE FUNCTION GetEditBoxHandle () AS HWND
   DECLARE FUNCTION GetListBoxHandle () AS HWND
   DECLARE FUNCTION GetCount () AS LONG
   DECLARE FUNCTION GetCueBannerText (BYREF wszText AS WSTRING, BYVAL cchText AS DWORD) AS BOOLEAN
   DECLARE FUNCTION GetCurSel () AS LONG
   DECLARE SUB GetDroppedControlRect (BYREF rc AS RECT)
   DECLARE FUNCTION GetDroppedControlRect () AS RECT
   DECLARE FUNCTION GetDroppedState () AS BOOLEAN
   DECLARE FUNCTION GetDroppedWidth () AS LONG
   DECLARE SUB GetEditSel (BYREF dwStart AS DWORD, BYREF dwEnd AS DWORD)
   DECLARE FUNCTION GetExtendedUI () AS UINT
   DECLARE FUNCTION GetHorizontalExtent () AS LONG
   DECLARE FUNCTION GetItemData (BYVAL index AS LONG) AS LRESULT
   DECLARE FUNCTION GetItemHeight () AS LONG
   DECLARE FUNCTION GetLbTextLen (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION GetLbText (BYVAL Index AS LONG) AS DWSTRING
   DECLARE FUNCTION GetLocale () AS LONG
   DECLARE FUNCTION GetMinVisible () AS LONG
   DECLARE FUNCTION GetTopIndex () AS LONG
   DECLARE FUNCTION InitStorage (BYVAL items AS LONG, BYVAL cb AS LONG) AS LONG
   DECLARE FUNCTION InsertString (BYVAL Index AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE SUB LimitText (BYVAL chLimit AS LONG)
   DECLARE SUB ResetContent
   DECLARE FUNCTION SelectString (BYVAL Index AS DWORD, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION SetCueBannerText (BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION SetCurSel (BYVAL Index AS DWORD) AS LONG
   DECLARE FUNCTION SetDroppedWidth (BYVAL nWidth AS LONG) AS LONG
   DECLARE FUNCTION SetEditSel (BYVAL nStart AS INTEGER, BYVAL nEnd AS INTEGER) AS LONG
   DECLARE FUNCTION SetExtendedUI (BYVAL flags AS UINT) AS LONG
   DECLARE SUB SetHorizontalExtent (BYVAL nWidth AS LONG)
   DECLARE FUNCTION SetItemData (BYVAL index AS DWORD, BYVAL nData AS LPARAM) AS LONG
   DECLARE FUNCTION SetItemHeight (BYVAL index AS LONG, BYVAL nHeight AS LONG) AS LONG
   DECLARE FUNCTION SetLocale (BYVAL hComboBox AS HWND, BYVAL lcid AS LONG) AS LONG
   DECLARE FUNCTION SetMinVisible (BYVAL iMinVisible AS LONG) AS LONG
   DECLARE FUNCTION SetTopIndex (BYVAL index AS DWORD) AS LONG
   DECLARE SUB ShowDropDown
   DECLARE SUB HideDropDown
   DECLARE FUNCTION CalcHorizontalExtent () AS LONG
   DECLARE FUNCTION AddItemData (BYVAL pData AS LRESULT) AS LONG
   DECLARE FUNCTION InsertItemData (BYVAL index AS LONG, BYVAL pData AS LRESULT) AS LONG
   DECLARE FUNCTION FindItemData (BYVAL indexStart AS LONG, BYVAL pData AS LRESULT) AS LONG
   DECLARE FUNCTION SelectItemData (BYVAL indexStart AS LONG, BYVAL pData AS LRESULT) AS LONG
   DECLARE FUNCTION GetTextLength () AS LONG
   DECLARE FUNCTION GetText () AS DWSTRING
   DECLARE FUNCTION SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CComboBox constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CComboBox (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CComboBox (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CComboBox (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CComboBox (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CComboBox destructor
' ========================================================================================
PRIVATE DESTRUCTOR CComboBox
END DESTRUCTOR
' ========================================================================================

' ########################################################################################
'                                    *** COMBOBOX ***
' ########################################################################################

' ========================================================================================
' Enables a combobox
' ========================================================================================
PRIVATE FUNCTION CComboBox.Enable () AS BOOLEAN
   FUNCTION = EnableWindow(m_hCtl, CTRUE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Disables a combobox
' ========================================================================================
PRIVATE FUNCTION CComboBox.Disable () AS BOOLEAN
   DIM res AS LONG = EnableWindow(m_hCtl, FALSE)
   DIM bEnable AS BOOLEAN
   IF res = 0 THEN bEnable = TRUE ELSE bEnable = FALSE
   RETURN bEnable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a string to the list box of a combo box. If the combo box does not have the
' CBS_SORT style, the string is added to the end of the list.
' ========================================================================================
PRIVATE FUNCTION CCombobox.AddString (BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_ADDSTRING, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a string in the list box of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.DeleteString (BYVAL Index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, CB_DELETESTRING, Index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds names to the list displayed by the combo box. The message adds the names of
' directories and files that match a specified string and set of file attributes.
' CB_DIR can also add mapped drive letters to the list.
' ========================================================================================
PRIVATE FUNCTION CCombobox.Dir (BYVAL Attr AS UINT, BYVAL pwszFileSpec AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_DIR, Attr, CAST(LPARAM, pwszFileSpec))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches the list box of a combo box for an item beginning with the characters in a
' specified string.
' ========================================================================================
PRIVATE FUNCTION CCombobox.FindString (BYVAL start AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_FINDSTRING, start, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the first list box string in a combo box that matches the string specified in the
' start parameter.
' ========================================================================================
PRIVATE FUNCTION CCombobox.FindStringExact (BYVAL start AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_FINDSTRINGEXACT, start, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the specified combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetComboBoxInfo (BYREF cbi AS COMBOBOXINFO) AS BOOLEAN
   IF VARPTR(cbi) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, CB_GETCOMBOBOXINFO, 0, CAST(LPARAM, @cbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the edit box of the combobox
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetEditBoxHandle () AS HWND
   DIM tci AS COMBOBOXINFO
   tci.cbSize = SIZEOF(COMBOBOXINFO)
   SendMessageW(m_hCtl, CB_GETCOMBOBOXINFO, 0, CAST(LPARAM, @tci))
   RETURN tci.hwndItem
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the drop down of the combobox
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetListBoxHandle () AS HWND
   DIM tci AS COMBOBOXINFO
   tci.cbSize = SIZEOF(COMBOBOXINFO)
   SendMessageW(m_hCtl, CB_GETCOMBOBOXINFO, 0, CAST(LPARAM, @tci))
   RETURN tci.hwndList
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of items in the list box of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetCount () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the cue banner text displayed in the edit control of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetCueBannerText (BYREF wszText AS WSTRING, BYVAL cchText AS DWORD) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, CB_GETCUEBANNER, CAST(LPARAM, @wszText), cchText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the index of the currently selected item, if any, in the list box of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetCurSel () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETCURSEL, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the screen coordinates of a combo box in its dropped-down state
' ========================================================================================
PRIVATE SUB CCombobox.GetDroppedControlRect (BYREF rc AS RECT)
   SendMessageW(m_hCtl, CB_GETDROPPEDCONTROLRECT,  0, CAST(LPARAM, @rc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetDroppedControlRect () AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, CB_GETDROPPEDCONTROLRECT,  0, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the list box of a combo box is dropped down.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetDroppedState () AS BOOLEAN
   RETURN CBOOL(SendMessageW(m_hCtl, CB_GETDROPPEDSTATE, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the minimum allowable width, in pixels, of the list box of a combo box with
' the CBS_DROPDOWN or CBS_DROPDOWNLIST style.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetDroppedWidth () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETDROPPEDWIDTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the starting and ending character positions of the current selection in the edit
' control of a combo box.
' ========================================================================================
PRIVATE SUB CComboBox.GetEditSel (BYREF dwStart AS DWORD, BYREF dwEnd AS DWORD)
   IF VARPTR(dwStart) = 0 OR VARPTR(dwEnd) = 0 THEN EXIT SUB
   SendMessageW m_hCtl, CB_GETEDITSEL, CAST(WPARAM, @dwStart), CAST(LPARAM, @dwEnd)
END SUB
' ========================================================================================

' ========================================================================================
' Determines whether a combo box has the default user interface or the extended user interface.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetExtendedUI () AS UINT
   RETURN SendMessageW(m_hCtl, CB_GETEXTENDEDUI, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves from a combo box the width, in pixels, by which the list box can be scrolled
' horizontally (the scrollable width). This is applicable only if the list box has a
' horizontal scroll bar.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetHorizontalExtent () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETHORIZONTALEXTENT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the application-supplied value associated with the specified item in the combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetItemData (BYVAL index AS LONG) AS LRESULT
   RETURN SendMessageW(m_hCtl, CB_GETITEMDATA, IIF(index < 0, this.GetCurSel, index), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the height of list items in a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetItemHeight () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETITEMHEIGHT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieve the length, in characters, of a string in the list of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetLbTextLen (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETLBTEXTLEN, Index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a string from the list of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetLbText (BYVAL Index AS LONG) AS DWSTRING
   DIM dwsText AS DWSTRING
   DIM nLen AS LONG
   IF Index < 0 THEN Index = this.GetCurSel
   nLen = this.GetLbTextLen(Index)
   dwsText = WSPACE(nLen + 1)
   SendMessageW m_hCtl, CB_GETLBTEXT, Index, CAST(LPARAM, STRPTR(dwsText))
   RETURN LEFT(dwsText, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current locale of the combo box. The locale is used to determine the
' correct sorting order of displayed text for combo boxes with the CBS_SORT style and text
' added by using the CB_ADDSTRING message.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetLocale () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETLOCALE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the minimum number of visible items in the drop-down list of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetMinVisible () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETMINVISIBLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the zero-based index of the first visible item in the list box portion of a
' combo box. Initially, the item with index 0 is at the top of the list box, but if the
' list box contents have been scrolled, another item may be at the top.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetTopIndex () AS LONG
   RETURN SendMessageW(m_hCtl, CB_GETTOPINDEX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a string into the list box of a combo box. Unlike the CB_ADDSTRING message, the
' CB_INSERTSTRING message does not cause a list with the CBS_SORT style to be sorted.
' ========================================================================================
PRIVATE FUNCTION CCombobox.InsertString (BYVAL Index AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_INSERTSTRING, Index, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Limits the length of the text the user may type into the edit control of a combo box.
' ========================================================================================
PRIVATE SUB CComboBox.LimitText (BYVAL chLimit AS LONG)
   SendMessage m_hCtl, CB_LIMITTEXT, chLimit, 0
END SUB
' ========================================================================================

' ========================================================================================
' Removes all items from the list box and edit control of a combo box.
' ========================================================================================
PRIVATE SUB CComboBox.ResetContent
   SendMessage m_hCtl, CB_RESETCONTENT, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Searches the list of a combo box for an item that begins with the characters in a
' specified string. If a matching item is found, it is selected and copied to the edit control.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SelectString (BYVAL Index AS DWORD, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SELECTSTRING, Index, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the cue banner text displayed in the edit control of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetCueBannerText (BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETCUEBANNER, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a string in the list of a combo box. If necessary, the list scrolls the string
' into view. The text in the edit control of the combo box changes to reflect the new
' selection, and any previous selection in the list is removed.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetCurSel (BYVAL Index AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETCURSEL, Index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the maximum allowable width, in pixels, of the list box of a combo box with the
' CBS_DROPDOWN or CBS_DROPDOWNLIST style.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetDroppedWidth (BYVAL nWidth AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETDROPPEDWIDTH, nWidth, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects characters in the edit control of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetEditSel (BYVAL nStart AS INTEGER, BYVAL nEnd AS INTEGER) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETEDITSEL, 0, MAKELONG(nStart, nEnd))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects either the default user interface or the extended user interface for a combo box
' that has the CBS_DROPDOWN or CBS_DROPDOWNLIST style.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetExtendedUI (BYVAL flags AS UINT) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETEXTENDEDUI, flags, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width, in pixels, by which a list box can be scrolled horizontally (the
' scrollable width). If the width of the list box is smaller than this value, the
' horizontal scroll bar horizontally scrolls items in the list box. If the width of the
' list box is equal to or greater than this value, the horizontal scroll bar is hidden or,
' if the combo box has the CBS_DISABLENOSCROLL style, disabled.
' ========================================================================================
PRIVATE SUB CComboBox.SetHorizontalExtent (BYVAL nWidth AS LONG)
   SendMessage m_hCtl, CB_SETHORIZONTALEXTENT, nWidth, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the value associated with the specified item in a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetItemData (BYVAL index AS DWORD, BYVAL nData AS LPARAM) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETITEMDATA, index, nData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the height of list items or the selection field in a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetItemHeight (BYVAL index AS LONG, BYVAL nHeight AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETITEMHEIGHT, index, nHeight)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current locale of the combo box. If the combo box has the CBS_SORT style and
' strings are added using CB_ADDSTRING, the locale of a combo box affects how list items
' are sorted.  The return value is the previous locale identifier. If lcid specifies a
' locale not installed on the system, the return value is CB_ERR and the current combo box
' locale is not changed.
' ========================================================================================
PRIVATE FUNCTION CComboBox.SetLocale (BYVAL hComboBox AS HWND, BYVAL lcid AS LONG) AS LONG
   FUNCTION = SendMessageW(hComboBox, CB_SETLOCALE, cast(WPARAM, lcid), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum number of visible items in the drop-down list of a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetMinVisible (BYVAL iMinVisible AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SETMINVISIBLE, iMinVisible, 0)
END FUNCTION
' ========================================================================================


' ========================================================================================
' Ensures that a particular item is visible in the list box of a combo box. The system
' scrolls the list box contents so that either the specified it. If the message is successful,
' the return value is zero. If the message fails, the return value is CB_ERR.
' ========================================================================================
PRIVATE FUNCTION CComboBox.SetTopIndex (BYVAL index AS DWORD) AS LONG
   FUNCTION = SendMessageW(m_hCtl, CB_SETTOPINDEX, cast(WPARAM, index), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows or hides the list box of a combo box that has the CBS_DROPDOWN or CBS_DROPDOWNLIST style.
' ========================================================================================
PRIVATE SUB CComboBox.ShowDropDown
   SendMessage m_hCtl, CB_SHOWDROPDOWN, CTRUE, 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CComboBox.HideDropDown
   SendMessage m_hCtl, CB_SHOWDROPDOWN, FALSE, 0
END SUB
' ========================================================================================

' ========================================================================================
' Calculates the size, in logical units, of the widest string item.
' ========================================================================================
PRIVATE FUNCTION CCombobox.CalcHorizontalExtent () AS LONG
   ' // Get an handle to the device context of the control
   DIM hDC AS HDC = GetDC(m_hCtl)
   ' // Get the number of items
   DIM i AS LONG, cxScrollWidth AS LONG
   DIM cItems AS LRESULT = SendMessageW(m_hCtl, CB_GETCOUNT, 0, 0)
   FOR i = 0 TO cItems - 1
      ' // Get the text of the item
      DIM cbLen AS LONG = SendMessageW(m_hCtl, CB_GETLBTEXTLEN, i, 0)
      IF cbLen = 0 THEN CONTINUE FOR
      DIM pbuffer AS WSTRING PTR = CAllocate(cbLen * 2)
      SendMessageW(m_hCtl, CB_GETLBTEXT, i, CAST(LPARAM, pbuffer))
      ' // Computes the width and height of the text
      DIM tsize AS SIZE
      GetTextExtentPoint32W hDC, pbuffer, cbLen, @tsize
      cxScrollWidth = MAX(cxScrollWidth, tsize.cx)
      Deallocate pbuffer
   NEXT
   ' // Release the device context
   ReleaseDC m_hCtl, hDC
   ' // Return the width of the text
   FUNCTION = cxScrollWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds item data to the combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.AddItemData (BYVAL pData AS LRESULT) AS LONG
   RETURN SendMessageW(m_hCtl, CB_ADDSTRING, 0, pData)
END FUNCTION
' ========================================================================================
 '========================================================================================
' Inserts item data to a combo box.
' ========================================================================================
PRIVATE FUNCTION CCombobox.InsertItemData (BYVAL index AS LONG, BYVAL pData AS LRESULT) AS LONG
   RETURN SendMessageW(m_hCtl, CB_INSERTSTRING, index, pData)
END FUNCTION
' ========================================================================================
 '========================================================================================
' Finds the first item in a combo box that has the specified item data.
' ========================================================================================
PRIVATE FUNCTION CCombobox.FindItemData (BYVAL indexStart AS LONG, BYVAL pData AS LRESULT) AS LONG
   RETURN SendMessageW(m_hCtl, CB_FINDSTRING, indexStart, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Searches a combo box for an item that has the specified item data. If a matching item is found, the item is selected.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SelectItemData (BYVAL indexStart AS LONG, BYVAL pData AS LRESULT) AS LONG
   RETURN SendMessageW(m_hCtl, CB_SELECTSTRING, indexStart, pData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the length of the text in a combobox control.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetTextLength () AS LONG
   RETURN GetWindowTextLengthW(m_hCtl)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the text in a combobox control.
' ========================================================================================
PRIVATE FUNCTION CCombobox.GetText () AS DWSTRING
   DIM nLen AS LONG = GetWindowTextLengthW(m_hCtl)
   DIM dwsText AS DWSTRING = WSPACE(nLen + 1)
   ..GetWindowTextW(m_hCtl, dwsText, nLen +  1)
   RETURN dwsText
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Sets the text of a combobox control.
' ========================================================================================
PRIVATE FUNCTION CCombobox.SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SetWindowTextW(m_hCtl, pwszText)
END FUNCTION
' ========================================================================================

END NAMESPACE

NAMESPACE Afx2

' ########################################################################################
'                                   *** COMBOBOXEX ***
' ########################################################################################

' ========================================================================================
' CComboBoxEx class
' ========================================================================================

TYPE CComboBoxEx EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION DeleteItem (BYVAL nIndex AS DWORD) AS LONG
   DECLARE FUNCTION GetComboControl () AS HWND
   DECLARE FUNCTION GetEditControl () AS HWND
   DECLARE FUNCTION GetExtendedStyle () AS DWORD
   DECLARE FUNCTION GetImageList () AS HIMAGELIST
   DECLARE FUNCTION GetItem (BYREF CBItem AS COMBOBOXEXITEMW) AS BOOLEAN
   DECLARE FUNCTION GetUnicodeFormat () AS DWORD
   DECLARE FUNCTION HasEditChanged () AS BOOLEAN
   DECLARE FUNCTION InsertItem (BYREF CBItem AS COMBOBOXEXITEMW) AS LONG
   DECLARE FUNCTION SetExtendedStyle (BYVAL dwMask AS DWORD, BYVAL dwExStyle AS DWORD) AS DWORD
   DECLARE FUNCTION SetImageList (BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetItem (BYREF CBItem AS COMBOBOXEXITEMW) AS BOOLEAN
   DECLARE FUNCTION SetUnicodeFormat (BYVAL fUnicode AS BOOLEAN) AS BOOLEAN
   DECLARE SUB SetWindowTheme (BYVAL pwszStyle AS WSTRING PTR)

END TYPE

' ========================================================================================
' Removes an item from a ComboBoxEx control. Returns an INT value that represents the number
' of items remaining in the control. If iIndex is invalid, the message returns CB_ERR.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.DeleteItem (BYVAL nIndex AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, CBEM_DELETEITEM, cast(WPARAM, nIndex), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle to the combo box control within the ComboBoxEx control.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.GetComboControl () AS HWND
   RETURN cast(HWND, SendMessageW(m_hCtl, CBEM_GETCOMBOCONTROL, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle to the edit control within the ComboBoxEx control if it uses the
' CBS_DROPDOWN style. Otherwise, the message returns NULL.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.GetEditControl () AS HWND
   RETURN cast(HWND, SendMessageW(m_hCtl, CBEM_GETEDITCONTROL, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a DWORD value that contains the ComboBoxEx control extended styles in use for the control.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.GetExtendedStyle () AS DWORD
   RETURN SendMessageW(m_hCtl, CBEM_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to an image list assigned to a ComboBoxEx control.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.GetImageList () AS HIMAGELIST
   RETURN cast(HIMAGELIST, SendMessageW(m_hCtl, CBEM_GETIMAGELIST, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves item information for a given ComboBoxEx item. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.GetItem (BYREF CBItem AS COMBOBOXEXITEMW) AS BOOLEAN
   IF VARPTR(CBItem) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, CBEM_GETITEMW, 0, cast(LPARAM, @CBItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the Unicode format flag for the control. If this value is nonzero, the control
' is using Unicode characters. If this value is zero, the control is using ANSI characters.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.GetUnicodeFormat () AS DWORD
   RETURN SendMessageW(m_hCtl, CBEM_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether or not the user has changed the text of a ComboBoxEx edit control.
' Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.HasEditChanged () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, CBEM_HASEDITCHANGED, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the index at which the new item was inserted if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.InsertItem (BYREF CBItem AS COMBOBOXEXITEMW) AS LONG
   IF VARPTR(CBItem) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, CBEM_INSERTITEMW, 0, cast(LPARAM, @CBItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a DWORD value that contains the extended styles previously used for the control.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.SetExtendedStyle (BYVAL dwExMask AS DWORD, BYVAL dwExStyle AS DWORD) AS DWORD
   RETURN SendMessageW(m_hCtl, CBEM_SETEXTENDEDSTYLE, cast(WPARAM, dwExMask), cast(LPARAM, dwExStyle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle to the image list previously associated with the control, or returns
' NULL if no image list was previously set.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.SetImageList (BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   RETURN cast(HIMAGELIST, SendMessageW(m_hCtl, CBEM_SETIMAGELIST, 0, cast(LPARAM, himl)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the attributes for an item in a ComboBoxEx control.
' Returns nonzero if successful, or zero otherwise.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.SetItem (BYREF CBItem AS COMBOBOXEXITEMW) AS BOOLEAN
   IF VARPTR(CBItem) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, CBEM_SETITEMW, 0, cast(LPARAM, @CBItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the UNICODE character format flag for the control. This message enables you to
' change the character set used by the control at run time rather than having to re-create
' the control. Returns the previous Unicode format flag for the control.
' ========================================================================================
PRIVATE FUNCTION CComboBoxEx.SetUnicodeFormat (BYVAL fUnicode AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, CBEM_SETUNICODEFORMAT, fUnicode, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the visual style of a ComboboxEx control.
' ========================================================================================
PRIVATE SUB CComboBoxEx.SetWindowTheme (BYVAL pwszStyle AS WSTRING PTR)
   SendMessageW(m_hCtl, CBEM_SETWINDOWTHEME, 0, cast(LPARAM, pwszStyle))
END SUB
' ========================================================================================

END NAMESPACE
