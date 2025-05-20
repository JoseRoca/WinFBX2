' ########################################################################################
' Platform: Microsoft Windows
' Filename: CListBox.inc
' Purpose:  Windows common controls - ListBox
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
' CListBox class
' ########################################################################################

TYPE CListBox EXTENDS CControl

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
   DECLARE FUNCTION AddFile (BYVAL pwszFileName AS WSTRING PTR) AS LONG
   DECLARE FUNCTION AddString (BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION Add (BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION DeleteString (BYVAL Index AS LONG) AS LONG
   DECLARE FUNCTION Delete_ (BYVAL Index AS LONG) AS LONG
   DECLARE FUNCTION Dir (BYVAL Attr AS DWORD, BYVAL pwszFileSpec AS WSTRING PTR) AS LONG
   DECLARE FUNCTION FindString (BYVAL indexStart AS LONG, BYVAL pwszString AS WSTRING PTR) AS LONG
   DECLARE FUNCTION Find (BYVAL indexStart AS LONG, BYVAL pwszString AS WSTRING PTR) AS LONG
   DECLARE FUNCTION FindStringExact (BYVAL indexStart AS LONG, BYVAL pwszSearchString AS WSTRING PTR) AS LONG
   DECLARE FUNCTION FindExact (BYVAL indexStart AS LONG, BYVAL pwszSearchString AS WSTRING PTR) AS LONG
   DECLARE FUNCTION GetAnchorIndex () AS LONG
   DECLARE FUNCTION GetCaretIndex () AS LONG
   DECLARE FUNCTION GetCount () AS LONG
   DECLARE FUNCTION GetCurSel () AS LONG
   DECLARE FUNCTION GetHorizontalExtent () AS LONG
   DECLARE FUNCTION GetListBoxInfo () AS LONG
   DECLARE FUNCTION GetItemData (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION GetItemHeight (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION GetItemRect (BYVAL index AS LONG, BYREF rc AS RECT) AS LONG
   DECLARE FUNCTION GetItemRect (BYVAL index AS LONG) AS RECT
   DECLARE FUNCTION GetLocale () AS LONG
   DECLARE FUNCTION GetSel (BYVAL Index AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetState (BYVAL Index AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetSelCount () AS LONG
   DECLARE FUNCTION GetSelItems (BYVAL items AS DWORD, BYVAL rgItems AS LONG PTR) AS LONG
   DECLARE FUNCTION GetText (BYVAL Index AS LONG) AS DWSTRING
   DECLARE FUNCTION GetTextLen (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION GetTopIndex () AS LONG
   DECLARE FUNCTION InitStorage (BYVAL nItems AS DWORD, BYVAL nBytes AS DWORD) AS LONG
   DECLARE FUNCTION InsertString (BYVAL Index AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION Insert (BYVAL Index AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION ItemFromPoint (BYVAL x AS SHORT, BYVAL y AS SHORT) AS LONG
   DECLARE SUB ResetContent
   DECLARE SUB Reset
   DECLARE FUNCTION SelectString (BYVAL start AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION SelItemRange (BYVAL fSel AS LONG, BYVAL startIndex AS SHORT, BYVAL endIndex AS SHORT) AS LONG
   DECLARE FUNCTION SelItemRangeEx (BYVAL startIndex AS LONG, BYVAL endIndex AS LONG) AS LONG
   DECLARE FUNCTION SetAnchorIndex (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION SetCaretIndex (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION SetColumnWidth (BYVAL nWidth AS DWORD) AS LONG
   DECLARE FUNCTION SetCount (BYVAL nCount AS DWORD) AS LONG
   DECLARE FUNCTION SetCurSel (BYVAL Index AS LONG) AS LONG
   DECLARE SUB SetHorizontalExtent (BYVAL cxExtent AS LONG)
   DECLARE FUNCTION SetItemData (BYVAL index AS LONG, BYVAL pData AS LONG) AS LONG
   DECLARE FUNCTION SetItemHeight (BYVAL nHeight AS DWORD) AS LONG
   DECLARE FUNCTION SetLocale (BYVAL hListBox AS DWORD, BYVAL wLocaleID AS DWORD) AS LONG
   DECLARE FUNCTION SetSel (BYVAL fSelect AS BOOLEAN, BYVAL index AS LONG = -1) AS LONG
   DECLARE FUNCTION Select_ (BYVAL index AS LONG =-1) AS LONG
   DECLARE FUNCTION SelectString (BYVAL index AS LONG =-1) AS LONG
   DECLARE FUNCTION Unselect (BYVAL index AS LONG = -1) AS LONG
   DECLARE FUNCTION UnselectString (BYVAL index AS LONG = -1) AS LONG
   DECLARE SUB SetTabStops (BYVAL cTabs AS LONG, BYVAL prgTabs AS LONG PTR)
   DECLARE FUNCTION SetTopIndex (BYVAL index AS LONG) AS LONG

   DECLARE FUNCTION ReplaceString (BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
   DECLARE FUNCTION SetText (BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION CalcHorizontalExtent () AS LONG
   DECLARE FUNCTION AddItemData (BYVAL pData AS ANY PTR) AS LONG
   DECLARE FUNCTION InsertItemData (BYVAL index AS LONG, BYVAL pData AS ANY PTR) AS LONG
   DECLARE FUNCTION FindItemData (BYVAL indexStart AS LONG, BYVAL pData AS ANY PTR) AS LONG
   DECLARE FUNCTION SelectItemData (BYVAL indexStart AS LONG, BYVAL pData AS ANY PTR) AS LONG
   DECLARE FUNCTION GetSelect (BYVAL startIndex AS LONG = 0) AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' * CListBox constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CListBox (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListBox (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListBox (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListBox (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CListBox destructor
' ========================================================================================
PRIVATE DESTRUCTOR CListBox
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Enables the control.
' Return value: False if the control was previous enabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CListBox.Enable () AS BOOLEAN
   RETURN EnableWindow(m_hCtl, TRUE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Disables the control
' Return value: False if the control was previous disabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CListBox.Disable () AS BOOLEAN
   DIM res AS LONG = EnableWindow(m_hCtl, FALSE)
   DIM bEnable AS BOOLEAN
   IF res = 0 THEN bEnable = TRUE ELSE bEnable = FALSE
   RETURN bEnable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds the specified filename to a list box that contains a directory listing.
' ========================================================================================
PRIVATE FUNCTION CListBox.AddFile (BYVAL pwszFileName AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_ADDFILE, 0, CAST(LPARAM, pwszFileName))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a string to a list box. If the list box does not have the LBS_SORT style, the
' string is added to the end of the list. Otherwise, the string is inserted into the list
' and the list is sorted.
' ========================================================================================
PRIVATE FUNCTION CListBox.AddString (BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_ADDSTRING, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.Add (BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_ADDSTRING, 0, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a string in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.DeleteString (BYVAL Index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_DELETESTRING, Index, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.Delete_ (BYVAL Index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_DELETESTRING, Index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds names to the list displayed by a list box. The message adds the names of
' directories and files that match a specified string and set of file attributes.
' LB_DIR can also add mapped drive letters to the list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.Dir (BYVAL Attr AS DWORD, BYVAL pwszFileSpec AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_DIR, Attr, CAST(LPARAM, pwszFileSpec))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the first string in a list box that begins with the specified string.
' ========================================================================================
PRIVATE FUNCTION CListBox.FindString (BYVAL indexStart AS LONG, BYVAL pwszString AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_FINDSTRING, indexStart, CAST(LPARAM, pwszString))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.Find (BYVAL indexStart AS LONG, BYVAL pwszString AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_FINDSTRING, indexStart, CAST(LPARAM, pwszString))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the first list box string that exactly matches the specified string, except that
' the search is not case sensitive.
' ========================================================================================
PRIVATE FUNCTION CListBox.FindStringExact (BYVAL indexStart AS LONG, BYVAL pwszSearchString AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_FINDSTRINGEXACT, indexStart, CAST(LPARAM, pwszSearchString))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.FindExact (BYVAL indexStart AS LONG, BYVAL pwszSearchString AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_FINDSTRINGEXACT, indexStart, CAST(LPARAM, pwszSearchString))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the anchor item--that is, the item from which a multiple selection
' starts. A multiple selection spans all items from the anchor item to the caret item.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetAnchorIndex () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETANCHORINDEX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the index of the item that has the focus rectangle in a multiple-selection
' list box. The item may or may not be selected.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetCaretIndex () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETCARETINDEX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of items in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetCount () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the currently selected item, if any, in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetCurSel () AS LONG
   RETURN SendMessageW(m_hCtl, IIF((GetWindowLong(m_hCtl, GWL_STYLE) AND LBS_MULTIPLESEL), LB_GETCARETINDEX, LB_GETCURSEL), 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width, in pixels, that a list box can be scrolled horizontally (the scrollable
' width) if the list box has a horizontal scroll bar.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetHorizontalExtent () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETHORIZONTALEXTENT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of items per column in a specified list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetListBoxInfo () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETLISTBOXINFO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the application-defined value associated with the specified list box item.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetItemData (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETITEMDATA, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height of items in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetItemHeight (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETITEMHEIGHT, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the dimensions of the rectangle that bounds a list box item as it is currently
' displayed in the list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetItemRect (BYVAL index AS LONG, BYREF rc AS RECT) AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETITEMRECT, index, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.GetItemRect (BYVAL index AS LONG) AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, LB_GETITEMRECT, index, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current locale of the list box. You can use the locale to determine the correct
' sorting order of displayed text (for list boxes with the LBS_SORT style) and of text
' added by the LB_ADDSTRING message.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetLocale () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETLOCALE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the selection state of an item.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetSel (BYVAL Index AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, LB_GETSEL, Index, 0) > 0
END FUNCTION
' ========================================================================================
' ========================================================================================
' Checks if the item is selected or not.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetState (BYVAL Index AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, LB_GETSEL, Index, 0) > 0
END FUNCTION
' ========================================================================================
' ========================================================================================
' The listbox is searched to find the first selected item. If the item parameter is included,
' searching starts at that position to facilitate retrieving multiple selected items.
' If item is omitted, the search starts at the first data item. The index number of the
' selected item is returned as the result of the function. If no item is selected, the
' value -1 is returned.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetSelect (BYVAL startIndex AS LONG = 0) AS LONG
   IF (GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      ' // Search for the first selected item starting from startIndex (zero-based)
      DIM nCount AS LONG = SendMessageW(m_hCtl, LB_GETCOUNT, 0, 0)
      FOR selectedIndex AS LONG = startIndex TO nCount - 1
         IF SendMessageW(m_hCtl, LB_GETSEL, selectedIndex, 0) <> 0 THEN
            RETURN selectedIndex ' Zero-based index
         END IF
      NEXT
      ' // If no item is selected, return -1
      RETURN -1
   ELSE
      '// Single slection listbox
      RETURN this.GetCursel
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the total number of selected items in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetSelCount () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETSELCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fills a buffer with an array of integers that specify the item numbers of selected items
' in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetSelItems (BYVAL items AS DWORD, BYVAL prgItems AS LONG PTR) AS LONG
   RETURN SendMessage(m_hCtl, LB_GETSELITEMS, items, CAST(LPARAM, prgItems))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a string from a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetText (BYVAL Index AS LONG) AS DWSTRING
   DIM nLen AS LONG, buffer AS DWSTRING
   IF index < 0 THEN index = this.GetCurSel
   nLen = SendMessageW(m_hCtl, LB_GETTEXTLEN, index, 0)
   buffer = WSPACE(nLen + 1)
   SendMessageW m_hCtl, LB_GETTEXT, Index, CAST(LPARAM, STRPTR(buffer))
   RETURN LEFT(buffer, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the length of a string in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetTextLen (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETTEXTLEN, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the first visible item in a list box. Initially the item with index 0
' is at the top of the list box, but if the list box contents have been scrolled another
' item may be at the top.
' ========================================================================================
PRIVATE FUNCTION CListBox.GetTopIndex () AS LONG
   RETURN SendMessageW(m_hCtl, LB_GETTOPINDEX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allocates memory for storing list box items. This message is used before an application
' adds a large number of items to a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.InitStorage (BYVAL nItems AS DWORD, BYVAL nBytes AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, LB_INITSTORAGE, nItems, nBytes)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a string into a list box. Unlike the LB_ADDSTRING message, the LB_INSERTSTRING
' message does not cause a list with the LBS_SORT style to be sorted.
' ========================================================================================
PRIVATE FUNCTION CListBox.InsertString (BYVAL Index AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_INSERTSTRING, Index, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.Insert (BYVAL Index AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_INSERTSTRING, Index, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the zero-based index of the item nearest the specified point in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.ItemFromPoint (BYVAL x AS SHORT, BYVAL y AS SHORT) AS LONG
   RETURN SendMessageW(m_hCtl, LB_ITEMFROMPOINT, 0, MAKELONG(x, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all items from a list box.
' ========================================================================================
PRIVATE SUB CListBox.ResetContent
   SendMessageW m_hCtl, LB_RESETCONTENT, 0, 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CListBox.Reset
   SendMessageW m_hCtl, LB_RESETCONTENT, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Searches a list box for an item that begins with the characters in a specified string.
' If a matching item is found, the item is selected.
' ========================================================================================
PRIVATE FUNCTION CListBox.SelectString (BYVAL start AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SELECTSTRING, start, CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects one or more consecutive items in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.SelItemRange (BYVAL fSel AS LONG, BYVAL startIndex AS SHORT, BYVAL endIndex AS SHORT) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SELITEMRANGE, fSel, MAKELONG(startIndex, endIndex))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects one or more consecutive items in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.SelItemRangeEx (BYVAL startIndex AS LONG, BYVAL endIndex AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SELITEMRANGEEX, startIndex, endIndex)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the anchor item--that is, the item from which a multiple selection starts.
' A multiple selection spans all items from the anchor item to the caret item.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetAnchorIndex (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETANCHORINDEX, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus rectangle to the item at the specified index in a multiple-selection list
' box. If the item is not visible, it is scrolled into view.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetCaretIndex (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETCARETINDEX, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width, in pixels, of all columns in a multiple-column list box (created with
' the LBS_MULTICOLUMN style).
' ========================================================================================
PRIVATE FUNCTION CListBox.SetColumnWidth (BYVAL nWidth AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETCOLUMNWIDTH, nWidth, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the count of items in a list box created with the LBS_NODATA style and not created
' with the LBS_HASSTRINGS style.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetCount (BYVAL nCount AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETCOUNT, nCount, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a string and scroll it into view, if necessary. When the new string is selected,
' the list box removes the highlight from the previously selected string.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetCurSel (BYVAL Index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETCURSEL, Index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' An application sends an LB_SETHORIZONTALEXTENT message to set the width, in pixels, by
' which a list box can be scrolled horizontally (the scrollable width). If the width of
' the list box is smaller than this value, the horizontal scroll bar horizontally scrolls
' items in the list box. If the width of the list box is equal to or greater than this
' value, the horizontal scroll bar is hidden.
' ========================================================================================
PRIVATE SUB CListBox.SetHorizontalExtent (BYVAL cxExtent AS LONG)
   SendMessageW m_hCtl, LB_SETHORIZONTALEXTENT, cxExtent, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets a value associated with the specified item in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetItemData (BYVAL index AS LONG, BYVAL pData AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETITEMDATA, index, pData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the height, in pixels, of items in a list box. If the list box has the
' LBS_OWNERDRAWVARIABLE style, this message sets the height of the item specified by the
' nHeight parameter. Otherwise, this message sets the height of all items in the list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetItemHeight (BYVAL nHeight AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETITEMHEIGHT, nHeight, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current locale of the list box. You can use the locale to determine the correct
' sorting order of displayed text (for list boxes with the LBS_SORT style) and of text
' added by the LB_ADDSTRING message.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetLocale (BYVAL hListBox AS DWORD, BYVAL wLocaleID AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETLOCALE, wLocaleID, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a string in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetSel (BYVAL fSelect AS BOOLEAN, BYVAL index AS LONG = -1) AS LONG
   IF (GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      RETURN SendMessageW(m_hCtl, LB_SETSEL, fSelect, index)
   ELSE
      '// Single slection listbox
      RETURN SendMessageW(m_hCtl, LB_SETCURSEL, index, 0)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Selects the specified string in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.Select_ (BYVAL index AS LONG =-1) AS LONG
   IF (GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      RETURN SendMessageW(m_hCtl, LB_SETSEL, CTRUE, index)
   ELSE
      '// Single slection listbox
      RETURN SendMessageW(m_hCtl, LB_SETCURSEL, Index, 0)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Selects the specified string in a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.SelectString (BYVAL index AS LONG =-1) AS LONG
   IF (GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      RETURN SendMessageW(m_hCtl, LB_SETSEL, CTRUE, index)
   ELSE
      '// Single slection listbox
      RETURN SendMessageW(m_hCtl, LB_SETCURSEL, Index, 0)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Unselects the specified string in a multiple-selection list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.Unselect (BYVAL index AS LONG = -1) AS LONG
   IF (GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      RETURN SendMessageW(m_hCtl, LB_SETSEL, FALSE, index)
   ELSE
      ' // Single slection listbox
      RETURN SendMessageW(m_hCtl, LB_SETCURSEL, -1, 0)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.UnselectString (BYVAL index AS LONG = -1) AS LONG
   IF (GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL THEN
      ' // Multiple-selection listbox
      RETURN SendMessageW(m_hCtl, LB_SETSEL, FALSE, index)
   ELSE
      ' // Single slection listbox
      RETURN SendMessageW(m_hCtl, LB_SETCURSEL, -1, 0)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the tab-stop positions in a list box.
' ========================================================================================
PRIVATE SUB CListBox.SetTabStops (BYVAL cTabs AS LONG, BYVAL prgTabs AS LONG PTR)
   SendMessage m_hCtl, LB_SETTABSTOPS, cTabs, CAST(LPARAM, prgTabs)
END SUB
' ========================================================================================

' ========================================================================================
' Ensures that a particular item in a list box is visible.
' ========================================================================================
PRIVATE FUNCTION CListBox.SetTopIndex (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SETTOPINDEX, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces a string, and his associated data, into a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.ReplaceString (BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
   DIM lRes AS LONG = SendMessageW(m_hCtl, LB_DELETESTRING, index, 0)
   IF lRes = LB_ERR THEN RETURN lRes
   ' // Insert the new string
   index = SendMessageW(m_hCtl, LB_INSERTSTRING, index, CAST(LPARAM, pwszNewText))
   IF index = LB_ERR OR index = LB_ERRSPACE THEN RETURN index
   lRes = SendMessageW(m_hCtl, LB_SETITEMDATA, index, CAST(LPARAM, pNewData))
   IF lRes = LB_ERR THEN RETURN lRes
   RETURN SendMessageW(m_hCtl, IIF((GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_MULTIPLESEL), LB_SETSEL, LB_SETCURSEL), index, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListBox.SetText (BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR) AS LONG
   DIM lRes AS LONG = SendMessageW(m_hCtl, LB_DELETESTRING, index, 0)
   IF lRes = LB_ERR THEN RETURN lRes
   ' // Insert the new string
   index = SendMessageW(m_hCtl, LB_INSERTSTRING, index, CAST(LPARAM, pwszNewText))
   IF index = LB_ERR OR index = LB_ERRSPACE THEN RETURN index
   RETURN SendMessageW(m_hCtl, IIF((GetWindowLongPtrW(m_hCtl, GWL_STYLE) AND LBS_MULTIPLESEL), LB_SETSEL, LB_SETCURSEL), index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the size, in logical units, of the widest string item.
' ========================================================================================
PRIVATE FUNCTION CListBox.CalcHorizontalExtent () AS LONG
   ' // Get an handle to the device context of the control
   DIM hDC AS HDC = GetDC(m_hCtl)
   ' // Get the number of items
   DIM i AS LONG, cxScrollWidth AS LONG
   DIM cItems AS LRESULT = SendMessageW(m_hCtl, LB_GETCOUNT, 0, 0)
   FOR i = 0 TO cItems - 1
      ' // Get the text of the item
      DIM cbLen AS LONG = SendMessageW(m_hCtl, LB_GETTEXTLEN, i, 0)
      IF cbLen = 0 THEN CONTINUE FOR
      DIM pbuffer AS WSTRING PTR = CAllocate(cbLen * 2)
      SendMessageW(m_hCtl, LB_GETTEXT, i, CAST(LPARAM, pbuffer))
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
' Adds item data to the list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.AddItemData (BYVAL pData AS ANY PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_ADDSTRING, 0, CAST(LPARAM, pData))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts item data to a list box.
' ========================================================================================
PRIVATE FUNCTION CListBox.InsertItemData (BYVAL index AS LONG, BYVAL pData AS ANY PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_INSERTSTRING, index, CAST(LPARAM, pData))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Finds the first item in a list box that has the specified item data.
' ========================================================================================
PRIVATE FUNCTION CListBox.FindItemData (BYVAL indexStart AS LONG, BYVAL pData AS ANY PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_FINDSTRING, indexStart, CAST(LPARAM, pData))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Searches a list box for an item that has the specified item data. If a matching item is found, the item is selected.
' ========================================================================================
PRIVATE FUNCTION CListBox.SelectItemData (BYVAL indexStart AS LONG, BYVAL pData AS ANY PTR) AS LONG
   RETURN SendMessageW(m_hCtl, LB_SELECTSTRING, indexStart, CAST(LPARAM, pData))
END FUNCTION
' ========================================================================================

END NAMESPACE
