' ########################################################################################
' Platform: Microsoft Windows
' Filename: CListView.inc
' Purpose:  Listview control
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
#include once "Afx2/CDialog.inc"
#include once "Afx2/CControl.inc"
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CListView class
' ########################################################################################

TYPE CListView EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION ApproximateViewRect (BYVAL hCtl AS HWND, BYVAL iWidth AS SHORT, BYVAL iHeight AS SHORT, BYVAL iCount AS LONG) AS DWORD
   DECLARE FUNCTION Arrange (BYVAL hCtl AS HWND, BYVAL code AS UINT) AS BOOLEAN
   DECLARE SUB CancelEditLabel (BYVAL hCtl AS HWND)
   DECLARE FUNCTION CreateDragImage (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYREF lpptUpLeft AS POINT) AS HIMAGELIST
   DECLARE FUNCTION CreateDragImage (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL x AS LONG, BYVAL y AS LONG) AS HIMAGELIST
   DECLARE FUNCTION DeleteAllItems (BYVAL hCtl AS HWND) AS BOOLEAN
   DECLARE FUNCTION DeleteColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS BOOLEAN
   DECLARE FUNCTION DeleteItem (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS BOOLEAN
   DECLARE FUNCTION EditLabel (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS HWND
   DECLARE FUNCTION EnableGroupView (BYVAL hCtl AS HWND, BYVAL fEnable AS BOOLEAN) AS LONG
   DECLARE FUNCTION EnsureVisible (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL fPartialOK AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION FindItem (BYVAL hCtl AS HWND, BYVAL iStart AS LONG, BYREF vfi AS LVFINDINFOW) AS LONG
   DECLARE FUNCTION FindString (BYVAL hCtl AS HWND, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION GetBkColor (BYVAL hCtl AS HWND) AS COLORREF
   DECLARE FUNCTION GetBkImage (BYVAL hCtl AS HWND, BYREF vbki AS LVBKIMAGE) AS BOOLEAN
   DECLARE FUNCTION GetBkImage (BYVAL hCtl AS HWND) AS LVBKIMAGE
   DECLARE FUNCTION GetCallbackMask (BYVAL hCtl AS HWND) AS UINT
   DECLARE FUNCTION GetCheckState (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS UINT
   DECLARE FUNCTION GetColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYREF pcol AS LVCOLUMN) AS BOOLEAN
   DECLARE FUNCTION GetColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LVCOLUMN
   DECLARE FUNCTION GetColumnText (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL cchTextMax AS LONG = MAX_PATH + 1) AS DWSTRING
   DECLARE FUNCTION GetColumnOrderArray (BYVAL hCtl AS HWND, BYVAL iCount AS LONG, BYVAL lpiArray AS LONG PTR) AS LONG
   DECLARE FUNCTION GetColumnWidth (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LONG
   DECLARE FUNCTION GetCountPerPage (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetEditControl (BYVAL hCtl AS HWND) AS HWND
   DECLARE FUNCTION GetEmptyText (BYVAL hCtl AS HWND, BYREF wszText AS WSTRING, BYVAL cchText AS DWORD) AS BOOLEAN
   DECLARE FUNCTION GetExtendedListViewStyle (BYVAL hCtl AS HWND) AS DWORD
   DECLARE FUNCTION GetFocusedGroup (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetFooterInfo (BYVAL hCtl AS HWND, BYREF vfi AS LVFOOTERINFO) AS BOOLEaN
   DECLARE FUNCTION GetFooterInfo (BYVAL hCtl AS HWND) AS LVFOOTERINFO
   DECLARE FUNCTION GetFooterItem (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYREF fi AS LVFOOTERITEM) AS BOOLEAN
   DECLARE FUNCTION GetFooterItem (BYVAL hCtl AS HWND, BYVAL iItem AS UINT) AS LVFOOTERITEM
   DECLARE FUNCTION GetFooterItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS UINT, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetFooterItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS UINT) AS RECT
   DECLARE FUNCTION GetFooterRect (BYVAL hCtl AS HWND, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetFooterRect (BYVAL hCtl AS HWND) AS RECT
   DECLARE FUNCTION GetGroupCount (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetGroupHeaderImageList (BYVAL hCtl AS HWND) AS HIMAGELIST
   DECLARE FUNCTION GetGroupInfo (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYREF grp AS LVGROUP) AS LONG
   DECLARE FUNCTION GetGroupInfo (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG) AS LVGROUP
   DECLARE FUNCTION GetGroupInfoByIndex (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG, BYREF grp AS LVGROUP) AS BOOLEAN
   DECLARE FUNCTION GetGroupInfoByIndex (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG) AS LVGROUP
   DECLARE SUB GetGroupMetrics (BYVAL hCtl AS HWND, BYREF pGroupMetrics AS LVGROUPMETRICS)
   DECLARE FUNCTION GetGroupMetrics (BYVAL hCtl AS HWND) AS LVGROUPMETRICS
   DECLARE FUNCTION GetGroupRect (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYVAL nType AS LONG, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetGroupRect (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYVAL nType AS LONG) AS RECT
   DECLARE FUNCTION GetGroupState (BYVAL hCtl AS HWND, BYVAL dwGroupId AS UINT, BYVAL dwMask AS UINT) AS UINT
   DECLARE FUNCTION GetHeader (BYVAL hCtl AS HWND) AS HWND
   DECLARE FUNCTION GetHotCursor (BYVAL hCtl AS HWND) AS HCURSOR
   DECLARE FUNCTION GetHotItem (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetHoverTime (BYVAL hCtl AS HWND) AS DWORD
   DECLARE FUNCTION GetImageList (BYVAL hCtl AS HWND, BYVAL iImageList AS LONG) AS HIMAGELIST
   DECLARE FUNCTION GetInsertMark (BYVAL hCtl AS HWND, BYREF vim AS LVINSERTMARK) AS BOOLEAN
   DECLARE FUNCTION GetInsertMark (BYVAL hCtl AS HWND) AS LVINSERTMARK
   DECLARE FUNCTION GetInsertMarkColor (BYVAL hCtl AS HWND) AS COLORREF
   DECLARE FUNCTION GetInsertMarkRect (BYVAL hCtl AS HWND, BYREF rc AS RECT) AS LONG
   DECLARE FUNCTION GetInsertMarkRect (BYVAL hCtl AS HWND) AS RECT
   DECLARE FUNCTION GetISearchStringLen (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetISearchString (BYVAL hCtl AS HWND) AS DWSTRING
   DECLARE FUNCTION GetItem (BYVAL hCtl AS HWND, BYREF item AS LVITEM) AS BOOLEAN
   DECLARE FUNCTION GetItemCount (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetItemIndexRect (BYVAL hCtl AS HWND, BYREF vii AS LVITEMINDEX, BYVAL iSubItem AS LONG, BYVAL code AS LONG, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetItemIndexRect (BYVAL hCtl AS HWND, BYREF vii AS LVITEMINDEX, BYVAL iSubItem AS LONG, BYVAL code AS LONG) AS RECT
   DECLARE FUNCTION GetItemPosition (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYREF pt AS POINT) AS BOOLEAN
   DECLARE FUNCTION GetItemPosition (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS POINT
   DECLARE FUNCTION GetItemRect (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYREF rc AS RECT, BYVAL code AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetItemRect (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL code AS LONG) AS RECT
   DECLARE FUNCTION GetItemSpacing (BYVAL hCtl AS HWND, BYVAL fSmall AS LONG) AS DWORD
   DECLARE FUNCTION GetItemState (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL mask AS LONG) AS UINT
   DECLARE SUB GetItemText (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL iSubItem AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG = MAX_PATH + 1)
   DECLARE FUNCTION GetTextW (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL iSubItem AS LONG, BYVAL cchTextMax AS LONG = MAX_PATH + 1) AS DWSTRING
   DECLARE FUNCTION GetNextItem (BYVAL hCtl AS HWND, BYVAL iStart AS LONG, BYVAL flags AS UINT) AS LONG
   DECLARE FUNCTION GetSelection (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetFocusedItem (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetNextItemIndex (BYVAL hCtl AS HWND, BYREF plvii AS LVITEMINDEX, BYVAL flags AS LONG) AS LONG
   DECLARE FUNCTION GetNumberOfWorkAreas (BYVAL hCtl AS HWND) AS DWORD
   DECLARE FUNCTION GetOrigin (BYVAL hCtl AS HWND, BYREF ptOrg AS POINT) AS BOOLEAN
   DECLARE FUNCTION GetOrigin (BYVAL hCtl AS HWND) AS POINT
   DECLARE FUNCTION GetOutlineColor (BYVAL hCtl AS HWND) AS COLORREF
   DECLARE FUNCTION GetSelectedColumn (BYVAL hCtl AS HWND) AS UINT
   DECLARE FUNCTION GetSelectedCount (BYVAL hCtl AS HWND) AS UINT
   DECLARE FUNCTION GetSelectionMark (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetStringWidth (BYVAL hCtl AS HWND, BYREF wszText AS WSTRING) AS LONG
   DECLARE FUNCTION GetSubItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYVAL iSubItem AS LONG, BYVAL nCode AS LONG, BYREF rc AS RECT) AS LONG
   DECLARE FUNCTION GetSubItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYVAL iSubItem AS LONG, BYVAL nCode AS LONG) AS RECT
   DECLARE FUNCTION GetTextBkColor (BYVAL hCtl AS HWND) AS COLORREF
   DECLARE FUNCTION GetTextColor (BYVAL hCtl AS HWND) AS COLORREF
   DECLARE SUB GetTileInfo (BYVAL hCtl AS HWND, BYREF ti AS LVTILEINFO)
   DECLARE FUNCTION GetTileInfo (BYVAL hCtl AS HWND) AS LVTILEINFO
   DECLARE SUB GetTileViewInfo (BYVAL hCtl AS HWND, BYREF tvi AS LVTILEVIEWINFO)
   DECLARE FUNCTION GetTileViewInfo (BYVAL hCtl AS HWND) AS LVTILEVIEWINFO
   DECLARE FUNCTION GetToolTips (BYVAL hCtl AS HWND) AS HWND
   DECLARE FUNCTION GetTopIndex (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetUnicodeFormat (BYVAL hCtl AS HWND) AS BOOLEAN
   DECLARE FUNCTION GetView (BYVAL hCtl AS HWND) AS DWORD
   DECLARE FUNCTION GetViewRect (BYVAL hCtl AS HWND, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetViewRect (BYVAL hCtl AS HWND) AS RECT
   DECLARE SUB GetWorkAreas (BYVAL hCtl AS HWND, BYVAL nWorkAreas AS LONG, BYREF rc AS RECT)
   DECLARE FUNCTION GetWorkAreas (BYVAL hCtl AS HWND, BYVAL nWorkAreas AS LONG) AS RECT
   DECLARE FUNCTION HasGroup (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG) AS BOOLEAN
   DECLARE FUNCTION HitTest (BYVAL hCtl AS HWND, BYREF info AS LVHITTESTINFO) AS LONG
   DECLARE FUNCTION HitTest (BYVAL hCtl AS HWND) AS LVHITTESTINFO
   DECLARE FUNCTION HitTestEx (BYVAL hCtl AS HWND, BYREF info AS LVHITTESTINFO) AS LONG
   DECLARE FUNCTION HitTestEx (BYVAL hCtl AS HWND) AS LVHITTESTINFO
   DECLARE FUNCTION InsertColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYREF col AS LVCOLUMN) AS LONG
   DECLARE FUNCTION InsertColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LVCOLUMN
   DECLARE FUNCTION AddColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL nWidth AS LONG, BYVAL nFormat AS LONG = LVCFMT_LEFT) AS LONG
   DECLARE FUNCTION InsertGroup (BYVAL hCtl AS HWND, BYVAL index AS LONG, BYREF grp AS LVGROUP) AS LONG
   DECLARE SUB InsertGroupSorted (BYVAL hCtl AS HWND, BYREF structInsert AS LVINSERTGROUPSORTED)
   DECLARE FUNCTION InsertItem (BYVAL hCtl AS HWND, BYREF item AS LVITEM) AS LONG
   DECLARE FUNCTION AddItem (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION AppendItem (BYVAL hCtl AS HWND, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DECLARE FUNCTION InsertMarkHitTest (BYVAL hCtl AS HWND, BYREF pt AS POINT, BYREF vim AS LVINSERTMARK) AS BOOLEAN
   DECLARE FUNCTION IsGroupViewEnabled (BYVAL hCtl AS HWND) AS BOOLEAN
   DECLARE FUNCTION IsItemVisible (BYVAL hCtl AS HWND, BYVAL index AS DWORD) AS BOOLEAN
   DECLARE FUNCTION MapIDToIndex (BYVAL hCtl AS HWND, BYVAL id AS UINT) AS UINT
   DECLARE FUNCTION MapIndexToID (BYVAL hCtl AS HWND, BYVAL index AS UINT) AS UINT
   DECLARE FUNCTION RedrawItems (BYVAL hCtl AS HWND, BYVAL iFirst AS LONG, BYVAL iLast AS LONG) AS BOOLEAN
   DECLARE SUB RemoveAllGroups (BYVAL hCtl AS HWND)
   DECLARE FUNCTION RemoveGroup (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG) AS LONG
   DECLARE FUNCTION Scroll (BYVAL hCtl AS HWND, BYVAL dx AS LONG, BYVAL dy AS LONG) AS BOOLEAN
   DECLARE FUNCTION SetBkColor (BYVAL hCtl AS HWND, BYVAL clrBk AS COLORREF) AS BOOLEAN
   DECLARE FUNCTION SetBkImage (BYVAL hCtl AS HWND, BYREF vbki AS LVBKIMAGE) AS BOOLEAN
   DECLARE FUNCTION SetCallbackMask (BYVAL hCtl AS HWND, BYVAL mask AS DWORD) AS BOOLEAN
   DECLARE SUB SetCheckState (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG, BYVAL fCheck AS BOOLEAN)
   DECLARE FUNCTION SetColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYREF col AS LVCOLUMN) AS BOOLEAN
   DECLARE FUNCTION SetColumnOrderArray (BYVAL hCtl AS HWND, BYVAL iCount AS LONG, BYVAL pi AS LONG PTR) AS BOOLEAN
   DECLARE FUNCTION SetColumnWidth (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL cx AS LONG) AS BOOLEAN
   DECLARE SUB SetExtendedListViewStyle (BYVAL hCtl AS HWND, BYVAL dwExStyle AS DWORD)
   DECLARE SUB SetExtendedListViewStyleEx (BYVAL hCtl AS HWND, BYVAL dwExMask AS DWORD, BYVAL dwExStyle AS DWORD)
   DECLARE FUNCTION SetGroupHeaderImageList (BYVAL hCtl AS HWND, BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   DECLARE FUNCTION SetGroupInfo (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYREF grp AS LVGROUP) AS LONG
   DECLARE SUB SetGroupMetrics (BYVAL hCtl AS HWND, BYREF GroupMetrics AS LVGROUPMETRICS)
   DECLARE FUNCTION SetGroupState (BYVAL hCtl AS HWND, BYVAL dwGroupId AS UINT, BYVAL dwMask AS UINT, BYVAL dwState AS UINT) AS LONG
   DECLARE FUNCTION SetHotCursor (BYVAL hCtl AS HWND, BYVAL hCur AS HCURSOR) AS HCURSOR
   DECLARE FUNCTION SetHotItem (BYVAL hCtl AS HWND, BYVAL iIndex AS DWORD) AS LONG
   DECLARE SUB SetHoverTime (BYVAL hCtl AS HWND, BYVAL dwHoverTimeMs AS DWORD)
   DECLARE FUNCTION SetIconSpacing (BYVAL hCtl AS HWND, BYVAL cx AS SHORT, BYVAL cy AS SHORT) AS DWORD
   DECLARE FUNCTION SetImageList (BYVAL hCtl AS HWND, BYVAL himl AS HIMAGELIST, BYVAL nImageList AS LONG) AS HIMAGELIST
   DECLARE FUNCTION SetInfoTip (BYVAL hCtl AS HWND, BYREF lvInfoTip AS LVSETINFOTIP) AS LONG
   DECLARE FUNCTION SetInsertMark (BYVAL hCtl AS HWND, BYREF lvim AS LVINSERTMARK) AS LONG
   DECLARE FUNCTION SetInsertMarkColor (BYVAL hCtl AS HWND, BYVAL colr AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetItem (BYVAL hCtl AS HWND, BYREF item AS LVITEM) AS BOOLEAN
   DECLARE SUB SetItemCount (BYVAL hCtl AS HWND, BYVAL cItems AS LONG)
   DECLARE SUB SetItemCountEx (BYVAL hCtl AS HWND, BYVAL cItems AS LONG, BYVAL dwFlags AS LONG)
   DECLARE FUNCTION SetItemIndexState (BYVAL hCtl AS HWND, BYREF lvii AS LVITEMINDEX, BYVAL dwData AS UINT, BYVAL mask AS UINT) AS HRESULT
   DECLARE FUNCTION SetItemPosition (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL x AS SHORT, BYVAL y AS SHORT) AS BOOLEAN
   DECLARE SUB SetItemPosition32 (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL x AS LONG, BYVAL y AS LONG)
   DECLARE SUB SetItemState (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL dwState AS DWORD, BYVAL mask AS DWORD)
   DECLARE SUB SetItemText (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL iSubItem AS LONG, BYREF wszText AS WSTRING)
   DECLARE FUNCTION SetOutlineColor (BYVAL hCtl AS HWND, BYVAL colr AS COLORREF) AS COLORREF
   DECLARE SUB SetSelectedColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG)
   DECLARE FUNCTION SetSelectionMark (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG) AS LONG
   DECLARE FUNCTION SetTextBkColor (BYVAL hCtl AS HWND, BYVAL clrTextBk AS COLORREF) AS BOOLEAN
   DECLARE FUNCTION SetTextColor (BYVAL hCtl AS HWND, BYVAL clrText AS COLORREF) AS BOOLEAN
   DECLARE FUNCTION SetTileInfo (BYVAL hCtl AS HWND, BYREF lvtinfo AS LVTILEINFO) AS BOOLEAN
   DECLARE FUNCTION SetTileViewInfo (BYVAL hCtl AS HWND, BYREF lvtvinfo AS LVTILEVIEWINFO) AS BOOLEAN
   DECLARE FUNCTION SetToolTips (BYVAL hCtl AS HWND, BYVAL hwndToolTip AS HWND) AS HWND
   DECLARE FUNCTION SetUnicodeFormat (BYVAL hCtl AS HWND, BYVAL fUnicode AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION SetView (BYVAL hCtl AS HWND, BYVAL iView AS DWORD) AS LONG
   DECLARE SUB SetWorkAreas (BYVAL hCtl AS HWND, BYVAL nWorkAreas AS LONG, BYREF rc AS RECT)
   DECLARE FUNCTION SortGroups (BYVAL hCtl AS HWND, BYVAL _pfnGroupCompare AS PFNLVGROUPCOMPARE PTR, BYVAL _plv AS ANY PTR) AS BOOLEAN
   DECLARE FUNCTION SortItems (BYVAL hCtl AS HWND, BYVAL _pfnCompare AS PFNLVCOMPARE PTR, BYVAL _lPrm AS LONG) AS BOOLEAN
   DECLARE FUNCTION SortItemsEx (BYVAL hCtl AS HWND, BYVAL _pfnCompare AS PFNLVCOMPARE PTR, BYVAL _lPrm AS LONG) AS BOOLEAN
   DECLARE FUNCTION SubItemHitTest (BYVAL hCtl AS HWND, BYREF lvhti AS LVHITTESTINFO) AS LONG
   DECLARE FUNCTION SubItemHitTestEx (BYVAL hCtl AS HWND, BYREF lvhti AS LVHITTESTINFO) AS LONG
   DECLARE FUNCTION Update (BYVAL hCtl AS HWND, BYVAL iItem AS LONG) AS BOOLEAN
   ' // Additional procedures
   DECLARE SUB SelectItem (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG)
   DECLARE SUB SelectAllItems (BYVAL hCtl AS HWND)
   DECLARE SUB UnselectItem (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG)
   DECLARE SUB UnselectAllItems (BYVAL hCtl AS HWND)
   DECLARE SUB MakeHeaderFlat (BYVAL hCtl AS HWND)
   DECLARE FUNCTION FitContent (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS BOOLEAN
   DECLARE FUNCTION AutoSizeColumns (BYVAL hCtl AS HWND) AS BOOLEAN
   DECLARE FUNCTION FitHeader (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS BOOLEAN
   DECLARE FUNCTION AutoSizeHeader (BYVAL hCtl AS HWND) AS BOOLEAN
   DECLARE FUNCTION GetColumnCount (BYVAL hCtl AS HWND) AS LONG
   DECLARE FUNCTION GetHeaderText (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL cchTextMax AS LONG = MAX_PATH + 1) AS DWSTRING
   DECLARE FUNCTION GetColumnOrder (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LONG
   DECLARE FUNCTION GetOrderArray (BYVAL hCtl AS HWND, BYVAL iSize AS DWORD, BYVAL pi AS LONG PTR) AS LONG
   DECLARE FUNCTION GetTooltipsFont (BYVAL hCtl AS HWND) AS HFONT
   DECLARE SUB SetTooltipsFont (BYVAL hCtl AS HWND, BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE FUNCTION SetHeaderFont (BYVAL hCtl AS HWND, BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, _
      BYVAL lWeight AS LONG, BYVAL bItalic AS BYTE, BYVAL bUnderline AS BYTE, BYVAL bStrikeOut AS BYTE, _
      BYVAL bCharSet AS BYTE) AS HFONT
   DECLARE FUNCTION ModifyHeaderFont (BYVAL hCtl AS HWND, BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, _
      BYVAL lWeight AS LONG, BYVAL bItalic AS BYTE, BYVAL bUnderline AS BYTE, BYVAL bStrikeOut AS BYTE, _
      BYVAL bCharSet AS BYTE) AS HFONT

END TYPE
' ========================================================================================

' ========================================================================================
' * CListView constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CListView (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListView (BYVAL hParent AS HWND, BYVAL hCtl AS HWND)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListView (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListView (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CListView (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CListView destructor
' ========================================================================================
PRIVATE DESTRUCTOR CListView
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Calculates the approximate width and height required to display a given number of items.
' ========================================================================================
PRIVATE FUNCTION CListView.ApproximateViewRect (BYVAL hCtl AS HWND, BYVAL iWidth AS SHORT, BYVAL iHeight AS SHORT, BYVAL iCount AS LONG) AS DWORD
   RETURN SendMessageW(hCtl, LVM_APPROXIMATEVIEWRECT, iCount, MAKELONG(iWidth, iHeight))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Arranges items in icon view.
' ========================================================================================
PRIVATE FUNCTION CListView.Arrange (BYVAL hCtl AS HWND, BYVAL code AS UINT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_ARRANGE, code, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Cancels an item text editing operation.
' ========================================================================================
PRIVATE SUB CListView.CancelEditLabel (BYVAL hCtl AS HWND)
   SendMessageW hCtl, LVM_CANCELEDITLABEL, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Creates a drag image list for the specified item.
' ========================================================================================
PRIVATE FUNCTION CListView.CreateDragImage (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYREF lpptUpLeft AS POINT) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(hCtl, LVM_CREATEDRAGIMAGE, i, CAST(LPARAM, @lpptUpLeft)))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.CreateDragImage (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL x AS LONG, BYVAL y AS LONG) AS HIMAGELIST
   DIM tUpLeft AS POINT = TYPE<POINT>(x, y)
   RETURN CAST(HIMAGELIST, SendMessageW(hCtl, LVM_CREATEDRAGIMAGE, i, CAST(LPARAM, @tUpLeft)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all items from a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.DeleteAllItems (BYVAL hCtl AS HWND) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_DELETEALLITEMS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes a column from a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.DeleteColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_DELETECOLUMN, iCol, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an item from a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.DeleteItem (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_DELETEITEM, i, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Begins in-place editing of the specified list-view item's text. The message implicitly
' selects and focuses the specified item.
' ========================================================================================
PRIVATE FUNCTION CListView.EditLabel (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS HWND
   RETURN CAST(HWND, SendMessageW(hCtl, LVM_EDITLABEL, i, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables or disables whether the items in a list-view control display as a group.
' ========================================================================================
PRIVATE FUNCTION CListView.EnableGroupView (BYVAL hCtl AS HWND, BYVAL fEnable AS BOOLEAN) AS LONG
   RETURN SendMessageW(hCtl, LVM_ENABLEGROUPVIEW, fEnable, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Ensures that a list-view item is either entirely or partially visible, scrolling the
' list-view control if necessary.
' ========================================================================================
PRIVATE FUNCTION CListView.EnsureVisible (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL fPartialOK AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_ENSUREVISIBLE, i, MAKELONG(fPartialOK, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for a list-view item with the specified characteristics.
' ========================================================================================
PRIVATE FUNCTION CListView.FindItem (BYVAL hCtl AS HWND, BYVAL iStart AS LONG, BYREF vfi AS LVFINDINFOW) AS LONG
   RETURN SendMessageW(hCtl, LVM_FINDITEMW, iStart, CAST(LPARAM, @vfi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for a list-view item with the specified text.
' Returns the index of the item if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CListView.FindString (BYVAL hCtl AS HWND, BYVAL pwszText AS WSTRING PTR) AS LONG
   DIM lvf AS LVFINDINFOW
   lvf.flags = LVFI_STRING
   lvf.psz = pwszText
   RETURN SendMessageW(hCtl, LVM_FINDITEMW, -1, CAST(LPARAM, @lvf))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the background color of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetBkColor (BYVAL hCtl AS HWND) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the background image in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetBkImage (BYVAL hCtl AS HWND, BYREF vbki AS LVBKIMAGE) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETBKIMAGE, 0, CAST(LPARAM, @vbki))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetBkImage (BYVAL hCtl AS HWND) AS LVBKIMAGE
   DIM vbki AS LVBKIMAGE
   SendMessageW(hCtl, LVM_GETBKIMAGE, 0, CAST(LPARAM, @vbki))
   RETURN vbki
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the callback mask for a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetCallbackMask (BYVAL hCtl AS HWND) AS UINT
   RETURN SendMessageW(hCtl, LVM_GETCALLBACKMASK, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if an item in a list-view control is selected. This should be used only for
' list-view controls that have the LVS_EX_CHECKBOXES style.
' ========================================================================================
PRIVATE FUNCTION CListView.GetCheckState (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS UINT
   DIM nCheck AS LONG
   nCheck = SendMessageW(hCtl, LVM_GETITEMSTATE, i, LVIS_STATEIMAGEMASK)
   RETURN BIT(nCheck, 12) - 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the attributes of a list-view control's column.
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYREF pcol AS LVCOLUMN) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETCOLUMN, iCol, CAST(LPARAM, @pcol))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LVCOLUMN
   DIM pcol AS LVCOLUMN
   SendMessageW(hCtl, LVM_GETCOLUMN, iCol, CAST(LPARAM, @pcol))
   RETURN pcol
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of a list-view control's column.
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumnText (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL cchTextMax AS LONG = MAX_PATH + 1 ) AS DWSTRING
   DIM dwsText AS DWSTRING
   IF cchTextMax < 1 THEN cchTextMax = MAX_PATH + 1
   dwsText = WSPACE(cchTextMax + 1)
   DIM lvc AS LVCOLUMNW
   lvc.mask = LVCF_TEXT
   lvc.cchTextMax = cchTextMax + 1
   lvc.pszText = STRPTR(dwsText)
   IF SendMessageW(hCtl, LVM_GETCOLUMNW, iCol, CAST(LPARAM, @lvc)) THEN RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current left-to-right order of columns in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumnOrderArray (BYVAL hCtl AS HWND, BYVAL iCount AS LONG, BYVAL lpiArray AS LONG PTR) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETCOLUMNORDERARRAY, iCount, CAST(LPARAM, lpiArray))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of a column in report or list view.
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumnWidth (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETCOLUMNWIDTH, iCol, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the number of items that can fit vertically in the visible area of a
' list-view control when in list or report view.
' ========================================================================================
PRIVATE FUNCTION CListView.GetCountPerPage (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETCOUNTPERPAGE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle to the edit control being used to edit a list-view item's text.
' ========================================================================================
PRIVATE FUNCTION CListView.GetEditControl (BYVAL hCtl AS HWND) AS HWND
   RETURN CAST(HWND, SendMessageW(hCtl, LVM_GETEDITCONTROL, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text meant for display when the list-view control appears empty.
' ========================================================================================
PRIVATE FUNCTION CListView.GetEmptyText (BYVAL hCtl AS HWND, BYREF wszText AS WSTRING, BYVAL cchText AS DWORD) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETEMPTYTEXT, cchText, CAST(LPARAM, @wszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the extended styles that are currently in use for a given list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetExtendedListViewStyle (BYVAL hCtl AS HWND) AS DWORD
   RETURN SendMessageW(hCtl, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the group that has the focus.
' ========================================================================================
PRIVATE FUNCTION CListView.GetFocusedGroup (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETFOCUSEDGROUP, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information on the footer of a specified list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterInfo (BYVAL hCtl AS HWND, BYREF vfi AS LVFOOTERINFO) AS BOOLEaN
   RETURN SendMessageW(hCtl, LVM_GETFOOTERINFO, 0, CAST(LPARAM, @vfi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterInfo (BYVAL hCtl AS HWND) AS LVFOOTERINFO
   DIM vfi AS LVFOOTERINFO
   SendMessageW(hCtl, LVM_GETFOOTERINFO, 0, CAST(LPARAM, @vfi))
   RETURN vfi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information on a footer item for a specified list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterItem (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYREF fi AS LVFOOTERITEM) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETFOOTERITEM, iItem, CAST(LPARAM, @fi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterItem (BYVAL hCtl AS HWND, BYVAL iItem AS UINT) AS LVFOOTERITEM
   DIM fi AS LVFOOTERITEM
   SendMessageW(hCtl, LVM_GETFOOTERITEM, iItem, CAST(LPARAM, @fi))
   RETURN fi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information on a footer item for a specified list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS UINT, BYREF rc AS RECT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETFOOTERITEMRECT, iItem, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS UINT) AS RECT
   DIM rc AS RECT
   SendMessageW(hCtl, LVM_GETFOOTERITEMRECT, iItem, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the coordinates of the footer for a specified list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterRect (BYVAL hCtl AS HWND, BYREF rc AS RECT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETFOOTERRECT, 0, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetFooterRect (BYVAL hCtl AS HWND) AS RECT
   DIM rc AS RECT
   SendMessageW(hCtl, LVM_GETFOOTERRECT, 0, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of groups.
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupCount (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETGROUPCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the group header image list that has been set for an existing list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupHeaderImageList (BYVAL hCtl AS HWND) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(hCtl, LVM_GETIMAGELIST, LVSIL_GROUPHEADER, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets group information.
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupInfo (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYREF grp AS LVGROUP) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETGROUPINFO, iGroupId, CAST(LPARAM, @grp))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupInfo (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG) AS LVGROUP
   DIM grp AS LVGROUP
   SendMessageW(hCtl, LVM_GETGROUPINFO, iGroupId, CAST(LPARAM, @grp))
   RETURN grp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information on a specified group.
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupInfoByIndex (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG, BYREF grp AS LVGROUP) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETGROUPINFOBYINDEX, iIndex, CAST(LPARAM, @grp))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupInfoByIndex (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG) AS LVGROUP
   DIM grp AS LVGROUP
   SendMessageW(hCtl, LVM_GETGROUPINFOBYINDEX, iIndex, CAST(LPARAM, @grp))
   RETURN grp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about the display of groups.
' ========================================================================================
PRIVATE SUB CListView.GetGroupMetrics (BYVAL hCtl AS HWND, BYREF GroupMetrics AS LVGROUPMETRICS)
   SendMessageW hCtl, LVM_GETGROUPMETRICS, 0, CAST(LPARAM, @GroupMetrics)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupMetrics (BYVAL hCtl AS HWND) AS LVGROUPMETRICS
   DIM GroupMetrics AS LVGROUPMETRICS
   SendMessageW hCtl, LVM_GETGROUPMETRICS, 0, CAST(LPARAM, @GroupMetrics)
   RETURN GroupMetrics
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the rectangle for a specified group.
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupRect (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYVAL nType AS LONG, BYREF rc AS RECT) AS BOOLEAN
   rc.Top = nType
   RETURN SendMessageW(hCtl, LVM_GETGROUPRECT, iGroupId, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupRect (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYVAL nType AS LONG) AS RECT
   DIM rc AS RECT
   rc.Top = nType
   SendMessageW(hCtl, LVM_GETGROUPRECT, iGroupId, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the state for a specified group.
' ========================================================================================
PRIVATE FUNCTION CListView.GetGroupState (BYVAL hCtl AS HWND, BYVAL dwGroupId AS UINT, BYVAL dwMask AS UINT) AS UINT
   RETURN SendMessageW(hCtl, LVM_GETGROUPSTATE, dwGroupId, dwMask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle to the header control used by a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetHeader (BYVAL hCtl AS HWND) AS HWND
   RETURN CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the HCURSOR used when the pointer is over an item while hot tracking is enabled.
' ========================================================================================
PRIVATE FUNCTION CListView.GetHotCursor (BYVAL hCtl AS HWND) AS HCURSOR
   RETURN CAST(HCURSOR, SendMessageW(hCtl, LVM_GETHOTCURSOR, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the hot item.
' ========================================================================================
PRIVATE FUNCTION CListView.GetHotItem (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETHOTITEM, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the amount of time that the mouse cursor must hover over an item before it is selected.
' ========================================================================================
PRIVATE FUNCTION CListView.GetHoverTime (BYVAL hCtl AS HWND) AS DWORD
   RETURN SendMessageW(hCtl, LVM_GETHOVERTIME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle to an image list used for drawing list-view items.
' ========================================================================================
PRIVATE FUNCTION CListView.GetImageList (BYVAL hCtl AS HWND, BYVAL iImageList AS LONG) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(hCtl, LVM_GETIMAGELIST, iImageList, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the position of the insertion point.
' ========================================================================================
PRIVATE FUNCTION CListView.GetInsertMark (BYVAL hCtl AS HWND, BYREF vim AS LVINSERTMARK) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETINSERTMARK, 0, CAST(LPARAM, @vim))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetInsertMark (BYVAL hCtl AS HWND) AS LVINSERTMARK
   DIM vim AS LVINSERTMARK
   SendMessageW(hCtl, LVM_GETINSERTMARK, 0, CAST(LPARAM, @vim))
   RETURN vim
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the color of the insertion point.
' ========================================================================================
PRIVATE FUNCTION CListView.GetInsertMarkColor (BYVAL hCtl AS HWND) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_GETINSERTMARKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the rectangle that bounds the insertion point.
' ========================================================================================
PRIVATE FUNCTION CListView.GetInsertMarkRect (BYVAL hCtl AS HWND, BYREF rc AS RECT) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETINSERTMARKRECT, 0, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetInsertMarkRect (BYVAL hCtl AS HWND) AS RECT
   DIM rc AS RECT
   SendMessageW(hCtl, LVM_GETINSERTMARKRECT, 0, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the length of the incremental search string of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetISearchStringLen (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETISEARCHSTRINGW, 0, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the incremental search string of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetISearchString (BYVAL hCtl AS HWND) AS DWSTRING
   DIM nLen AS LONG = SendMessageW(hCtl, LVM_GETISEARCHSTRINGW, 0, BYVAL NULL)
   DIM dwsSearch AS DWSTRING =WSPACE(nLen + 1)
   nLen = SendMessageW(hCtl, LVM_GETISEARCHSTRINGW, 0, BYVAL CAST(LPARAM, STRPTR(dwsSearch)))
   RETURN LEFT(dwsSearch, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets some or all of a list-view item's attributes.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItem (BYVAL hCtl AS HWND, BYREF item AS LVITEM) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETITEM, 0, CAST(LPARAM, @item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of items in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemCount (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETITEMCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the bounding rectangle for all or part of a subitem in the current view of a
' specified list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemIndexRect (BYVAL hCtl AS HWND, BYREF vii AS LVITEMINDEX, BYVAL iSubItem AS LONG, BYVAL code AS LONG, BYREF rc AS RECT) AS BOOLEAN
   rc.Top = iSubItem : rc.Left = Code
   RETURN SendMessageW(hCtl, LVM_GETITEMINDEXRECT, CAST(WPARAM, @vii), CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemIndexRect (BYVAL hCtl AS HWND, BYREF vii AS LVITEMINDEX, BYVAL iSubItem AS LONG, BYVAL code AS LONG) AS RECT
   dim rc AS RECT: rc.Top = iSubItem : rc.Left = Code
   SendMessageW(hCtl, LVM_GETITEMINDEXRECT, CAST(WPARAM, @vii), CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the position of a list-view item.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemPosition (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYREF pt AS POINT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETITEMPOSITION, i, CAST(LPARAM, @pt))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemPosition (BYVAL hCtl AS HWND, BYVAL i AS LONG) AS POINT
   DIM pt AS POINT
   SendMessageW(hCtl, LVM_GETITEMPOSITION, i, CAST(LPARAM, @pt))
   RETURN pt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the bounding rectangle for all or part of an item in the current view.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemRect (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYREF rc AS RECT, BYVAL code AS LONG) AS BOOLEAN
   rc.Left = code
   RETURN SendMessageW(hCtl, LVM_GETITEMRECT, i, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemRect (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL code AS LONG) AS RECT
   DIM rc AS RECT : rc.Left = code
   SendMessageW(hCtl, LVM_GETITEMRECT, i, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the spacing between items in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemSpacing (BYVAL hCtl AS HWND, BYVAL fSmall AS LONG) AS DWORD
   RETURN SendMessageW(hCtl, LVM_GETITEMSPACING, fSmall, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the state of a list-view item.
' ========================================================================================
PRIVATE FUNCTION CListView.GetItemState (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL mask AS LONG) AS UINT
   RETURN SendMessageW(hCtl, LVM_GETITEMSTATE, i, mask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of a list-view item or subitem.
' ========================================================================================
PRIVATE SUB CListView.GetItemText (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL iSubItem AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG  = MAX_PATH + 1)
   DIM lvi AS LVITEMW
   lvi.iSubItem = iSubItem
   lvi.cchTextMax = cchTextMax
   lvi.pszText = pwszText
   SendMessageW hCtl, LVM_GETITEMTEXTW, i, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetTextW (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL iSubItem AS LONG, BYVAL cchTextMax AS LONG = MAX_PATH + 1) AS DWSTRING
   DIM dwsText AS DWSTRING
   DIM lvi AS LVITEMW
   DIM nLen AS LONG
   IF cchTextMax < 1 THEN cchTextMax = MAX_PATH + 1
   dwsText = WSPACE(cchTextMax + 1)
   lvi.iSubItem = iSubItem
   lvi.cchTextMax = cchTextMax + 1
   lvi.pszText = STRPTR(dwsText)
   nLen = SendMessageW(hCtl, LVM_GETITEMTEXTW, i, CAST(LPARAM, @lvi))
   RETURN LEFT(dwsText, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for a list-view item that has the specified properties and bears the specified
' relationship to a specified item.
' ========================================================================================
PRIVATE FUNCTION CListView.GetNextItem (BYVAL hCtl AS HWND, BYVAL iStart AS LONG, BYVAL flags AS UINT) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETNEXTITEM, iStart, MAKELONG(flags, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the zero-based index of the currently selected item.
' See the following Microsoft article: How To Select a Listview Item Programmatically
' http://support.microsoft.com/kb/131284
' ========================================================================================
PRIVATE FUNCTION CListView.GetSelection (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETNEXTITEM, -1, LVNI_SELECTED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the item that has the focus, which may or may not be selected.
' Returns the index of the next item if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CListView.GetFocusedItem (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETNEXTITEM, -1, LVNI_ALL OR LVNI_FOCUSED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the item in a particular list-view control that has the specified
' properties and relationship to another specific item.
' ========================================================================================
PRIVATE FUNCTION CListView.GetNextItemIndex (BYVAL hCtl AS HWND, BYREF plvii AS LVITEMINDEX, BYVAL flags AS LONG) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETNEXTITEMINDEX, CAST(WPARAM, @plvii), MAKELONG(flags, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of working areas in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetNumberOfWorkAreas (BYVAL hCtl AS HWND) AS DWORD
   DIM pnWorkAreas AS DWORD
   RETURN SendMessageW(hCtl, LVM_GETNUMBEROFWORKAREAS, 0, CAST(LPARAM, @pnWorkAreas))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current view origin for a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetOrigin (BYVAL hCtl AS HWND, BYREF ptOrg AS POINT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETORIGIN, 0, CAST(LPARAM, @ptOrg))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetOrigin (BYVAL hCtl AS HWND) AS POINT
   DIM ptOrg AS POINT
   SendMessageW(hCtl, LVM_GETORIGIN, 0, CAST(LPARAM, @ptOrg))
   RETURN ptOrg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the color of the border of a list-view control if the LVS_EX_BORDERSELECT extended
' window style is set.
' ========================================================================================
PRIVATE FUNCTION CListView.GetOutlineColor (BYVAL hCtl AS HWND) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_GETOUTLINECOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an integer that specifies the selected column.
' ========================================================================================
PRIVATE FUNCTION CListView.GetSelectedColumn (BYVAL hCtl AS HWND) AS UINT
   RETURN SendMessageW(hCtl, LVM_GETSELECTEDCOLUMN, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the number of selected items in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetSelectedCount (BYVAL hCtl AS HWND) AS UINT
   RETURN SendMessageW(hCtl, LVM_GETSELECTEDCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the selection mark from a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetSelectionMark (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETSELECTIONMARK, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the width of a specified string using the specified list-view control's
' current font.
' ========================================================================================
PRIVATE FUNCTION CListView.GetStringWidth (BYVAL hCtl AS HWND, BYREF wszText AS WSTRING) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETSTRINGWIDTHW, 0, CAST(LPARAM, @wszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about the rectangle that surrounds a subitem in a list-view control.
' This function is intended to be used only on list-view controls that use the LVS_REPORT style.
' ========================================================================================
PRIVATE FUNCTION CListView.GetSubItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYVAL iSubItem AS LONG, BYVAL nCode AS LONG, BYREF rc AS RECT) AS LONG
   IF VARPTR(rc) THEN
      rc.Top = iSubItem
      rc.Left = nCode
   END IF
   RETURN SendMessageW(hCtl, LVM_GETSUBITEMRECT, iItem, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetSubItemRect (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYVAL iSubItem AS LONG, BYVAL nCode AS LONG) AS RECT
   DIM rc AS RECT
   rc.Top = iSubItem
   rc.Left = nCode
   SendMessageW(hCtl, LVM_GETSUBITEMRECT, iItem, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text background color of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetTextBkColor (BYVAL hCtl AS HWND) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_GETTEXTBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text color of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetTextColor (BYVAL hCtl AS HWND) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_GETTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about a tile in a list-view control.
' ========================================================================================
PRIVATE SUB CListView.GetTileInfo (BYVAL hCtl AS HWND, BYREF ti AS LVTILEINFO)
   SendMessageW hCtl, LVM_GETTILEINFO, 0, CAST(LPARAM, @ti)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetTileInfo (BYVAL hCtl AS HWND) AS LVTILEINFO
   DIM ti AS LVTILEINFO
   SendMessageW hCtl, LVM_GETTILEINFO, 0, CAST(LPARAM, @ti)
   RETURN ti
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about a list-view control in tile view.
' ========================================================================================
PRIVATE SUB CListView.GetTileViewInfo (BYVAL hCtl AS HWND, BYREF tvi AS LVTILEVIEWINFO)
   SendMessageW hCtl, LVM_GETTILEVIEWINFO, 0, CAST(LPARAM, @tvi)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetTileViewInfo (BYVAL hCtl AS HWND) AS LVTILEVIEWINFO
   DIM tvi AS LVTILEVIEWINFO
   SendMessageW hCtl, LVM_GETTILEVIEWINFO, 0, CAST(LPARAM, @tvi)
   RETURN tvi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the ToolTip control that the list-view control uses to display ToolTips.
' ========================================================================================
PRIVATE FUNCTION CListView.GetToolTips (BYVAL hCtl AS HWND) AS HWND
   RETURN CAST(HWND, SendMessageW(hCtl, LVM_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of the topmost visible item when in list or report view.
' ========================================================================================
PRIVATE FUNCTION CListView.GetTopIndex (BYVAL hCtl AS HWND) AS LONG
   RETURN SendMessageW(hCtl, LVM_GETTOPINDEX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Unicode character format flag for the control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetUnicodeFormat (BYVAL hCtl AS HWND) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETUNICODEFORMAT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current view of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetView (BYVAL hCtl AS HWND) AS DWORD
   RETURN SendMessageW(hCtl, LVM_GETVIEW, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the bounding rectangle of all items in the list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetViewRect (BYVAL hCtl AS HWND, BYREF rc AS RECT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_GETVIEWRECT, 0, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetViewRect (BYVAL hCtl AS HWND) AS RECT
   DIM rc AS RECT
   SendMessageW(hCtl, LVM_GETVIEWRECT, 0, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the working areas from a list-view control.
' ========================================================================================
PRIVATE SUB CListView.GetWorkAreas (BYVAL hCtl AS HWND, BYVAL nWorkAreas AS LONG, BYREF rc AS RECT)
   SendMessageW hCtl, LVM_GETWORKAREAS, nWorkAreas, CAST(LPARAM, @rc)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.GetWorkAreas (BYVAL hCtl AS HWND, BYVAL nWorkAreas AS LONG) AS RECT
   DIM rc AS RECT
   SendMessageW hCtl, LVM_GETWORKAREAS, nWorkAreas, CAST(LPARAM, @rc)
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the list-view control has a specified group.
' ========================================================================================
PRIVATE FUNCTION CListView.HasGroup (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_HASGROUP, iGroupId, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which list-view item, if any, is at a specified position.
' ========================================================================================
PRIVATE FUNCTION CListView.HitTest (BYVAL hCtl AS HWND, BYREF info AS LVHITTESTINFO) AS LONG
   RETURN SendMessageW(hCtl, LVM_HITTEST, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.HitTest (BYVAL hCtl AS HWND) AS LVHITTESTINFO
   DIM info AS LVHITTESTINFO
   SendMessageW(hCtl, LVM_HITTEST, 0, CAST(LPARAM, @info))
   RETURN info
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which list-view item, if any, is at a specified position.
' ========================================================================================
PRIVATE FUNCTION CListView.HitTestEx (BYVAL hCtl AS HWND, BYREF info AS LVHITTESTINFO) AS LONG
   RETURN SendMessageW(hCtl, LVM_HITTEST, -1, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.HitTestEx (BYVAL hCtl AS HWND) AS LVHITTESTINFO
   DIM info AS LVHITTESTINFO
   SendMessageW(hCtl, LVM_HITTEST, -1, CAST(LPARAM, @info))
   RETURN info
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new column in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.InsertColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYREF col AS LVCOLUMN) AS LONG
   RETURN SendMessageW(hCtl, LVM_INSERTCOLUMN, iCol, CAST(LPARAM, @col))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CListView.InsertColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LVCOLUMN
   DIM col AS LVCOLUMN
   SendMessageW(hCtl, LVM_INSERTCOLUMN, iCol, CAST(LPARAM, @col))
   RETURN col
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new column in a list-view control.
' - hCtl Handle to the ListView control.
' - iCol: Index of the new column.
' - pwszText: The column header text.
' - nWidth: The width of the column, in pixels.
' - nFormat: Alignment of the column header and the subitem text in the column.
'   The alignment of the leftmost column is always left-justified; it cannot be changed.
'   This member can be one of the following values:
'   see https://msdn.microsoft.com/en-us/library/windows/desktop/bb774743%28v=vs.85%29.aspx
' Returns the index of the new column if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION CListView.AddColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL pwszText AS WSTRING PTR, BYVAL nWidth AS LONG, BYVAL nFormat AS LONG = LVCFMT_LEFT) AS LONG
   DIM lvc AS LVCOLUMNW
   lvc.mask = LVCF_FMT OR LVCF_WIDTH OR LVCF_TEXT OR LVCF_SUBITEM
   lvc.fmt = nFormat
   lvc.pszText = pwszText
   lvc.cx = nWidth
   lvc.iSubItem = iCol
   RETURN SendMessageW(hCtl, LVM_INSERTCOLUMNW, iCol, CAST(LPARAM, @lvc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a group into a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.InsertGroup (BYVAL hCtl AS HWND, BYVAL index AS LONG, BYREF grp AS LVGROUP) AS LONG
   RETURN SendMessageW(hCtl, LVM_INSERTGROUP, index, CAST(LPARAM, @grp))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a group into an ordered list of groups.
' ========================================================================================
PRIVATE SUB CListView.InsertGroupSorted (BYVAL hCtl AS HWND, BYREF structInsert AS LVINSERTGROUPSORTED)
   SendMessageW hCtl, LVM_INSERTGROUPSORTED, CAST(LPARAM, @structInsert), 0
END SUB
' ========================================================================================

' ========================================================================================
' Inserts a new item in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.InsertItem (BYVAL hCtl AS HWND, BYREF item AS LVITEM) AS LONG
   RETURN SendMessageW(hCtl, LVM_INSERTITEM, 0, CAST(LPARAM, @item))
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
PRIVATE FUNCTION CListView.AddItem (BYVAL hCtl AS HWND, BYVAL iItem AS LONG, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DIM lvi AS LVITEMW
   lvi.mask = LVIF_TEXT
   lvi.pszText = pwszText
   lvi.iItem = iItem
   lvi.iImage = iImage
   FUNCTION = SendMessageW(hCtl, LVM_INSERTITEMW, iItem, cast(LPARAM, @lvi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a new item in a list-view control.
' - hwndLV: Handle to the ListView control.
' - nImage: Zero-based index of the item's icon in the control's image list.
' - pwszText: The item text.
' Returns the index of the new item if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION ListView_AppendItem (BYVAL hCtl AS HWND, BYVAL iImage AS LONG, BYVAL pwszText AS WSTRING PTR) AS LONG
   DIM lvi AS LVITEMW
   DIM iItem AS LONG = SendMessageW(hCtl, LVM_GETITEMCOUNT, 0, 0)
   lvi.mask = LVIF_TEXT
   lvi.pszText = pwszText
   lvi.iItem = iItem
   lvi.iImage = iImage
   RETURN SendMessageW(hCtl, LVM_INSERTITEMW, iItem, CAST(LPARAM, @lvi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the insertion point closest to a specified point.
' ========================================================================================
PRIVATE FUNCTION CListView.InsertMarkHitTest (BYVAL hCtl AS HWND, BYREF pt AS POINT, BYREF vim AS LVINSERTMARK) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_INSERTMARKHITTEST, CAST(WPARAM, @pt), CAST(LPARAM, @vim))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Checks whether the list-view control has group view enabled.
' ========================================================================================
PRIVATE FUNCTION CListView.IsGroupViewEnabled (BYVAL hCtl AS HWND) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_ISGROUPVIEWENABLED, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether an item in the list-view control is visible.
' ========================================================================================
PRIVATE FUNCTION CListView.IsItemVisible (BYVAL hCtl AS HWND, BYVAL index AS DWORD) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_ISITEMVISIBLE, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps the ID of an item to an index.
' ========================================================================================
PRIVATE FUNCTION CListView.MapIDToIndex (BYVAL hCtl AS HWND, BYVAL id AS UINT) AS UINT
   RETURN SendMessageW(hCtl, LVM_MAPIDTOINDEX, id, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps the index of an item to a unique ID.
' ========================================================================================
PRIVATE FUNCTION CListView.MapIndexToID (BYVAL hCtl AS HWND, BYVAL index AS UINT) AS UINT
   RETURN SendMessageW(hCtl, LVM_MAPINDEXTOID, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE SUB CListView.MoveGroup (BYVAL hwnd AS DWORD, BYVAL iGroupId AS LONG, BYVAL toIndex AS LONG)
'   SendMessageW(hCtl, LVM_MOVEGROUP, iGroupId, toIndex)
'END SUB
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE SUB ListView_MoveItemToGroup (BYVAL hCtl AS HWND, BYVAL idItemFrom AS LONG, BYVAL idGroupTo AS LONG)
'   SendMessageW(hCtl, LVM_MOVEITEMTOGROUP, idItemFrom, idGroupTo)
'END SUB
' ========================================================================================

' ========================================================================================
' Forces a list-view control to redraw a range of items.
' ========================================================================================
PRIVATE FUNCTION CListView.RedrawItems (BYVAL hCtl AS HWND, BYVAL iFirst AS LONG, BYVAL iLast AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_REDRAWITEMS, iFirst, iLast)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all groups from a list-view control.
' ========================================================================================
PRIVATE SUB CListView.RemoveAllGroups (BYVAL hCtl AS HWND)
   SendMessageW hCtl, LVM_REMOVEALLGROUPS, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Removes a group from a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.RemoveGroup (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG) AS LONG
   RETURN SendMessageW(hCtl, LVM_REMOVEGROUP, iGroupId, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the content of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.Scroll (BYVAL hCtl AS HWND, BYVAL dx AS LONG, BYVAL dy AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SCROLL, dx, dy)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetBkColor (BYVAL hCtl AS HWND, BYVAL clrBk AS COLORREF) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETBKCOLOR, 0, clrBk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background image in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetBkImage (BYVAL hCtl AS HWND, BYREF vbki AS LVBKIMAGE) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETBKIMAGE, 0, CAST(LPARAM, @vbki))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the callback mask for a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetCallbackMask (BYVAL hCtl AS HWND, BYVAL mask AS DWORD) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETCALLBACKMASK, mask, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects or deselects an item in a list-view control.
' ========================================================================================
PRIVATE SUB CListView.SetCheckState (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG, BYVAL fCheck AS BOOLEAN)
   DIM lvi AS LVITEM
   lvi.stateMask = LVIS_STATEIMAGEMASK
   lvi.state = INDEXTOSTATEIMAGEMASK(IIF(fCheck, 2, 1))
   SendMessageW hCtl, LVM_SETITEMSTATE, iIndex, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the attributes of a list-view column.
' ========================================================================================
PRIVATE FUNCTION CListView.SetColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYREF col AS LVCOLUMN) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETCOLUMN, iCol, CAST(LPARAM, @col))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the left-to-right order of columns in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetColumnOrderArray (BYVAL hCtl AS HWND, BYVAL iCount AS LONG, BYVAL pi AS LONG PTR) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETCOLUMNORDERARRAY, iCount, CAST(LPARAM, pi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Used to change the width of a column in report view or the width of all columns in list-view mode.
' ========================================================================================
PRIVATE FUNCTION CListView.SetColumnWidth (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL cx AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, iCol, MAKELONG(cx, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets extended styles for list-view controls
' ========================================================================================
PRIVATE SUB CListView.SetExtendedListViewStyle (BYVAL hCtl AS HWND, BYVAL dwExStyle AS DWORD)
   SendMessageW hCtl, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, dwExStyle
END SUB
' ========================================================================================

' ========================================================================================
' Sets extended styles for list-view controls using the style mask.
' ========================================================================================
PRIVATE SUB CListView.SetExtendedListViewStyleEx (BYVAL hCtl AS HWND, BYVAL dwExMask AS DWORD, BYVAL dwExStyle AS DWORD)
   SendMessageW hCtl, LVM_SETEXTENDEDLISTVIEWSTYLE, dwExMask, dwExStyle
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an image list to the group header of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetGroupHeaderImageList (BYVAL hCtl AS HWND, BYVAL himl AS HIMAGELIST) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(hCtl, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, CAST(LPARAM, himl)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets group information.
' ========================================================================================
PRIVATE FUNCTION CListView.SetGroupInfo (BYVAL hCtl AS HWND, BYVAL iGroupId AS LONG, BYREF grp AS LVGROUP) AS LONG
   RETURN SendMessageW(hCtl, LVM_SETGROUPINFO, iGroupId, CAST(LPARaM, @grp))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets information about the display of groups.
' ========================================================================================
PRIVATE SUB CListView.SetGroupMetrics (BYVAL hCtl AS HWND, BYREF GroupMetrics AS LVGROUPMETRICS)
   SendMessageW hCtl, LVM_SETGROUPMETRICS, 0, CAST(LPARAM, @GroupMetrics)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the state for a specified group.
' ========================================================================================
PRIVATE FUNCTION CListView.SetGroupState (BYVAL hCtl AS HWND, BYVAL dwGroupId AS UINT, BYVAL dwMask AS UINT, BYVAL dwState AS UINT) AS LONG
   DIM lvg AS LVGROUP
   lvg.cbSize = SIZEOF(lvg)
   lvg.mask = LVGF_STATE
   lvg.stateMask = dwMask
   lvg.state = dwState
   RETURN SendMessageW(hCtl, LVM_SETGROUPINFO, dwGroupId, CAST(LPARAM, @lvg))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the HCURSOR that the list-view control uses when the pointer is over an item while
' hot tracking is enabled.
' ========================================================================================
PRIVATE FUNCTION CListView.SetHotCursor (BYVAL hCtl AS HWND, BYVAL hCur AS HCURSOR) AS HCURSOR
   RETURN CAST(HCURSOR, SendMessageW(hCtl, LVM_SETHOTCURSOR, 0, CAST(LPARAM, hCur)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the hot item in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetHotItem (BYVAL hCtl AS HWND, BYVAL iIndex AS DWORD) AS LONG
   RETURN SendMessageW(hCtl, LVM_SETHOTITEM, iIndex, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the amount of time that the mouse cursor must hover over an item before it is selected.
' ========================================================================================
PRIVATE SUB CListView.SetHoverTime (BYVAL hCtl AS HWND, BYVAL dwHoverTimeMs AS DWORD)
   SendMessageW(hCtl, LVM_SETHOVERTIME, 0, dwHoverTimeMs)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the spacing between icons in list-view controls set to the LVS_ICON style.
' ========================================================================================
PRIVATE FUNCTION CListView.SetIconSpacing (BYVAL hCtl AS HWND, BYVAL cx AS SHORT, BYVAL cy AS SHORT) AS DWORD
   RETURN SendMessageW(hCtl, LVM_SETICONSPACING, 0, MAKELONG(cx, cy))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an image list to a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetImageList (BYVAL hCtl AS HWND, BYVAL himl AS HIMAGELIST, BYVAL nImageList AS LONG) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(hCtl, LVM_SETIMAGELIST, nImageList, CAST(LPARAM, himl)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets tooltip text.
' ========================================================================================
PRIVATE FUNCTION CListView.SetInfoTip (BYVAL hCtl AS HWND, BYREF lvInfoTip AS LVSETINFOTIP) AS LONG
   RETURN SendMessageW(hCtl, LVM_SETINFOTIP, 0, CAST(LPARAM, @lvInfoTip))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the insertion point to the defined position.
' ========================================================================================
PRIVATE FUNCTION CListView.SetInsertMark (BYVAL hCtl AS HWND, BYREF lvim AS LVINSERTMARK) AS LONG
   RETURN SendMessageW(hCtl, LVM_SETINSERTMARK, 0, CAST(LPARAM, @lvim))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the color of the insertion point.
' ========================================================================================
PRIVATE FUNCTION CListView.SetInsertMarkColor (BYVAL hCtl AS HWND, BYVAL colr AS COLORREF) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_SETINSERTMARKCOLOR, 0, colr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets some or all of a list-view item's attributes. You can also use ListView_SetItem to
' set the text of a subitem.
' ========================================================================================
PRIVATE FUNCTION CListView.SetItem (BYVAL hCtl AS HWND, BYREF item AS LVITEM) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETITEM, 0, CAST(LPARAM, @item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes the list-view control to allocate memory for the specified number of items.
' ========================================================================================
PRIVATE SUB CListView.SetItemCount (BYVAL hCtl AS HWND, BYVAL cItems AS LONG)
   SendMessageW hCtl, LVM_SETITEMCOUNT, cItems, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the virtual number of items in a virtual list view.
' ========================================================================================
PRIVATE SUB CListView.SetItemCountEx (BYVAL hCtl AS HWND, BYVAL cItems AS LONG, BYVAL dwFlags AS LONG)
   SendMessageW hCtl, LVM_SETITEMCOUNT, cItems, dwFlags
END SUB
' ========================================================================================

' ========================================================================================
' Sets the state of a specified list-view item.
' ========================================================================================
PRIVATE FUNCTION CListView.SetItemIndexState (BYVAL hCtl AS HWND, BYREF lvii AS LVITEMINDEX, BYVAL dwData AS UINT, BYVAL mask AS UINT) AS HRESULT
   DIM lvi AS LVITEM
   lvi.stateMask = mask
   lvi.state = dwData
   RETURN SendMessageW(hCtl, LVM_SETITEMINDEXSTATE, CAST(WPARAM, @lvii), Cast(LPARAM, @lvi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves an item to a specified position in a list-view control (in icon or small icon view).
' ========================================================================================
PRIVATE FUNCTION CListView.SetItemPosition (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL x AS SHORT, BYVAL y AS SHORT) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETITEMPOSITION, i, MAKELONG(x, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves an item to a specified position in a list-view control (in icon or small icon view).
' This macro differs from the ListView_SetItemPosition macro in that it uses 32-bit coordinates.
' ========================================================================================
PRIVATE SUB CListView.SetItemPosition32 (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL x AS LONG, BYVAL y AS LONG)
   DIM ptNewPos AS POINT
   ptNewPos.x = x
   ptNewPos.y = y
   SendMessageW  hCtl, LVM_SETITEMPOSITION32, i, CAST(LPARAM, @ptNewPos)
END SUB
' ========================================================================================

' ========================================================================================
' Changes the state of an item in a list-view control.
' ========================================================================================
PRIVATE SUB CListView.SetItemState (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL dwState AS DWORD, BYVAL mask AS DWORD)
   DIM lvi AS LVITEM
   lvi.stateMask = mask
   lvi.state = dwState
   SendMessageW hCtl, LVM_SETITEMSTATE, i, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Changes the text of a list-view item or subitem.
' ========================================================================================
PRIVATE SUB CListView.SetItemText (BYVAL hCtl AS HWND, BYVAL i AS LONG, BYVAL iSubItem AS LONG, BYREF wszText AS WSTRING)
   DIM lvi AS LVITEMW
   lvi.iSubItem = iSubItem
   lvi.pszText = @wszText
   SendMessageW hCtl, LVM_SETITEMTEXTW, i, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the color of the border of a list-view control if the LVS_EX_BORDERSELECT extended
' window style is set.
' ========================================================================================
PRIVATE FUNCTION CListView.SetOutlineColor (BYVAL hCtl AS HWND, BYVAL colr AS COLORREF) AS COLORREF
   RETURN SendMessageW(hCtl, LVM_SETOUTLINECOLOR, 0, colr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the index of the selected column.
' ========================================================================================
PRIVATE SUB CListView.SetSelectedColumn (BYVAL hCtl AS HWND, BYVAL iCol AS LONG)
   SendMessageW hCtl, LVM_SETSELECTEDCOLUMN, iCol, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the selection mark in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetSelectionMark (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG) AS LONG
   RETURN SendMessageW(hCtl, LVM_SETSELECTIONMARK, 0, iIndex)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of text in a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetTextBkColor (BYVAL hCtl AS HWND, BYVAL clrTextBk AS COLORREF) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETTEXTBKCOLOR, 0, clrTextBk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text color of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetTextColor (BYVAL hCtl AS HWND, BYVAL clrText AS COLORREF) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETTEXTCOLOR, 0, clrText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets information for an existing tile of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetTileInfo (BYVAL hCtl AS HWND, BYREF lvtinfo AS LVTILEINFO) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETTILEINFO, 0, CAST(LPARAM, @lvtinfo))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets information that a list-view control uses in tile view.
' ========================================================================================
PRIVATE FUNCTION CListView.SetTileViewInfo (BYVAL hCtl AS HWND, BYREF lvtvinfo AS LVTILEVIEWINFO) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETTILEVIEWINFO, 0, CAST(LPARAM, @lvtvinfo))
END FUNCTION
' ========================================================================================

' ========================================================================================
' This message is supported only through Microsoft Windows XP.
' Currently not supported.
' ========================================================================================
'PRIVATE FUNCTION CListView.SetTileWidth (BYVAL hwnd AS DWORD, BYVAL cpWidth AS DWORD) AS LONG
'   RETURN SendMessageW(hwnd, LVM_SETTILEWIDTH, cpWidth, 0)
'END FUNCTION
' ========================================================================================
 
 ' ========================================================================================
' Sets the ToolTip control that the list-view control will use to display ToolTips.
' ========================================================================================
PRIVATE FUNCTION CListView.SetToolTips (BYVAL hCtl AS HWND, BYVAL hwndToolTip AS HWND) AS HWND
   RETURN CAST(HWND, SendMessageW(hCtl, LVM_SETTOOLTIPS, CAST(LPARAM, hwndToolTip), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the Unicode character format flag for the control. This message allows you to
' change the character set used by the control at run time rather than having to re-create
' the control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetUnicodeFormat (BYVAL hCtl AS HWND, BYVAL fUnicode AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETUNICODEFORMAT, fUnicode, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the view of a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SetView (BYVAL hCtl AS HWND, BYVAL iView AS DWORD) AS LONG
   RETURN SendMessageW(hCtl, LVM_SETVIEW, iView, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the working areas within a list-view control.
' ========================================================================================
PRIVATE SUB CListView.SetWorkAreas (BYVAL hCtl AS HWND, BYVAL nWorkAreas AS LONG, BYREF rc AS RECT)
   SendMessageW hCtl, LVM_SETWORKAREAS, nWorkAreas, CAST(LPARAM, @rc)
END SUB
' ========================================================================================

' ========================================================================================
' Uses an application-defined comparison function to sort groups by ID within a list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.SortGroups (BYVAL hCtl AS HWND, BYVAL _pfnGroupCompare AS PFNLVGROUPCOMPARE PTR, BYVAL _plv AS ANY PTR) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SORTGROUPS, CAST(WPARAM, _pfnGroupCompare), CAST(LPARAM, _plv))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses an application-defined comparison function to sort the items of a list-view control.
' The index of each item changes to reflect the new sequence.
' ========================================================================================
PRIVATE FUNCTION CListView.SortItems (BYVAL hCtl AS HWND, BYVAL _pfnCompare AS PFNLVCOMPARE PTR, BYVAL _lPrm AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SORTITEMS, CAST(WPARAM, _pfnCompare), CAST(LPARAM, _lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses an application-defined comparison function to sort the items of a list-view control.
' The index of each item changes to reflect the new sequence.
' ========================================================================================
PRIVATE FUNCTION CListView.SortItemsEx (BYVAL hCtl AS HWND, BYVAL _pfnCompare AS PFNLVCOMPARE PTR, BYVAL _lPrm AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SORTITEMSEX, CAST(WPARAM, _pfnCompare), CAST(LPARAM, _lPrm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which list-view item or subitem is located at a given position.
' ========================================================================================
PRIVATE FUNCTION CListView.SubItemHitTest (BYVAL hCtl AS HWND, BYREF lvhti AS LVHITTESTINFO) AS LONG
   RETURN SendMessageW(hCtl, LVM_SUBITEMHITTEST, 0, CAST(LPARAM, @lvhti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which list-view item or subitem is located at a given position.
' ========================================================================================
PRIVATE FUNCTION CListView.SubItemHitTestEx (BYVAL hCtl AS HWND, BYREF lvhti AS LVHITTESTINFO) AS LONG
   RETURN SendMessageW(hCtl, LVM_SUBITEMHITTEST, -1, CAST(LPARAM, @lvhti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates a list-view item. If the list-view control has the LVS_AUTOARRANGE style, this
' macro causes the list-view control to be arranged.
' ========================================================================================
PRIVATE FUNCTION CListView.Update (BYVAL hCtl AS HWND, BYVAL iItem AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_UPDATE, iItem, 0)
END FUNCTION
' ========================================================================================

' ########################################################################################


' ========================================================================================
' Selects a ListView item.
' Windows does not provide a separate message or function to set the current selection in
' a listview. Instead, it defines item states or LVIS_* values that determine the listview
' item's appearance and functionality. LVIS_FOCUSED and LVIS_SELECTED in particular are
' the states that determine a listview item's selection state.
' Note that the last parameter passed to this macro is a mask specifying which bits are
' about to change. LVIS_FOCUSED and LVIS_SELECTED are defined in commctrl.inc as &H0001
' and &H0002 respectively, so you need to set the last four bits of the mask.
' See the following Microsoft article: How To Select a Listview Item Programmatically
' http://support.microsoft.com/kb/131284
' ========================================================================================
PRIVATE SUB CListView.SelectItem (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG)
   DIM lvi AS LVITEM
   lvi.stateMask = &h000F
   lvi.state = LVIS_FOCUSED OR LVIS_SELECTED
   SendMessageW hCtl, LVM_SETITEMSTATE, iIndex, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Selects all items.
' ========================================================================================
PRIVATE SUB CListView.SelectAllItems (BYVAL hCtl AS HWND)
   DIM lvi AS LVITEM
   lvi.mask = LVIF_STATE
   lvi.state = -1
   lvi.stateMask = LVIS_SELECTED
   SendMessageW hCtl, LVM_SETITEMSTATE, -1, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Unselects an item in a ListView.
' ========================================================================================
PRIVATE SUB CListView.UnselectItem (BYVAL hCtl AS HWND, BYVAL iIndex AS LONG)
   DIM lvi AS LVITEM
   lvi.mask = LVIF_STATE
   lvi.state = 0
   lvi.state = LVIS_SELECTED
   SendMessageW hCtl, LVM_SETITEMSTATE, iIndex, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Unselects all items.
' ========================================================================================
PRIVATE SUB CListView.UnselectAllItems (BYVAL hCtl AS HWND)
   DIM lvi AS LVITEM
   lvi.stateMask = LVIS_SELECTED OR LVIS_FOCUSED
   lvi.state = 0
   SendMessageW hCtl, LVM_SETITEMSTATE, -1, CAST(LPARAM, @lvi)
END SUB
' ========================================================================================

' ========================================================================================
' Removes the HDS_BUTTONS style from the header control to give it a flat appearance.
' ========================================================================================
PRIVATE SUB CListView.MakeHeaderFlat (BYVAL hCtl AS HWND)
   DIM hLvHeader AS HWND
   hLvHeader = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN EXIT SUB
   SetWindowLong hLvHeader, GWL_STYLE, GetWindowLong(hLvHeader, GWL_STYLE) XOR HDS_BUTTONS
   SetWindowPos GetParent(hCtl), NULL, 0, 0, 0, 0, SWP_NOZORDER OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_DRAWFRAME
END SUB
' ========================================================================================

' ========================================================================================
' Automatically sizes the specified column.
' ========================================================================================
PRIVATE FUNCTION CListView.FitContent (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, iCol, LVSCW_AUTOSIZE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Automatically sizes all the columns of a listview
' ========================================================================================
PRIVATE FUNCTION CListView.AutoSizeColumns (BYVAL hCtl AS HWND) AS BOOLEAN
   DIM AS LONG i, nCount
   DIM hLvHeader AS HWND
   hLvHeader = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN EXIT FUNCTION
   nCount = SendMessageW(hLvHeader, HDM_GETITEMCOUNT, 0, 0)
   FOR i = 0 TO nCount - 1
      FUNCTION = SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE)
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Automatically sizes the column to fit the header text. If you use this value with the
' last column, its width is set to fill the remaining width of the list-view control.
' ========================================================================================
PRIVATE FUNCTION CListView.FitHeader (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS BOOLEAN
   RETURN SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, iCol, LVSCW_AUTOSIZE_USEHEADER)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Automatically sizes all the columns of the header of a listview
' ========================================================================================
PRIVATE FUNCTION CListView.AutoSizeHeader (BYVAL hCtl AS HWND) AS BOOLEAN
   DIM AS LONG i, nCount
   DIM hLvHeader AS HWND
   hLvHeader = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN EXIT FUNCTION
   nCount = SendMessageW(hLvHeader, HDM_GETITEMCOUNT, 0, 0)
   FOR i = 0 TO nCount - 1
      FUNCTION = SendMessageW(hCtl, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER)
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of columns of a listview.
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumnCount (BYVAL hCtl AS HWND) AS LONG
   DIM hLvHeader AS HWND = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN EXIT FUNCTION
   RETURN SendMessageW(hLvHeader, HDM_GETITEMCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of the specified column of the header of a listview control.
' - hCtl Handle to the ListView control.
' - nCol: Zero-based index of the column
' - pwszText: A pointer to a buffer that receives the item or subitem text.
' - cchTextMax : Maximum number of characters to return. Both this value and the size of the
'   buffer pointed by pwszText must be one character bigger that the wanted length of the text
'   to return to make room for the null character terminator.
' Returns the number of characters in the pszText member of the LVITEM structure.
' ========================================================================================
PRIVATE FUNCTION CListView.GetHeaderText (BYVAL hCtl AS HWND, BYVAL iCol AS LONG, BYVAL cchTextMax AS LONG = MAX_PATH + 1) AS DWSTRING
   DIM hLvHeader AS HWND = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN RETURN ""
   DIM dwsText AS DWSTRING
   IF cchTextMax < 1 THEN cchTextMax = MAX_PATH + 1
   dwsText = WSPACE(cchTextMax + 1)
   DIM hdi AS HDITEMW
   hdi.mask = HDI_TEXT
   hdi.cchTextMax = cchTextMax + 1
   hdi.pszText = *dwsText
   SendMessageW(hLvHeader, HDM_GETITEMW, iCol, CAST(LPARAM, @hdi))
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the order of the specified column of the listview's header control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetColumnOrder (BYVAL hCtl AS HWND, BYVAL iCol AS LONG) AS LONG
   DIM hLvHeader AS HWND =CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN RETURN 0
   DIM hdi AS HDITEMW
   hdi.mask = HDI_ORDER
   SendMessageW(hLvHeader, HDM_GETITEM, iCol, CAST(LPARAM, @hdi))
   RETURN hdi.iOrder
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current left-to-right order of items in the header of a listview control.
' ========================================================================================
PRIVATE FUNCTION CListView.GetOrderArray (BYVAL hCtl AS HWND, BYVAL iSize AS DWORD, BYVAL pi AS LONG PTR) AS LONG
   DIM hLvHeader AS HWND = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN RETURN 0
   RETURN SendMessageW(hLvHeader, HDM_GETORDERARRAY, iSize, CAST(LPARAM, pi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the font used by the ToolTip control that the list-view control uses to display ToolTips.
' The return value is a handle to the font used by the control, or NULL if the control is
' using the system font.
' ========================================================================================
PRIVATE FUNCTION CListView.GetTooltipsFont (BYVAL hCtl AS HWND) AS HFONT
   DIM hwndTooltips AS HWND = CAST(HWND, SendMessageW(hCtl, LVM_GETTOOLTIPS, 0, 0))
   FUNCTION = CAST(HFONT, SendMessageW(hwndTooltips, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font used by the ToolTip control that the list-view control uses to display ToolTips.
' Parameters:
' - hCtl    = Handle to the ListView control.
' - hFont   = Handle to the font. If this parameter is NULL, the control uses the default system font to draw text.
' - fRedraw = Specifies whether the control should be redrawn immediately upon setting the font.
'             If this parameter is TRUE, the control redraws itself.
' ========================================================================================
PRIVATE SUB CListView.SetTooltipsFont (BYVAL hCtl AS HWND, BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
   DIM hwndTooltips AS HWND = CAST(HWND, SendMessageW(hCtl, LVM_GETTOOLTIPS, 0, 0))
   IF hwndTooltips = NULL THEN EXIT SUB
   SendMessageW(hwndTooltips, WM_SETFONT, CAST(WPARAM, hFont), fRedraw)
END SUB
' ========================================================================================


' ========================================================================================
PRIVATE FUNCTION CListView.SetHeaderFont ( _
   BYVAL hCtl        AS HWND, _      ' __in Handle to the listview
   BYREF wszFaceName AS WSTRING, _   ' __in Typeface name of font
   BYVAL lPointSize  AS LONG, _      ' __in Point size
   BYVAL lWeight     AS LONG, _      ' __in Font weight(bold etc.)
   BYVAL bItalic     AS BYTE, _      ' __in TRUE = italic
   BYVAL bUnderline  AS BYTE, _      ' __in TRUE = underline
   BYVAL bStrikeOut  AS BYTE, _      ' __in TRUE = strikeout
   BYVAL bCharSet    AS BYTE _       ' __in character set
   ) AS HFONT                        ' Handle of font or NULL on failure.

   DIM hLvHeader AS HWND
   DIM hLvHeaderFont AS HFONT
   DIM hOldFont AS HFONT
   DIM lf AS LOGFONTW
   DIM hDC AS HDC

   ' // Get the handle of the header
   hLvHeader = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN RETURN NULL

   ' // Create the font
   hDC = GetDC(HWND_DESKTOP)
   lf.lfHeight         = -MulDiv(lPointSize, GetDeviceCaps(hDC, LOGPIXELSY), 72) ' logical font height
   lf.lfWidth          =  0                                                       ' average character width
   lf.lfEscapement     =  0                                                       ' escapement
   lf.lfOrientation    =  0                                                       ' orientation angles
   lf.lfWeight         =  lWeight                                                 ' font weight
   lf.lfItalic         =  bItalic                                                 ' italic(TRUE/FALSE)
   lf.lfUnderline      =  bUnderline                                              ' underline(TRUE/FALSE)
   lf.lfStrikeOut      =  bStrikeOut                                              ' strikeout(TRUE/FALSE)
   lf.lfCharSet        =  bCharset                                                ' character set
   lf.lfOutPrecision   =  OUT_TT_PRECIS                                           ' output precision
   lf.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                     ' clipping precision
   lf.lfQuality        =  DEFAULT_QUALITY                                         ' output quality
   lf.lfPitchAndFamily =  FF_DONTCARE                                             ' pitch and family
   lf.lfFaceName       =  wszFaceName                                             ' typeface name
   ReleaseDC HWND_DESKTOP, hDC
   hLvHeaderFont = CreateFontIndirectW(@lf)
   IF hLvHeaderFont = NULL THEN RETURN NULL

   ' // Select the font
   hDC = GetDC(hLvHeader)
   hOldFont = SelectObject(hDC, hLvHeaderFont)
   ReleaseDC (hLvHeader, hDC)
   SendMessageW(CAST(HWND, hLvHeader), WM_SETFONT, CAST(WPARAM, hLvHeaderFont), TRUE)
   IF hOldFont THEN DeleteObject(hOldFont)
   RETURN hLvHeaderFont

END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies the font used by the listview header.
' Only the passed values that are not an empty string or zero will be modified.
' Note: The returned font must be destroyed with DeleteObject when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION CListView.ModifyHeaderFont ( _
   BYVAL hCtl        AS HWND, _      ' __in Handle to the listview
   BYREF wszFaceName AS WSTRING, _   ' __in Typeface name of font
   BYVAL lPointSize  AS LONG, _      ' __in Point size
   BYVAL lWeight     AS LONG, _      ' __in Font weight(bold etc.)
   BYVAL bItalic     AS BYTE, _      ' __in TRUE = italic
   BYVAL bUnderline  AS BYTE, _      ' __in TRUE = underline
   BYVAL bStrikeOut  AS BYTE, _      ' __in TRUE = strikeout
   BYVAL bCharSet    AS BYTE _       ' __in character set
   ) AS HFONT                        ' Handle of font or NULL on failure.

   DIM hLvHeader AS HWND
   DIM hLvHeaderFont AS HFONT
   DIM hCurFont AS HFONT
   DIM hOldFont AS HFONT
   DIM lf AS LOGFONTW
   DIM hDC AS HDC

   ' // Get the handle of the header
   hLvHeader = CAST(HWND, SendMessageW(hCtl, LVM_GETHEADER, 0, 0))
   IF hLvHeader = 0 THEN RETURN NULL

   ' // Get the handle of the font used by the header
   hCurFont = CAST(HFONT, SendMessageW(hLvHeader, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN EXIT FUNCTION
   ' // Get the LOGFONT structure
   IF GetObjectW(hCurFont, SIZEOF(lf), @lf) = 0 THEN RETURN NULL

   ' // Change the requested values
   IF lPointSize THEN
      hDC = GetDC(HWND_DESKTOP)
      lf.lfHeight = -MulDiv(lPointSize, GetDeviceCaps(hDC, LOGPIXELSY), 72)
      ReleaseDC HWND_DESKTOP, hDC
   END IF
   IF lWeight THEN lf.lfWeight =  lWeight
   IF bItalic THEN lf.lfItalic =  bItalic
   IF bUnderline THEN lf.lfUnderline =  bUnderline
   IF bStrikeOut THEN lf.lfStrikeOut =  bStrikeOut
   IF bCharset THEN lf.lfCharSet =  bCharset
   IF wszFaceName <> "" THEN lf.lfFaceName =  wszFaceName

   ' // Create the font
   hLvHeaderFont = CreateFontIndirectW(@lf)
   IF hLvHeaderFont = 0 THEN EXIT FUNCTION

   ' // Select the font
   hDC = GetDC(hLvHeader)
   hOldFont = SelectObject(hDC, hLvHeaderFont)
   ReleaseDC (hLvHeader, hDC)
   SendMessageW(hLvHeader, WM_SETFONT, CAST(WPARAM, hLvHeaderFont), TRUE)
   IF hOldFont THEN DeleteObject(hOldFont)
   RETURN hLvHeaderFont

END FUNCTION
' ========================================================================================

END NAMESPACE
' ########################################################################################
