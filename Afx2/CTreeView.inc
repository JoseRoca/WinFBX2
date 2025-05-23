' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTreeView.inc
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
' CTreeView class
' ########################################################################################

TYPE CTreeView EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION GetItem (BYVAL pitem AS TVITEMW PTR) AS BOOLEAN
   DECLARE FUNCTION GetItemEx (BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
   DECLARE FUNCTION InsertItem (BYVAL pis AS TVINSERTSTRUCTW PTR) AS HTREEITEM
   DECLARE FUNCTION Insert (BYVAL hParent AS HTREEITEM, BYVAL hInsertAfter AS HTREEITEM, _
      BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0, BYVAL pwszText AS WSTRING PTR) AS HTREEITEM
   DECLARE FUNCTION InsertItem (BYVAL hParent AS HTREEITEM, BYVAL hInsertAfter AS HTREEITEM, _
      BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   DECLARE FUNCTION AppendItem (BYVAL hParent AS HTREEITEM, _
      BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   DECLARE FUNCTION InsertFirstItem (BYVAL hParent AS HTREEITEM, _
      BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   DECLARE FUNCTION AddRootItem (BYVAL pwszText AS WSTRING PTR, _
      BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   DECLARE FUNCTION InsertSortedItem (BYVAL hParent AS HTREEITEM, _
      BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   DECLARE FUNCTION DeleteItem (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION Delete_ (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION DeleteAllItems () AS BOOLEAN
   DECLARE FUNCTION Reset () AS BOOLEAN
   DECLARE FUNCTION GetItemState (BYVAL hItem AS HTREEITEM, BYVAL mask AS UINT) AS UINT
   DECLARE FUNCTION GetBold (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE SUB SetItemState (BYVAL hItem AS HTREEITEM, BYVAL nState AS UINT, BYVAL statemask AS UINT)
   DECLARE SUB SetBold (BYVAL hItem AS HTREEITEM, BYVAL fBold AS BOOLEAN)
   DECLARE FUNCTION GetCheckState (BYVAL hItem AS HTREEITEM) AS UINT
   DECLARE FUNCTION IsItemChecked (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION GetCheck (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION GetISearchStringLen () AS LONG
   DECLARE FUNCTION GetISearchString (BYVAL pwsz AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION GetISearchString () AS DWSTRING
   DECLARE FUNCTION GetlParam (BYVAL hItem AS HTREEITEM) AS LPARAM
   DECLARE FUNCTION Expand (BYVAL hItem AS HTREEITEM, BYVAL code AS UINT) AS BOOLEAN
   DECLARE FUNCTION Expand (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION ExpandPartial (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE SUB ExpandChildItems (BYVAL hNode AS HTREEITEM)
   DECLARE SUB ExpandAllChildItems (BYVAL hNode AS HTREEITEM)
   DECLARE SUB ExpandSiblings (BYVAL hNode AS HTREEITEM)
   DECLARE SUB ExpandAllItems 
   DECLARE FUNCTION Collapse (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION CollapseReset (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE SUB CollapseChildItems (BYVAL hNode AS HTREEITEM)
   DECLARE SUB CollapseSiblings (BYVAL hNode AS HTREEITEM)
   DECLARE SUB CollapseAllChildItems (BYVAL hNode AS HTREEITEM)
   DECLARE SUB CollapseAllItems 
   DECLARE FUNCTION Toggle (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION GetNextItem (BYVAL hItem AS HTREEITEM, BYVAL code AS UINT) AS HTREEITEM
   DECLARE FUNCTION GetChild (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetSelection () AS HTREEITEM
   DECLARE FUNCTION GetSelect () AS HTREEITEM
   DECLARE FUNCTION GetDropHilight () AS HTREEITEM
   DECLARE FUNCTION GetFirstVisible () AS HTREEITEM
   DECLARE FUNCTION GetNextSibling (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetNext (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetNextSelected (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetNextVisible (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetParent (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetItemText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchtextmax AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetItemText (BYVAL hItem AS HTREEITEM, BYVAL cchTextMax AS LONG = 0) AS DWSTRING
   DECLARE FUNCTION GetParentText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetParentText (BYVAL hItem AS HTREEITEM, BYVAL cchTextMax AS LONG = 0) AS DWSTRING
   DECLARE FUNCTION GetRoot () AS HTREEITEM
   DECLARE FUNCTION GetPrevSibling (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetPrevious (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetPrevVisible (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   DECLARE FUNCTION GetLastVisible () AS HTREEITEM
   DECLARE FUNCTION SetItem (BYVAL pitem AS TVITEMW PTR) AS BOOLEAN
   DECLARE FUNCTION SetItemEx (BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
   DECLARE FUNCTION SetItemText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION SetText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION ItemExists (BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION ItemExistsIgnoreCase (BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN

   DECLARE FUNCTION CreateDragImage ( BYVAL hItem AS HTREEITEM) AS HIMAGELIST
   DECLARE FUNCTION EditLabel (BYVAL hItem AS HTREEITEM) AS HWND
   DECLARE FUNCTION EndEditLabelNow (BYVAL fCancel AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION EnsureVisible (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION GetBkColor () AS COLORREF
   DECLARE FUNCTION GetCount () AS UINT
   DECLARE FUNCTION GetEditControl () AS HWND
   DECLARE FUNCTION GetExtendedStyle () AS DWORD
   DECLARE FUNCTION GetImageList (BYVAL iImage AS LONG) AS HIMAGELIST
   DECLARE FUNCTION GetIndent () AS UINT
   DECLARE FUNCTION GetInsertMarkColor () AS COLORREF
   DECLARE FUNCTION GetItemHeight () AS LONG
   DECLARE FUNCTION GetItemPartRect (BYVAL hItem AS HTREEITEM, BYREF rc AS RECT, BYVAL partid AS TVITEMPART = TVGIPR_BUTTON) AS BOOLEAN
   DECLARE FUNCTION GetItemPartRect (BYVAL hItem AS HTREEITEM, BYVAL partid AS TVITEMPART = TVGIPR_BUTTON) AS RECT
   DECLARE FUNCTION GetItemRect (BYVAL hItem AS HTREEITEM, BYREF rc AS RECT, BYVAL code AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION GetItemRect (BYVAL hItem AS HTREEITEM, BYVAL code AS BOOLEAN) AS RECT
   DECLARE FUNCTION GetLineColor () AS COLORREF
   DECLARE FUNCTION GetScrollTime () AS UINT
   DECLARE FUNCTION GetSelectedCount () AS DWORD
   DECLARE FUNCTION GetTextColor () AS COLORREF
   DECLARE FUNCTION GetToolTips () AS HWND
   DECLARE FUNCTION GetVisibleCount () AS UINT
   DECLARE FUNCTION HitTest (BYREF ht AS TVHITTESTINFO) AS HTREEITEM
   DECLARE FUNCTION MapAccIDToHTREEITEM (BYVAL id AS UINT) AS HTREEITEM
   DECLARE FUNCTION MapHTREEITEMToAccID (BYVAL hItem AS HTREEITEM) AS UINT
   DECLARE FUNCTION Select_ (BYVAL hItem AS HTREEITEM, BYVAL code AS UINT) AS BOOLEAN
   DECLARE FUNCTION SelectDropTarget (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION SelectItem (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION SelectSetFirstVisible (BYVAL hitem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION UnSelect (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   DECLARE FUNCTION SetAutoScrollInfo (BYVAL uPixPerSec AS UINT, BYVAL uUpdateTime AS UINT) AS LRESULT
   DECLARE FUNCTION SetBkColor (BYVAL clr AS COLORREF) AS COLORREF
   DECLARE SUB SetCheckState (BYVAL hItem AS HTREEITEM, BYVAL fCheck AS BOOLEAN)
   DECLARE SUB SetCheck (BYVAL hItem AS HTREEITEM, BYVAL fCheck AS BOOLEAN)
   DECLARE FUNCTION SetExtendedStyle (BYVAL dwExStyle AS DWORD, BYVAL mask AS UINT) AS HRESULT
   DECLARE FUNCTION SetImageList (BYVAL himl AS HIMAGELIST, BYVAL iImage AS LONG) AS HIMAGELIST
   DECLARE FUNCTION SetIndent (BYVAL indent AS LONG) AS BOOLEAN
   DECLARE FUNCTION SetInsertMark (BYVAL hItem AS HTREEITEM, BYVAL fAfter AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION SetInsertMarkColor (BYVAL clr AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetItemHeight (BYVAL cyItem AS SHORT) AS LONG
   DECLARE FUNCTION SetLineColor (BYVAL clrLine AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetScrollTime (BYVAL uMaxScrollTime AS UINT) AS UINT
   DECLARE FUNCTION SetTextColor (BYVAL clrText AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetToolTips (BYVAL hwndTT AS HWND) AS HWND
   DECLARE FUNCTION ShowInfoTip (BYVAL hitem AS HTREEITEM) AS DWORD
   DECLARE FUNCTION SortChildren (BYVAL hitem AS HTREEITEM, BYVAL fRecurse AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION SortChildrenCB (BYVAL psort AS TVSORTCB PTR, BYVAL fRecurse AS BOOLEAN) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CTreeView constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTreeView (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTreeView (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTreeView (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTreeView (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTreeView destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTreeView
END DESTRUCTOR
' ========================================================================================

' ########################################################################################
'                                   *** TREEVIEW ***
' ########################################################################################

' ========================================================================================
' Retrieves some or all of a treeview item's attributes.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItem (BYVAL pitem AS TVITEMW PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_GETITEMW, 0, cast(LPARAM, pitem))
END FUNCTION
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemEx (BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_GETITEMW, 0, cast(LPARAM, pitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CtreeView.InsertItem (BYVAL pis AS TVINSERTSTRUCTW PTR) AS HTREEITEM
   RETURN CAST(HTREEITEM, SendMessageW(m_hCtl, TVM_INSERTITEMW, 0, CAST(LPARAM, CAST(TVINSERTSTRUCTW PTR, @pis))))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.InsertItem (BYVAL hParent AS HTREEITEM, BYVAL hInsertAfter AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
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
   RETURN CAST(HTREEITEM, SendMessageW(m_hCtl, TVM_INSERTITEMW, 0, CAST(..LPARAM, CAST(TVINSERTSTRUCTW PTR, @tvinsert))))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.Insert (BYVAL hParent AS HTREEITEM, BYVAL hInsertAfter AS HTREEITEM, _
BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0, BYVAL pwszText AS WSTRING PTR) AS HTREEITEM
   DIM tvinsert AS TVINSERTSTRUCTW
   tvinsert.hParent             = hParent
   tvinsert.hInsertAfter        = hInsertAfter
   tvinsert.Item.iImage         = iImage
   tvinsert.Item.iSelectedImage = iSelectedImage
   tvinsert.Item.lParam         = 0
   tvinsert.Item.mask           = TVIF_TEXT OR TVIF_IMAGE OR TVIF_SELECTEDIMAGE OR TVIF_PARAM
   tvinsert.Item.pszText        = pwszText
   IF pwszText <> LPSTR_TEXTCALLBACK THEN
      tvinsert.Item.cchTextMax  = LEN(*pwszText)
   END IF
   RETURN CAST(HTREEITEM, SendMessageW(m_hCtl, TVM_INSERTITEMW, 0, CAST(..LPARAM, CAST(TVINSERTSTRUCTW PTR, @tvinsert))))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts the item at the end of the list in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.AppendItem (BYVAL hParent AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   RETURN this.InsertItem(hParent, TVI_LAST, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts the item at the beginning of the list in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.InsertFirstItem (BYVAL hParent AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   RETURN this.InsertItem(hParent, TVI_FIRST, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Add the item as the root item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.AddRootItem (BYVAL pwszText AS WSTRING PTR, _
BYVAL lParam AS ..LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   RETURN this.InsertItem(NULL, TVI_ROOT, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Inserts the item into the list in alphabetical order.
' ========================================================================================
PRIVATE FUNCTION CTreeView.InsertSortedItem (BYVAL hParent AS HTREEITEM, _
BYVAL pwszText AS WSTRING PTR, BYVAL lParam AS LPARAM = 0, BYVAL iImage AS LONG = 0, BYVAL iSelectedImage AS LONG = 0) AS HTREEITEM
   RETURN this.InsertItem(hParent, TVI_SORT, pwszText, lParam, iImage, iSelectedImage)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an item and all its children from a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.DeleteItem (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_DELETEITEM, 0, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.Delete_ (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_DELETEITEM, 0, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Deletes all items from a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.DeleteAllItems () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_DELETEITEM, 0, CAST(LPARAM, TVI_ROOT))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Deletes all items from a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.Reset () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_DELETEITEM, 0, CAST(LPARAM, TVI_ROOT))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Expands or collapses the list of child items associated with the specified parent item, if any.
' ========================================================================================
PRIVATE FUNCTION CTreeView.Expand (BYVAL hItem AS HTREEITEM, BYVAL code AS UINT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_EXPAND, code, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Expands the list
' ========================================================================================
PRIVATE FUNCTION CTreeView.Expand (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_EXPAND, TVE_EXPAND, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Partially expands the list
' ========================================================================================
PRIVATE FUNCTION CTreeView.ExpandPartial (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_EXPAND, TVE_EXPAND OR TVE_EXPANDPARTIAL, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Collapses the list
' ========================================================================================
PRIVATE FUNCTION CTreeView.Collapse (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_EXPAND, TVE_COLLAPSE, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Collapses the list and removes the child items.
' ========================================================================================
PRIVATE FUNCTION CTreeView.CollapseReset (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_EXPAND, TVE_COLLAPSE OR TVE_COLLAPSERESET, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Collapses the list if it is expanded or expands it if it is collapsed.
' ========================================================================================
PRIVATE FUNCTION CTreeView.Toggle (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_EXPAND, TVE_TOGGLE, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Expands the list of child items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB CTreeView.ExpandChildItems (BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Expand the node
      this.Expand(hNode, TVM_EXPAND)
      ' // Get the first child node
      hChildNode = this.GetChild(hNode)
      IF hChildNode THEN this.ExpandChildItems(hChildNode)
      hNode = NULL
   LOOP
END SUB
' ========================================================================================
' ========================================================================================
' Expands the list of child items, if any, associated with the specified parent node and
' the sibling nodes below it.
' ========================================================================================
PRIVATE SUB CTreeView.ExpandAllChildItems (BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Expand the node
      this.Expand(hNode, TVM_EXPAND)
      ' // Get the first child node
      hChildNode = this.GetChild(hNode)
      IF hChildNode THEN this.ExpandAllChildItems(hChildNode)
      ' // Get the next sibling item
      hNode = this.GetNextSibling(hNode)
   LOOP
END SUB
' ========================================================================================
' ========================================================================================
' Expands the list of sibling items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB CTreeView.ExpandSiblings (BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Expand the node
      this.Expand(hNode, TVM_EXPAND)
      ' // Get the next sibling item
      hNode = this.GetNextSibling(hNode)
   LOOP
END SUB
' ========================================================================================
' ========================================================================================
' Expands all the list of items.
' Note: This is a very slow process if the control has many items because of the continuous
' redrawing; therefore, disable redrawing before calling it with
'   SendMessageW(hTreeView, WM_SETREDRAW, FALSE, 0)
' and reenable drawing once it has ended with
'   SendMessageW(hTreeView, WM_SETREDRAW, TRUE, 0)
' followed by a redrawing of the control
'   InvalidateRect hTreeView, NULL, CTRUE
'   UpdateWindow hTreeView
' ========================================================================================
PRIVATE SUB CTreeView.ExpandAllItems
   DIM hNode AS HTREEITEM = this.GetRoot
   IF hNode THEN this.ExpandAllChildItems(hNode)
END SUB
' ========================================================================================
' ========================================================================================
' Collapses the list of child items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB CTreeView.CollapseChildItems (BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Collapse the node
      this.Expand(hNode, TVE_COLLAPSE)
      ' // Get the first child node
      hChildNode = this.GetChild(hNode)
      IF hChildNode THEN this.CollapseChildItems(hChildNode)
      hNode = NULL
   LOOP
END SUB
' ========================================================================================
' ========================================================================================
' Collapses the list of sibling items, if any, associated with the specified parent node.
' ========================================================================================
PRIVATE SUB CTreeView.CollapseSiblings (BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Collapse the node
      this.Expand(hNode, TVE_COLLAPSE)
      ' // Get the next sibling item
      hNode = this.GetNextSibling(hNode)
   LOOP
END SUB
' ========================================================================================
' ========================================================================================
' Collapses the list of child items, if any, associated with the specified parent node and
' the sibling nodes below it.
' ========================================================================================
PRIVATE SUB CTreeView.CollapseAllChildItems (BYVAL hNode AS HTREEITEM)
   DIM hChildNode AS HTREEITEM
   DO WHILE hNode
      ' // Collapse the node
      this.Expand(hNode, TVE_COLLAPSE)
      ' // Get the first child node
      hChildNode = this.GetChild(hNode)
      IF hChildNode THEN this.CollapseAllChildItems(hChildNode)
      ' // Get the next sibling item
      hNode = this.GetNextSibling(hNode)
   LOOP
END SUB
' ========================================================================================
' ========================================================================================
' Collapses all the list of items.
' ========================================================================================
PRIVATE SUB CTreeView.CollapseAllItems 
   DIM hNode AS HTREEITEM = this.GetRoot
   IF hNode THEN this.CollapseAllChildItems(hNode)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves some or all of a treeview item's state attributes.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemState (BYVAL hItem AS HTREEITEM, BYVAL mask AS UINT) AS UINT
   RETURN SendMessageW(m_hCtl, TVM_GETITEMSTATE, CAST(WPARAM, hItem), CAST(LPARAM, mask))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns TRUE if the bold attribute is set, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetBold (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN (this.GetItemState(hItem, TVIS_BOLD) AND TVIS_BOLD) = TVIS_BOLD
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves some or all of a treeview item's state attributes.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetCheckState (BYVAL hItem AS HTREEITEM) AS UINT
   RETURN SendMessageW(m_hCtl, TVM_GETITEMSTATE, CAST(WPARAM, hItem), CAST(LPARAM, TVIS_STATEIMAGEMASK)) SHR 12
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns TRUE if the item is checked, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.IsItemChecked (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN this.GetCheckState(hItem) = 1
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns TRUE if the item is checked, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetCheck (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN this.GetCheckState(hItem) = 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of characters in the incremental search string.
' If the treeview control is not in incremental search mode, the return value is zero.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetISearchStringLen () AS LONG
   RETURN SendMessageW(m_hCtl, TVM_GETISEARCHSTRING, 0, CAST(LPARAM, NULL))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the incremental search string for a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetISearchString (BYVAL pwsz AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_GETISEARCHSTRING, 0, CAST(LPARAM, pwsz))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetISearchString () AS DWSTRING
   DIM nLen AS LONG, dwsSearch AS DWSTRING
   nLen = SendMessageW(m_hCtl, TVM_GETISEARCHSTRINGW, 0, NULL)
   dwsSearch = WSPACE(nLen + 1)
   nLen = SendMessageW(m_hCtl, TVM_GETISEARCHSTRINGW, 0, CAST(LPARAM, STRPTR(dwsSearch)))
   RETURN LEFT(dwsSearch, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieve the lParam value from a Treeview
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetlParam (BYVAL hItem AS HTREEITEM) AS LPARAM
   DIM tvi AS TVITEMW
   tvi.hItem = hItem
   tvi.mask  = TVIF_HANDLE OR TVIF_PARAM
   SendMessageW(m_hCtl, TVM_GETITEMW, 0, CAST(LPARAM, @tvi))
   RETURN tvi.lParam
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a treeview item's state attributes.
' ========================================================================================
PRIVATE SUB CTreeView.SetItemState (BYVAL hItem AS HTREEITEM, BYVAL nState AS UINT, BYVAL statemask AS UINT)
   DIM tvi AS TVITEMW
   tvi.mask  = TVIF_STATE
   tvi.hItem = hItem
   tvi.stateMask = statemask
   tvi.state = nState
   SendMessageW m_hCtl, TVM_SETITEMW, 0, CAST(LPARAM, @tvi)
END SUB
' ========================================================================================

' ========================================================================================
' Sets/unsets the bold attribute of the specified item.
' ========================================================================================
PRIVATE SUB CTreeView.SetBold (BYVAL hItem AS HTREEITEM, BYVAL fBold AS BOOLEAN)
   this.SetItemState(hItem, IIF(fBold = 0, 0, TVIS_BOLD), TVIS_BOLD)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the text of the specified item.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchtextmax AS LONG) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM tvi AS TVITEMW
   tvi.hitem = hItem
   tvi.mask = TVIF_TEXT
   tvi.psztext = pwszText
   tvi.cchtextmax = cchtextmax
   RETURN SendMessageW(m_hCtl, TVM_GETITEMW, 0, cast(LPARAM, cast(TVITEMW PTR, @tvi)))
END FUNCTION
' ========================================================================================
' Retrieves the text of the specified item.
' Returns the retrieved text.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemText (BYVAL hItem AS HTREEITEM, BYVAL cchTextMax AS LONG = 0) AS DWSTRING
   DIM tvi AS TVITEMW, dwsText AS DWSTRING, p AS LONG
   IF cchtextmax < 1 THEN cchtextmax = MAX_PATH + 1
   dwsText = WSTRING(cchtextmax + 1, WCHR(0))
   tvi.hitem = hItem
   tvi.mask = TVIF_TEXT
   tvi.psztext = STRPTR(dwsText)
   tvi.cchtextmax = cchtextmax + 1
   IF SendMessageW(m_hCtl, TVM_GETITEMW, 0, CAST(LPARAM, @tvi)) = FALSE THEN EXIT FUNCTION
   p = INSTR(dwsText, CHR(0))
   IF p THEN dwsText = LEFT(dwsText, p - 1)
   FUNCTION = dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets some or all of a treeview item's attributes.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetItem (BYVAL pitem AS TVITEMW PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SETITEMW, 0, cast(WPARAM, pitem))
END FUNCTION
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetItemEx (BYVAL pitem AS TVITEMEXW PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SETITEMW, 0, cast(WPARAM, pitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetItemText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM tvi AS TVITEMW
   tvi.hitem = hItem
   tvi.mask = TVIF_TEXT
   tvi.psztext = pwszText
   tvi.cchtextmax = LEN(*pwszText)
   RETURN SendMessageW(m_hCtl, TVM_SETITEMW, 0, cast(LPARAM, @tvi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the specified item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DIM tvi AS TVITEMW
   tvi.hitem = hItem
   tvi.mask = TVIF_TEXT
   tvi.psztext = pwszText
   tvi.cchtextmax = LEN(*pwszText)
   RETURN SendMessageW(m_hCtl, TVM_SETITEMW, 0, cast(LPARAM, @tvi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the treeview item that bears the specified relationship to a specified item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetNextItem (BYVAL hItem AS HTREEITEM, BYVAL code AS UINT) AS HTREEITEM
   RETURN CAST(HTREEITEM, SendMessageW(m_hCtl, TVM_GETNEXTITEM, code, CAST(LPARAM, hItem)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the first child item of the specified treeview item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetChild (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_CHILD)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the currently selected item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetSelection () AS HTREEITEM
   RETURN this.GetNextItem(NULL, TVGN_CARET)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the currently selected item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetSelect () AS HTREEITEM
   RETURN this.GetNextItem(NULL, TVGN_CARET)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the treeview item that is the target of a drag-and-drop operation.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetDropHilight () AS HTREEITEM
   RETURN this.GetNextItem(NULL, TVGN_DROPHILITE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the first visible item in a treeview control window.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetFirstVisible () AS HTREEITEM
   RETURN this.GetNextItem(NULL, TVGN_FIRSTVISIBLE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the next sibling item of a specified item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetNextSibling (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_NEXT)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the next sibling item of a specified item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetNext (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_NEXT)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the treeview item that bears the TVGN_NEXTSELECTED relationship to a specified tree item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetNextSelected (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_NEXTSELECTED)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the next visible item that follows a specified item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetNextVisible (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_NEXTVISIBLE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the parent item of the specified treeview item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetParent (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_PARENT)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the text of the parent item of the specified treeview item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetParentText (BYVAL hItem AS HTREEITEM, BYVAL cchTextMax AS LONG = 0) AS DWSTRING
   DIM hNode AS HTREEITEM = this.GetNextItem(hItem, TVGN_PARENT)
   IF hNode THEN RETURN this.GetItemText(hNode, cchTextMax)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the topmost or very first item of the treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetRoot () AS HTREEITEM
   RETURN this.GetNextItem(NULL, TVGN_ROOT)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the previous sibling item of a specified item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetPrevSibling (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_PREVIOUS)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetPrevious (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_PREVIOUS)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the first visible item that precedes a specified item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetPrevVisible (BYVAL hItem AS HTREEITEM) AS HTREEITEM
   RETURN this.GetNextItem(hItem, TVGN_PREVIOUSVISIBLE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves some or all of a treeview item's state attributes.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetLastVisible () AS HTREEITEM
   RETURN this.GetNextItem(NULL, TVGN_LASTVISIBLE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text of the parent item of the specified treeview item.
' Usage example:
' DIM wszText AS WSTRING * 260
' CTreeView.GetParentText(hTreeView, hItem, @wszText, 260)
' AfxMsg(wszText)
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetParentText (BYVAL hItem AS HTREEITEM, BYVAL pwszText AS WSTRING PTR, BYVAL cchTextMax AS LONG) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM hNode AS HTREEITEM = this.GetNextItem(hItem, TVGN_PARENT)
   IF hNode THEN RETURN this.GetItemText(hNode, pwszText, cchTextMax)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for an exact string in a tree view node. Search is case sensitive.
' Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CTreeView.ItemExists (BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM cbLen AS LONG = LEN(*pwszText)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM dwsItemText AS DWSTRING = WSPACE(cbLen + 1)
   DIM hItem AS HTREEITEM = this.GetChild(hNode)
   DO
      IF hItem = NULL THEN EXIT FUNCTION
      IF this.GetItemText(hItem, *dwsItemText, LEN(dwsItemText)) = FALSE THEN EXIT DO
      IF dwsItemText = *pwszText THEN
         FUNCTION = TRUE
         EXIT DO
      END IF
      hItem = this.GetNextSibling(hItem)
   LOOP
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches for an exact string in a tree view node. Search is case insensitive.
' Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CTreeView.ItemExistsIgnoreCase (BYVAL hNode AS HTREEITEM, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   IF pwszText = NULL THEN EXIT FUNCTION
   DIM cbLen AS LONG = LEN(*pwszText)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM dwsItemText AS DWSTRING = WSPACE(cbLen + 1)
   DIM hItem AS HTREEITEM = this.GetChild(hNode)
   DO
      IF hItem = NULL THEN EXIT FUNCTION
      IF this.GetItemText(hItem, *dwsItemText, LEN(dwsItemText)) = FALSE THEN EXIT DO
      IF UCASE(dwsItemText) = UCASE(*pwszText) THEN
         FUNCTION = TRUE
         EXIT DO
      END IF
      hItem = this.GetNextSibling(hItem)
   LOOP
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a dragging bitmap for the specified item in a treeview control. The macro also
' creates an image list for the bitmap and adds the bitmap to the image list.
' ========================================================================================
PRIVATE FUNCTION CTreeView.CreateDragImage ( BYVAL hItem AS HTREEITEM) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, TVM_CREATEDRAGIMAGE, 0, CAST(LPARAM, hItem)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Begins in-place editing of the specified item's text, replacing the text of the item
' with a single-line edit control containing the text.
' ========================================================================================
PRIVATE FUNCTION CTreeView.EditLabel (BYVAL hItem AS HTREEITEM) AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TVM_EDITLABEL, 0, CAST(LPARAM, hItem)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Ends the editing of a treeview item's label.
' ========================================================================================
PRIVATE FUNCTION CTreeView.EndEditLabelNow (BYVAL fCancel AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_ENDEDITLABELNOW, fCancel, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Ensures that a treeview item is visible, expanding the parent item or scrolling the
' treeview control, if necessary.
' ========================================================================================
PRIVATE FUNCTION CTreeView.EnsureVisible (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_ENSUREVISIBLE, 0, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current background color of the control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetBkColor () AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a count of the items in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetCount () AS UINT
   RETURN SendMessageW(m_hCtl, TVM_GETCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the edit control being used to edit a treeview item's text.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetEditControl () AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TVM_GETEDITCONTROL, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the extended style for a specified treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetExtendedStyle () AS DWORD
   RETURN SendMessageW(m_hCtl, TVM_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the normal or state image list associated with a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetImageList (BYVAL iImage AS LONG) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, TVM_GETIMAGELIST, iImage, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the amount, in pixels, that child items are indented relative to their parent items.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetIndent () AS UINT
   RETURN SendMessageW(m_hCtl, TVM_GETINDENT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the color used to draw the insertion mark for the tree view.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetInsertMarkColor () AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_GETINSERTMARKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current height of the treeview items.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemHeight () AS LONG
   RETURN SendMessageW(m_hCtl, TVM_GETITEMHEIGHT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the largest possible bounding rectangle that constitutes the "hit zone" for a
' specified part of an item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemPartRect (BYVAL hItem AS HTREEITEM, BYREF rc AS RECT, BYVAL partid AS TVITEMPART = TVGIPR_BUTTON) AS BOOLEAN
   DIM info AS TVGETITEMPARTRECTINFO
   info.hti = hitem
   info.prc = VARPTR(rc)
   info.partID = partid
   RETURN SendMessageW(m_hCtl, TVM_GETITEMPARTRECT, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemPartRect (BYVAL hItem AS HTREEITEM, BYVAL partid AS TVITEMPART = TVGIPR_BUTTON) AS RECT
   DIM info AS TVGETITEMPARTRECTINFO, rc AS RECT
   info.hti = hitem
   info.prc = VARPTR(rc)
   info.partID = partid
   SendMessageW(m_hCtl, TVM_GETITEMPARTRECT, 0, CAST(LPARAM, @info))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle for a treeview item and indicates whether the item is visible.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemRect (BYVAL hItem AS HTREEITEM, BYREF rc AS RECT, BYVAL code AS BOOLEAN) AS BOOLEAN
   IF VARPTR(rc) = 0 THEN RETURN FALSE
   *CPTR(HTREEITEM PTR, @rc) = hitem
   RETURN SendMessageW(m_hCtl, TVM_GETITEMRECT, code, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetItemRect (BYVAL hItem AS HTREEITEM, BYVAL code AS BOOLEAN) AS RECT
   DIM rc AS RECT
   *CPTR(HTREEITEM PTR, @rc) = hItem
   SendMessageW(m_hCtl, TVM_GETITEMRECT, code, CAST(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current line color.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetLineColor () AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_GETLINECOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the maximum scroll time for the treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetScrollTime () AS UINT
   RETURN SendMessageW(m_hCtl, TVM_GETSCROLLTIME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of items selected.
' This message is not implemented.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetSelectedCount () AS DWORD
   RETURN SendMessageW(m_hCtl, TVM_GETSELECTEDCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current text color of the control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetTextColor () AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_GETTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the child ToolTip control used by a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.GetToolTips () AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TVM_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Obtains the number of items that can be fully visible in the client window of a
' treeview control.
' ========================================================================================
FUNCTION CTreeView.GetVisibleCount () AS UINT
   RETURN SendMessageW(m_hCtl, TVM_GETVISIBLECOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines the location of the specified point relative to the client area of a
' treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.HitTest (BYREF ht AS TVHITTESTINFO) AS HTREEITEM
   RETURN CAST(HTREEITEM, SendMessageW(m_hCtl, TVM_HITTEST, 0, CAST(LPARAM, @ht)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps an accessibility ID to an HTREEITEM.
' ========================================================================================
PRIVATE FUNCTION CTreeView.MapAccIDToHTREEITEM (BYVAL id AS UINT) AS HTREEITEM
   RETURN CAST(HTREEITEM, SendMessageW(m_hCtl, TVM_MAPACCIDTOHTREEITEM, id, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps an HTREEITEM to an accessibility ID.
' ========================================================================================
PRIVATE FUNCTION CTreeView.MapHTREEITEMToAccID (BYVAL hItem AS HTREEITEM) AS UINT
   RETURN SendMessageW(m_hCtl, TVM_MAPHTREEITEMTOACCID, CAST(WPARAM, hItem), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects the specified treeview item, scrolls the item into view, or redraws the item in
' the style used to indicate the target of a drag-and-drop operation.
' ========================================================================================
PRIVATE FUNCTION CTreeView.Select_ (BYVAL hItem AS HTREEITEM, BYVAL code AS UINT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SELECTITEM, code, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Unselects all items.
' ========================================================================================
PRIVATE FUNCTION CTreeView.UnSelect (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SELECTITEM, 0, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redraws a specified treeview control item in the style used to indicate the target of a
' drag-and-drop operation.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SelectDropTarget (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SELECTITEM, TVGN_DROPHILITE, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects the specified treeview item.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SelectItem (BYVAL hItem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SELECTITEM, TVGN_CARET, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the treeview control vertically to ensure that the specified item is visible.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SelectSetFirstVisible (BYVAL hitem AS HTREEITEM) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets information used to determine auto-scroll characteristics.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetAutoScrollInfo (BYVAL uPixPerSec AS UINT, BYVAL uUpdateTime AS UINT) AS LRESULT
   RETURN SendMessageW(m_hCtl, TVM_SETAUTOSCROLLINFO, uPixPerSec, uUpdateTime)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of the control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetBkColor (BYVAL clr AS COLORREF) AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_SETBKCOLOR, 0, clr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the item's state image to "checked" or "unchecked."
' ========================================================================================
PRIVATE SUB CTreeView.SetCheckState (BYVAL hItem AS HTREEITEM, BYVAL fCheck AS BOOLEAN)
   DIM ix AS LONG = IIF(fCheck = TRUE, IndexToStateImageMask(2), IndexToStateImageMask(1))
   this.SetItemState hItem, ix, TVIS_STATEIMAGEMASK
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CTreeView.SetCheck (BYVAL hItem AS HTREEITEM, BYVAL fCheck AS BOOLEAN)
   DIM ix AS LONG = IIF(fCheck = TRUE, IndexToStateImageMask(2), IndexToStateImageMask(1))
   this.SetItemState hItem, ix, TVIS_STATEIMAGEMASK
END SUB
' ========================================================================================

' ========================================================================================
' Sets the extended style for a specified TreeView control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetExtendedStyle (BYVAL dwExStyle AS DWORD, BYVAL mask AS UINT) AS HRESULT
   RETURN SendMessageW(m_hCtl, TVM_SETEXTENDEDSTYLE, mask, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' [Intended for internal use; not recommended for use in applications.
' This macro may not be supported in future versions of Windows.]
' ========================================================================================
'PRIVATE FUNCTION CTreeView.SetHot (BYVAL hItem AS HTREEITEM) AS LRESULT
'   RETURN SendMessageW(m_hCtl, TVM_SETHOT, 0, hItem)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the normal or state image list for a treeview control and redraws the control
' using the new images.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetImageList (BYVAL himl AS HIMAGELIST, BYVAL iImage AS LONG) AS HIMAGELIST
   RETURN CAST(HIMAGELIST, SendMessageW(m_hCtl, TVM_SETIMAGELIST, iImage, CAST(LPARAM, himl)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the width of indentation for a treeview control and redraws the control to reflect
' the new width.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetIndent (BYVAL indent AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SETINDENT, indent, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the insertion mark in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetInsertMark (BYVAL hItem AS HTREEITEM, BYVAL fAfter AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SETINSERTMARK, fAfter, CAST(LPARAM, hItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the color used to draw the insertion mark for the tree view.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetInsertMarkColor (BYVAL clr AS COLORREF) AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_SETINSERTMARKCOLOR, 0, clr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the height of the treeview items.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetItemHeight (BYVAL cyItem AS SHORT) AS LONG
   RETURN SendMessageW(m_hCtl, TVM_SETITEMHEIGHT, cyItem, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current line color.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetLineColor (BYVAL clrLine AS COLORREF) AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_SETLINECOLOR, 0, clrLine)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the maximum scroll time for the treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetScrollTime (BYVAL uMaxScrollTime AS UINT) AS UINT
   RETURN SendMessageW(m_hCtl, TVM_SETSCROLLTIME, uMaxScrollTime, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text color of the control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetTextColor (BYVAL clrText AS COLORREF) AS COLORREF
   RETURN SendMessageW(m_hCtl, TVM_SETTEXTCOLOR, 0, clrText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a treeview control's child ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SetToolTips (BYVAL hwndTT AS HWND) AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TVM_SETTOOLTIPS, CAST(WPARAM, hwndTT), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' [Intended for internal use; not recommended for use in applications.
' This macro may not be supported in future versions of Windows.]
' ========================================================================================
'PRIVATE FUNCTION CTreeView.SetBorder (BYVAL dwFlags AS DWORD, BYVAL xBorder AS SHORT, BYVAL yBorder AS SHORT) AS LONG
'   RETURN SendMessageW(m_hCtl, TVM_SETBORDER, dwFlags, MAKELONG(xBorder, yBorder))
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows the infotip for a specified item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.ShowInfoTip (BYVAL hitem AS HTREEITEM) AS DWORD
   RETURN SendMessageW(m_hCtl, TVM_SHOWINFOTIP, 0, CAST(LPARAM, hitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sorts the child items of the specified parent item in a treeview control.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SortChildren (BYVAL hitem AS HTREEITEM, BYVAL fRecurse AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SORTCHILDREN, fRecurse, CAST(LPARAM, hitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sorts treeview items using an application-defined callback function that compares the items.
' ========================================================================================
PRIVATE FUNCTION CTreeView.SortChildrenCB (BYVAL psort AS TVSORTCB PTR, BYVAL fRecurse AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TVM_SORTCHILDRENCB, fRecurse, CAST(LPARAM, psort))
END FUNCTION
' ========================================================================================

END NAMESPACE
