' ########################################################################################
' Platform: Microsoft Windows
' Filename: CRebar.inc
' Purpose:  Windows common controls - Rebar control
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
' CRebar class
' ########################################################################################

TYPE CRebar EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE SUB BeginDrag (BYVAL uBand AS DWORD, BYVAL dwPos AS DWORD)
   DECLARE FUNCTION DeleteBand (BYVAL uBand AS DWORD) AS BOOLEAN
   DECLARE SUB DragMove (BYVAL dwPos AS DWORD)
   DECLARE SUB EndDrag ()
   DECLARE SUB GetBandBorders (BYVAL uBand AS DWORD, BYREF rc AS RECT)
   DECLARE FUNCTION GetBandBorders (BYVAL uBand AS DWORD) AS RECT
   DECLARE FUNCTION GetBandCount () AS DWORD
   DECLARE FUNCTION GetBandInfo (BYVAL uBand AS DWORD, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
   DECLARE FUNCTION GetBandInfo (BYVAL uBand AS DWORD) AS REBARBANDINFOW
   DECLARE SUB GetBandMargins (BYVAL pMargins AS MARGINS PTR)
   DECLARE FUNCTION GetBarHeight () AS DWORD
   DECLARE FUNCTION GetBarInfo (BYREF rbi AS REBARINFO) AS BOOLEAN
   DECLARE FUNCTION GetBkColor () AS COLORREF
   DECLARE FUNCTION GetColorScheme (BYREF cs AS COLORSCHEME) AS BOOLEAN
   DECLARE FUNCTION GetColorScheme () AS COLORSCHEME
   DECLARE FUNCTION GetDropTarget () AS IDropTarget PTR
   DECLARE FUNCTION GetExtendedStyle () AS DWORD
   DECLARE FUNCTION GetPalette () AS DWORD
   DECLARE FUNCTION GetRect (BYVAL uBand AS DWORD, BYREF rc AS RECT) AS BOOLEAN
   DECLARE FUNCTION GetRect (BYVAL uBand AS DWORD) AS RECT
   DECLARE FUNCTION GetRowCount () AS DWORD
   DECLARE FUNCTION GetRowHeight (BYVAL uRow AS DWORD) AS DWORD
   DECLARE FUNCTION GetTextColor () AS COLORREF
   DECLARE FUNCTION GetTooltips () AS HWND
   DECLARE FUNCTION HitTest (BYREF rhbt AS RBHITTESTINFO) AS LONG
   DECLARE FUNCTION IdToIndex (BYVAL uBandID AS DWORD) AS LONG
   DECLARE FUNCTION InsertBand (BYVAL nIndex AS LONG, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
   DECLARE SUB MaximizeBand (BYVAL uBand AS DWORD, BYVAL fIdeal AS BOOLEAN)
   DECLARE SUB MinimizeBand (BYVAL uBand AS DWORD)
   DECLARE FUNCTION MoveBand (BYVAL nFrom AS DWORD, BYVAL nTo AS DWORD) AS BOOLEAN
   DECLARE SUB PushChevron (BYVAL uBand AS DWORD, BYVAL iAppValue AS DWORD)
   DECLARE FUNCTION SetBandInfo (BYVAL uBand AS DWORD, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
   #if _WIN32_WINNT = &h0602
   DECLARE FUNCTION SetBandWidth (BYVAL uBand AS DWORD, BYVAL nWidth AS DWORD) AS BOOLEAN
   #endif
   DECLARE FUNCTION SetBarInfo (BYREF rbi AS REBARINFO) AS BOOLEAN
   DECLARE FUNCTION SetBkColor (BYVAL clrBk AS COLORREF) AS COLORREF
   DECLARE SUB SetColorScheme (BYREF cs AS COLORSCHEME)
   DECLARE FUNCTION SetPalette (BYVAL hpal AS HPALETTE) AS HPALETTE
   DECLARE FUNCTION SetParent (BYVAL hwndParent AS hWND) AS HWND
   DECLARE FUNCTION SetTextColor (BYVAL clrText AS COLORREF) AS COLORREF
   DECLARE SUB SetTooltips (BYVAL hwndToolTip AS HWND)
   DECLARE SUB SetWindowTheme (BYVAL pwszTheme AS WSTRING PTR)
   DECLARE FUNCTION ShowBand (BYVAL uBand AS DWORD, BYVAL fShow AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION SizeToRect (BYVAL prc AS RECT PTR) AS BOOLEAN
   DECLARE FUNCTION SizeToRect () AS RECT

END TYPE
' ========================================================================================

' ========================================================================================
' * CRebar constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CRebar (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CRebar (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CRebar (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CRebar (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CRebar destructor
' ========================================================================================
PRIVATE DESTRUCTOR CRebar
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                      *** REBAR ***
' ########################################################################################

' ========================================================================================
' Puts the rebar control in drag-and-drop mode.
' ========================================================================================
PRIVATE SUB CRebar.BeginDrag (BYVAL uBand AS DWORD, BYVAL dwPos AS DWORD)
   SendMessageW(m_hCtl, RB_BEGINDRAG, uBand, dwPos)
END SUB
' ========================================================================================

' ========================================================================================
' Deletes a band from a rebar control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.DeleteBand (BYVAL uBand AS DWORD) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, RB_DELETEBAND, uBand, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the drag position in the rebar control after a previous RB_BEGINDRAG message.
' ========================================================================================
PRIVATE SUB CRebar.DragMove (BYVAL dwPos AS DWORD)
   SendMessageW(m_hCtl, RB_DRAGMOVE, 0, dwPos)
END SUB
' ========================================================================================

' ========================================================================================
' Terminates the rebar control's drag-and-drop operation. This message does not cause an
' RBN_ENDDRAG notification to be sent.
' ========================================================================================
PRIVATE SUB CRebar.EndDrag
   SendMessageW(m_hCtl, RB_ENDDRAG, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the borders of a band. The result of this message can be used to calculate the
' usable area in a band.
' ========================================================================================
PRIVATE SUB CRebar.GetBandBorders (BYVAL uBand AS DWORD, BYREF rc AS RECT)
   SendMessageW(m_hCtl, RB_GETBANDBORDERS, uBand, cast(LPARAM, @rc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBandBorders (BYVAL uBand AS DWORD) AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, RB_GETBANDBORDERS, uBand, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the count of bands currently in the rebar control.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBandCount () AS DWORD
   RETURN SendMessageW(m_hCtl, RB_GETBANDCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about a specified band in a rebar control.
' Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBandInfo (BYVAL uBand AS DWORD, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
   IF VARPTR(rbbi) = 0 THEN RETURN FALSE
   IF this.ComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   RETURN SendMessageW(m_hCtl, RB_GETBANDINFO, uBand, cast(LPARAM, @rbbi))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBandInfo (BYVAL uBand AS DWORD) AS REBARBANDINFOW
   DIM rbbi AS REBARBANDINFOW
   IF this.ComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   SendMessageW(m_hCtl, RB_GETBANDINFO, uBand, cast(LPARAM, @rbbi))
   RETURN rbbi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the margins of a band.
' ========================================================================================
PRIVATE SUB CRebar.GetBandMargins (BYVAL pMargins AS MARGINS PTR)
   SendMessageW(m_hCtl, RB_GETBANDMARGINS, 0, cast(LPARAM, pMargins))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the height of the rebar control.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBarHeight () AS DWORD
   RETURN SendMessageW(m_hCtl, RB_GETBARHEIGHT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the rebar control and the image list it uses.
' Return value: Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBarInfo (BYREF rbi AS REBARINFO) AS BOOLEAN
   IF VARPTR(rbi) = 0 THEN RETURN FALSE
   rbi.cbSize = SIZEOF(rbi)
   RETURN SendMessageW(m_hCtl, RB_GETBARINFO, 0, cast(LPARAM, @rbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a rebar control's default background color.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetBkColor () AS COLORREF
   RETURN SendMessageW(m_hCtl, RB_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the color scheme information from the rebar control.
' Return value: Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetColorScheme (BYREF cs AS COLORSCHEME) AS BOOLEAN
   IF VARPTR(cs) THEN RETURN FALSE
   cs.dwSize = SIZEOF(cs)
   RETURN SendMessageW(m_hCtl, RB_GETCOLORSCHEME, 0, CAST(LPARAM, @cs))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRebar.GetColorScheme () AS COLORSCHEME
   DIM cs AS COLORSCHEME
   cs.dwSize = SIZEOF(cs)
   SendMessageW(m_hCtl, RB_GETCOLORSCHEME, 0, CAST(LPARaM, @cs))
   RETURN cs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a rebar control's IDropTarget interface pointer.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetDropTarget () AS IDropTarget PTR
   DIM ppIDropTarget AS IDropTarget PTR
   SendMessageW(m_hCtl, RB_GETDROPTARGET, 0, CAST(LPARAM, @ppIDropTarget))
   RETURN ppIDropTarget
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the extended style.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetExtendedStyle () AS DWORD
   RETURN SendMessageW(m_hCtl, RB_GETEXTENDEDSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the rebar control's current palette.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetPalette () AS DWORD
   RETURN SendMessageW(m_hCtl, RB_GETPALETTE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the bounding rectangle for a given band in a rebar control.
' Return value: Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetRect (BYVAL uBand AS DWORD, BYREF rc AS RECT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, RB_GETRECT, uBand, cast(LPARAM, @rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRebar.GetRect (BYVAL uBand AS DWORD) AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, RB_GETRECT, uBand, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of rows of bands in a rebar control.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetRowCount () AS DWORD
   RETURN SendMessageW(m_hCtl, RB_GETROWCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the height of a specified row in a rebar control.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetRowHeight (BYVAL uRow AS DWORD) AS DWORD
   RETURN SendMessageW(m_hCtl, RB_GETROWHEIGHT, uRow, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a rebar control's default text color.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetTextColor () AS COLORREF
   RETURN SendMessageW(m_hCtl, RB_GETTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to any ToolTip control associated with the rebar control.
' ========================================================================================
PRIVATE FUNCTION CRebar.GetTooltips () AS HWND
   RETURN cast(HWND, SendMessageW(m_hCtl, RB_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which portion of a rebar band is at a given point on the screen, if a rebar
' band exists at that point.
' - m_hCtl: Pointer to the rebar control.
' - lprhbt: Pointer to an RBHITTESTINFO structure. Before sending the message, the pt member
'   of this structure must be initialized to the point that will be hit tested, in client coordinates.
' Returns the zero-based index of the band at the given point, or -1 if no rebar band was at the point.
' ========================================================================================
PRIVATE FUNCTION CRebar.HitTest (BYREF rhbt AS RBHITTESTINFO) AS LONG
   IF VARPTR(rhbt) = 0 THEN RETURN -1
   RETURN SendMessageW(m_hCtl, RB_HITTEST, 0, cast(LPARAM, @rhbt))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a band identifier to a band index in a rebar control.
' Returns the zero-based band index if successful, or -1 otherwise. If duplicate band
' identifiers exist, the first one is returned.
' ========================================================================================
PRIVATE FUNCTION CRebar.IdToIndex (BYVAL uBandID AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, RB_IDTOINDEX, uBandID, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a new band in a rebar control.
' ========================================================================================
PRIVATE FUNCTION CRebar.InsertBand (BYVAL nIndex AS LONG, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
   IF VARPTR(rbbi) = 0 THEN RETURN FALSE
   IF this.ComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   RETURN SendMessageW(m_hCtl, RB_INSERTBAND, nIndex, cast(LPARAM, @rbbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resizes a band in a rebar control to either its ideal or largest size.
' ========================================================================================
PRIVATE SUB CRebar.MaximizeBand (BYVAL uBand AS DWORD, BYVAL fIdeal AS BOOLEAN)
   SendMessageW(m_hCtl, RB_MAXIMIZEBAND, uBand, fIdeal)
END SUB
' ========================================================================================

' ========================================================================================
' Resizes a band in a rebar control to its smallest size.
' ========================================================================================
PRIVATE SUB CRebar.MinimizeBand (BYVAL uBand AS DWORD)
  SendMessageW(m_hCtl, RB_MINIMIZEBAND, cast(WPARAM, uBand), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Moves a band from one index to another. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.MoveBand (BYVAL nFrom AS DWORD, BYVAL nTo AS DWORD) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, RB_MOVEBAND, nFrom, nTo)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sent to a rebar control to programmatically push a chevron.
' ========================================================================================
PRIVATE SUB CRebar.PushChevron (BYVAL uBand AS DWORD, BYVAL iAppValue AS DWORD)
   SendMessageW(m_hCtl, RB_PUSHCHEVRON, uBand, iAppValue)
END SUB
' ========================================================================================

' ========================================================================================
' Sets characteristics of an existing band in a rebar control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.SetBandInfo (BYVAL uBand AS DWORD, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
   IF this.ComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   RETURN SendMessageW(m_hCtl, RB_SETBANDINFO, uBand, cast(LPARAM, @rbbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the characteristics of a rebar control. Returns TRUE or FALSE.
' ========================================================================================
#if _WIN32_WINNT = &h0602
PRIVATE FUNCTION CRebar.SetBandWidth (BYVAL uBand AS DWORD, BYVAL nWidth AS DWORD) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, RB_SETBANDWIDTH, uBand, nWidth)
END FUNCTION
#endif
' ========================================================================================

' ========================================================================================
' Sets the characteristics of a rebar control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.SetBarInfo (BYREF rbi AS REBARINFO) AS BOOLEAN
   IF VARPTR(rbi) = 0 THEN RETURN FALSE
   rbi.cbSize = SIZEOF(rbi)
   RETURN SendMessageW(m_hCtl, RB_SETBARINFO, 0, cast(LPARAM, @rbi))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a rebar control's default background color.
' Returns a COLORREF value that represents the previous default background color.
' ========================================================================================
PRIVATE FUNCTION CRebar.SetBkColor (BYVAL clrBk AS DWORD) AS DWORD
   RETURN SendMessageW(m_hCtl, RB_SETBKCOLOR, 0, clrBk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the color scheme information for the rebar control.
' ========================================================================================
PRIVATE SUB CRebar.SetColorScheme (BYREF cs AS COLORSCHEME)
   IF VARPTR(cs) THEN EXIT SUB
   cs.dwSize = SIZEOF(cs)
   SendMessageW(m_hCtl, RB_SETCOLORSCHEME, 0, CAST(LPARAM, @cs))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the rebar control's current palette.
' Returns an HPALETTE that specifies the rebar control's previous palette.
' ========================================================================================
PRIVATE FUNCTION CRebar.SetPalette (BYVAL hpal AS HPALETTE) AS HPALETTE
   RETURN cast(HPALETTE, SendMessageW(m_hCtl, RB_SETPALETTE, 0, cast(LPARAM, hpal)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a rebar control's parent window.
' Returns the handle to the previous parent window, or NULL if there is no previous parent.
' ========================================================================================
PRIVATE FUNCTION CRebar.SetParent (BYVAL hwndParent AS hWND) AS HWND
   RETURN cast(HWND, SendMessageW(m_hCtl, RB_SETPARENT, cast(WPARAM, hwndParent), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a rebar control's default text color.
' Returns a COLORREF value that represents the previous default text color.
' ========================================================================================
PRIVATE FUNCTION CRebar.SetTextColor (BYVAL clrText AS COLORREF) AS COLORREF
   RETURN SendMessageW(m_hCtl, RB_SETTEXTCOLOR, 0, clrText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates a tooltip control with the rebar control.
' ========================================================================================
PRIVATE SUB CRebar.SetTooltips (BYVAL hwndToolTip AS HWND)
   SendMessageW(m_hCtl, RB_SETTOOLTIPS, cast(WPARAM, hwndToolTip), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the visual style of a rebar control.
' ========================================================================================
PRIVATE SUB CRebar.SetWindowTheme (BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(m_hCtl, RB_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
END SUB
' ========================================================================================

' ========================================================================================
' Shows or hides a given band in a rebar control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.ShowBand (BYVAL uBand AS DWORD, BYVAL fShow AS BOOLEAN) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, RB_SHOWBAND, uBand, fShow)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attempts to find the best layout of the bands for the given rectangle. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CRebar.SizeToRect (BYVAL prc AS RECT PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, RB_SIZETORECT, 0, cast(LPARAM, prc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CRebar.SizeToRect () AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, RB_SIZETORECT, 0, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

END NAMESPACE
