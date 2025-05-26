' ########################################################################################
' Platform: Microsoft Windows
' Filename: CScrollWindow.inc
' Purpose:  Scrollable window
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
'using Afx2

NAMESPACE Afx2

' ########################################################################################
'                              *** CSCROLLWINDOW CLASS ***
' ########################################################################################

TYPE CScrollWindow

Private:
   Private:
      m_hwnd AS HWND             ' // Handle of the window
      m_ClientRect AS RECT       ' // Coordinates of the window's client area
      m_nHorzInc AS LONG         ' // Horizontal increment
      m_nVertInc AS LONG         ' // Vertical increment
      m_HScrollMax AS LONG       ' // Maximum horizontal scroll value
      m_VScrollMax AS LONG       ' // Maximum vertical scroll value
      m_HScrollPos AS LONG       ' // Horizontal scroll position
      m_VScrollPos AS LONG       ' // Vertical scroll position
      m_HorzUnits AS LONG        ' // Amount, in device units, of horizontal scrolling
      m_VertUnits AS LONG        ' // Amount, in device units, of vertical scrolling

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScrollbars
      DECLARE SUB ResetScrollbars
'      DECLARE SUB SetClientRect (BYVAL pRect AS RECT PTR)
'      DECLARE SUB SetScrollingUnits (BYVAL nHorzUnits AS LONG, BYVAL nVertUnits AS LONG)
      DECLARE SUB OnSize (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB OnVScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB OnHScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
'      DECLARE SUB OnMouseWheel (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

END TYPE
' ========================================================================================

' ========================================================================================
' CScrollWindow class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CScrollWindow (BYVAL hwnd AS HWND)
   m_hwnd = hwnd
   m_HorzUnits = 10
   m_VertUnits = 10
   GetClientRect(m_hwnd, @m_ClientRect)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CScrollWindow class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CScrollWindow
'   OutputDebugStringW "CScrollWindow.Destructor"
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Sets the amount, in device units, of horizontal and vertical scrolling
' ========================================================================================
'PRIVATE SUB CScrollWindow.SetClientRect (BYVAL pRect AS RECT PTR)
'   IF pRect = NULL THEN EXIT SUB
'   m_ClientRect = *pRect
'END SUB
' ========================================================================================

' ========================================================================================
' Sets the amount, in device units, of horizontal and vertical scrolling
' ========================================================================================
'PRIVATE SUB CScrollWindow.SetScrollingUnits (BYVAL nHorzUnits AS LONG, BYVAL nVertUnits AS LONG)
'   m_HorzUnits = nHorzUnits
'   m_VertUnits = nVertUnits
'END SUB
' ========================================================================================

' ========================================================================================
' Resets the window scrolling information
' ========================================================================================
PRIVATE SUB CScrollWindow.ResetScrollbars
   ScrollWindow(m_hwnd, m_HScrollPos * m_HorzUnits, 0, NULL, NULL)
   ScrollWindow(m_hwnd, 0, m_VScrollPos * m_HorzUnits, NULL, NULL)
   m_HScrollPos = 0
   m_VScrollPos = 0
   SetScrollPos(m_hwnd, SB_HORZ, m_HScrollPos, CTRUE)
   SetScrollPos(m_hwnd, SB_VERT, m_VScrollPos, CTRUE)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the scroll information
' ========================================================================================
PRIVATE SUB CScrollWindow.SetupScrollbars

   DIM bMaximized AS LONG, tempRect AS RECT, wp AS WINDOWPLACEMENT

   GetClientRect(m_hwnd, @tempRect)
   GetWindowPlacement(m_hwnd, @wp)
   bMaximized = (wp.showCmd = SW_MAXIMIZE)

   ' // Calculate how many scrolling increments for the client area
   m_nHorzInc = ((m_ClientRect.Right - m_ClientRect.Left) - (tempRect.Right  - tempRect.Left) ) \ m_HorzUnits
   m_nVertInc = ((m_ClientRect.Bottom - m_ClientRect.Top) - (tempRect.bottom - tempRect.Top ) ) \ m_VertUnits

   '// Set the vertical and horizontal scrolling info
   m_HScrollMax = MAX(0, m_nHorzInc)
   m_HScrollPos = MIN(m_HScrollPos, m_HScrollMax)
   SetScrollRange(m_hwnd, SB_HORZ, 0, m_HScrollMax, FALSE)
   SetScrollPos(m_hwnd, SB_HORZ, m_HScrollPos, CTRUE)
   m_VScrollMax = MAX(0, m_nVertInc)
   m_VScrollPos = MIN(m_VScrollPos, m_VScrollMax)
   SetScrollRange(m_hwnd, SB_VERT, 0, m_VScrollMax, FALSE)
   SetScrollPos(m_hwnd, SB_VERT, m_VScrollPos, TRUE)

END SUB
' ========================================================================================

' ========================================================================================
' Handle vertical scrollbar messages
' WM_VSCROLL
' nScrollCode = (int) LOWORD(wParam); // scroll bar value
' nPos = (short int) HIWORD(wParam);  // scroll box position
' hwndScrollBar = (HWND) lParam;      // handle of scroll bar
' ========================================================================================
PRIVATE SUB CScrollWindow.OnVScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

   DIM nInc AS LONG, iMove AS LONG

   SELECT CASE LOWORD(wParam)
      CASE SB_TOP        : nInc = -m_VScrollPos
      CASE SB_BOTTOM     : nInc = m_VScrollMax - m_VScrollPos
      CASE SB_LINEUP     : nInc = -1
      CASE SB_LINEDOWN   : nInc = 1
      CASE SB_PAGEUP     : nInc = MIN(-1, -m_nVertInc)
      CASE SB_PAGEDOWN   : nInc = MAX(1, m_nVertInc)
      CASE SB_THUMBTRACK : nInc = HIWORD(wParam) - m_VScrollPos
      CASE ELSE          : nInc = 0
   END SELECT

   nInc = MAX(-m_VScrollPos, MIN(nInc, m_VScrollMax - m_VScrollPos))
   IF nInc THEN
      m_VScrollPos += nInc
      iMove = -m_HorzUnits * nInc
      ScrollWindow(m_hwnd, 0, iMove, NULL, NULL)
      SetScrollPos(m_hwnd, SB_VERT, m_VScrollPos, CTRUE)
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' Handle horizontal scrollbar messages
' WM_HSCROLL
' nScrollCode = (int) LOWORD(wParam);  // scroll bar value
' nPos = (short int) HIWORD(wParam);   // scroll box position
' hwndScrollBar = (HWND) lParam;       // handle of scroll bar
' ========================================================================================
PRIVATE SUB CScrollWindow.OnHScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

   DIM nInc AS LONG, iMove AS LONG

   SELECT CASE LOWORD(wParam)
      CASE SB_TOP        : nInc = -m_HScrollPos
      CASE SB_BOTTOM     : nInc = m_HScrollMax - m_HScrollPos
      CASE SB_LINEUP     : nInc = -1
      CASE SB_LINEDOWN   : nInc = 1
      CASE SB_PAGEUP     : nInc = -m_HorzUnits
      CASE SB_PAGEDOWN   : nInc = m_HorzUnits
      CASE SB_THUMBTRACK : nInc = HIWORD(wParam) - m_HScrollPos
      CASE ELSE          : nInc = 0
   END SELECT

   nInc = MAX(-m_HScrollPos, MIN(nInc, m_HScrollMax - m_HScrollPos))
   IF nInc THEN
      m_HScrollPos += nInc
      iMove = -m_HorzUnits * nInc
      ScrollWindow(m_hwnd, iMove, 0, NULL, NULL)
      SetScrollPos(m_hwnd, SB_HORZ, m_HScrollPos, CTRUE)
  END IF

END SUB
' ========================================================================================

' ========================================================================================
' Handle WM_SIZE messges
' WM_SIZE
' fwSizeType = wParam;      // resizing flag
' nWidth = LOWORD(lParam);  // width of client area
' nHeight = HIWORD(lParam); // height of client area
' ========================================================================================
PRIVATE SUB CScrollWindow.OnSize (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
   this.ResetScrollbars
   this.SetupScrollbars
END SUB
' ========================================================================================

END NAMESPACE
