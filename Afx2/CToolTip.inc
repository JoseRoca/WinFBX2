' ########################################################################################
' Platform: Microsoft Windows
' Filename: CToolTip.inc
' Purpose:  Windows common controls - Tooltip control
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
' CToolTip class
' ########################################################################################

TYPE CToolTip EXTENDS CControl

Private:
   m_hwndTT AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND, BYVAL fBallon AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG, BYVAL fBallon AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG, BYVAL fBallon AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG, BYVAL fBallon AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR

   DECLARE SUB Activate
   DECLARE SUB Deactivate
   DECLARE FUNCTION AddTool (BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE FUNCTION AdjustRect (BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR) AS BOOLEAN
   DECLARE SUB DelTool (BYVAL lpti AS TOOLINFOW PTR)
   DECLARE FUNCTION EnumTools (BYVAL iTool AS DWORD, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE FUNCTION GetBubbleSize (BYVAL lpTtm AS TOOLINFOW PTR) AS DWORD
   DECLARE FUNCTION GetCurrentTool (BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE FUNCTION GetDelayTime (BYVAL dwDuration AS DWORD) AS LONG
   DECLARE FUNCTION GetFont () AS HFONT
   DECLARE SUB GetMargin (BYVAL lprc AS RECT PTR)
   DECLARE FUNCTION GetMaxTipWidth () AS LONG
   DECLARE SUB GetText (BYVAL numChars AS DWORD, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE FUNCTION GetTipBkColor () AS COLORREF
   DECLARE FUNCTION GetTipTextColor () AS COLORREF
   DECLARE SUB GetTitle (BYVAL pTTGetTitle AS TTGETTITLE PTR)
   DECLARE FUNCTION GetToolCount () AS LONG
   DECLARE FUNCTION GetToolInfo (BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE FUNCTION HitTest (BYVAL lphti AS TTHITTESTINFO PTR) AS BOOLEAN
   DECLARE SUB NewToolRect (BYVAL lpti AS TOOLINFOW PTR)
   DECLARE SUB Pop
   DECLARE SUB Popup
   DECLARE FUNCTION RelayEvent (BYVAL lpmsg AS tagMSG PTR) AS LONG
   DECLARE SUB SetDelayTime (BYVAL dwDuration AS DWORD, BYVAL iTime AS SHORT)
   DECLARE SUB SetFont (BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetMargin (BYVAL lprc AS RECT PTR)
   DECLARE FUNCTION SetMaxTipWidth (BYVAL iWidth AS LONG) AS LONG
   DECLARE SUB SetTipBkColor (BYVAL clr AS COLORREF)
   DECLARE SUB SetTipTextColor (BYVAL clr AS COLORREF)
   DECLARE FUNCTION SetTitle (BYVAL nIcon AS LONG, BYVAL pwszTitle AS WSTRING PTR) AS BOOLEAN
   DECLARE SUB SetToolInfo (BYVAL lpti AS TOOLINFOW PTR)
   DECLARE SUB SetWindowTheme (BYVAL pwszTheme AS WSTRING PTR)
   DECLARE SUB TrackActivate (BYVAL bActivate AS BOOLEAN, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE SUB TrackPosition (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   DECLARE SUB Update
   DECLARE SUB UpdateTipText (BYVAL lpti AS TOOLINFOW PTR)
   DECLARE FUNCTION WindowFromPoint (BYVAL lppt AS POINT PTR) AS HWND

   DECLARE FUNCTION AddTool (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING = "", BYVAL bCentered AS BOOLEAN = FALSE) AS BOOLEAN
   DECLARE FUNCTION SetText (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING) AS BOOLEAN
   DECLARE FUNCTION DelTool (BYVAL hwnd AS HWND) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CToolTip constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CToolTip (BYVAL hCtl AS HWND, BYVAL fBalloon AS BOOLEAN = FALSE)
   IF IsWindow(hCtl) THEN m_hwndTT = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
   ' // Create a tooltip control
   DIM dwStyle AS DWORD = WS_POPUP OR TTS_ALWAYSTIP
   IF fBalloon = TRUE THEN dwStyle = dwStyle OR TTS_BALLOON
   m_hwndTT = CreateWindowExW(0, "tooltips_class32", "", dwStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL)
   IF m_hwndTT = NULL THEN EXIT CONSTRUCTOR
   ' // You must explicitly define a tooltip control as topmost. Otherwise, it might be covered by the parent window.
   SetWindowPos(m_hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CToolTip (BYVAL hParent AS HWND, BYVAL cID AS LONG, BYVAL fBalloon AS BOOLEAN = FALSE)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hwndTT = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
   ' // Create a tooltip control
   DIM dwStyle AS DWORD = WS_POPUP OR TTS_ALWAYSTIP
   IF fBalloon = TRUE THEN dwStyle = dwStyle OR TTS_BALLOON
   m_hwndTT = CreateWindowExW(0, "tooltips_class32", "", dwStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL)
   IF m_hwndTT = NULL THEN EXIT CONSTRUCTOR
   ' // You must explicitly define a tooltip control as topmost. Otherwise, it might be covered by the parent window.
   SetWindowPos(m_hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CToolTip (BYREF pDlg AS CDialog, BYVAL cID AS LONG, BYVAL fBalloon AS BOOLEAN = FALSE)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hwndTT = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
   ' // Create a tooltip control
   DIM dwStyle AS DWORD = WS_POPUP OR TTS_ALWAYSTIP
   IF fBalloon = TRUE THEN dwStyle = dwStyle OR TTS_BALLOON
   m_hwndTT = CreateWindowExW(0, "tooltips_class32", "", dwStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL)
   IF m_hwndTT = NULL THEN EXIT CONSTRUCTOR
   ' // You must explicitly define a tooltip control as topmost. Otherwise, it might be covered by the parent window.
   SetWindowPos(m_hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CToolTip (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG, BYVAL fBalloon AS BOOLEAN = FALSE)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hwndTT = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
   ' // Create a tooltip control
   DIM dwStyle AS DWORD = WS_POPUP OR TTS_ALWAYSTIP
   IF fBalloon = TRUE THEN dwStyle = dwStyle OR TTS_BALLOON
   m_hwndTT = CreateWindowExW(0, "tooltips_class32", "", dwStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL)
   IF m_hwndTT = NULL THEN EXIT CONSTRUCTOR
   ' // You must explicitly define a tooltip control as topmost. Otherwise, it might be covered by the parent window.
   SetWindowPos(m_hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CToolTip destructor
' ========================================================================================
PRIVATE DESTRUCTOR CToolTip
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                   *** TOOLTIPS ***
' ########################################################################################

' ========================================================================================
' Activates a ToolTip control.
' ========================================================================================
PRIVATE SUB CTooltip.Activate
   SendMessageW(m_hwndTT, TTM_ACTIVATE, TRUE, 0)
END SUB
' ========================================================================================
' ========================================================================================
' Deactivates a ToolTip control.
' ========================================================================================
PRIVATE SUB CTooltip.Deactivate
   SendMessageW(m_hwndTT, TTM_ACTIVATE, FALSE, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Registers a tool with a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTooltip.AddTool (BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_ADDTOOLW, 0, CAST(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates a ToolTip control's text display rectangle from its window rectangle, or the
' ToolTip window rectangle needed to display a specified text display rectangle.
' ========================================================================================
PRIVATE FUNCTION CTooltip.AdjustRect (BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_ADJUSTRECT, cast(WPARAM, fLarger), cast(LPARAM, lprc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes a tool from a ToolTip control.
' ========================================================================================
PRIVATE SUB CTooltip.DelTool (BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(m_hwndTT, TTM_DELTOOLW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the information that a ToolTip control maintains about the current tool--that
' is, the tool for which the ToolTip is currently displaying text.
' ========================================================================================
PRIVATE FUNCTION CTooltip.EnumTools (BYVAL iTool AS DWORD, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_ENUMTOOLSW, cast(WPARAM, iTool), cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width and height of a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetBubbleSize (BYVAL lpTtm AS TOOLINFOW PTR) AS DWORD
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETBUBBLESIZE, 0, cast(LPARAM, lpTtm))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information for the current tool in a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetCurrentTool (BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETCURRENTTOOLW, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information for the current tool in a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetDelayTime (BYVAL dwDuration AS DWORD) AS LONG
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETDELAYTIME, cast(WPARAM, dwDuration), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle of the font used by the ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetFont () AS HFONT
   FUNCTION = cast(HFONT, SendMessageW(m_hwndTT, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the top, left, bottom, and right margins set for a ToolTip window. A margin is
' the distance, in pixels, between the ToolTip window border and the text contained within
' the ToolTip window.
' ========================================================================================
PRIVATE SUB CTooltip.GetMargin (BYVAL lprc AS RECT PTR)
   SendMessageW(m_hwndTT, TTM_GETMARGIN, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the maximum width for a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetMaxTipWidth () AS LONG
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETMAXTIPWIDTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information a ToolTip control maintains about a tool.
' ========================================================================================
PRIVATE SUB CTooltip.GetText (BYVAL numChars AS DWORD, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(m_hwndTT, TTM_GETTEXTW, numChars, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the background color in a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetTipBkColor () AS COLORREF
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETTIPBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text color in a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetTipTextColor () AS COLORREF
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETTIPTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information concerning the title of a tooltip control.
' ========================================================================================
PRIVATE SUB CTooltip.GetTitle (BYVAL pTTGetTitle AS TTGETTITLE PTR)
   SendMessageW(m_hwndTT, TTM_GETTITLE, 0, cast(LPARAM, pTTGetTitle))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves a count of the tools maintained by a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetToolCount () AS LONG
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETTOOLCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information that a ToolTip control maintains about a tool.
' ========================================================================================
PRIVATE FUNCTION CTooltip.GetToolInfo (BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_GETTOOLINFOW, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tests a point to determine whether it is within the bounding rectangle of the specified
' tool and, if it is, retrieves information about the tool.
' ========================================================================================
PRIVATE FUNCTION CTooltip.HitTest (BYVAL lphti AS TTHITTESTINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_HITTEST, 0, cast(LPARAM, lphti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a new bounding rectangle for a tool.
' ========================================================================================
PRIVATE SUB CTooltip.NewToolRect (BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(m_hwndTT, TTM_NEWTOOLRECTW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Removes a displayed ToolTip window from view.
' ========================================================================================
PRIVATE SUB CTooltip.Pop
   SendMessageW m_hwndTT, TTM_POP, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Causes the ToolTip to display at the coordinates of the last mouse message.
' ========================================================================================
PRIVATE SUB CTooltip.Popup
   SendMessageW m_hwndTT, TTM_POPUP, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Passes a mouse message to a ToolTip control for processing.
' ========================================================================================
PRIVATE FUNCTION CTooltip.RelayEvent (BYVAL lpmsg AS tagMSG PTR) AS LONG
   FUNCTION = SendMessageW(m_hwndTT, TTM_RELAYEVENT, 0, cast(LPARAM, lpmsg))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the initial, pop-up, and reshow durations for a ToolTip control.
' ========================================================================================
PRIVATE SUB CTooltip.SetDelayTime (BYVAL dwDuration AS DWORD, BYVAL iTime AS SHORT)
   SendMessageW(m_hwndTT, TTM_SETDELAYTIME, dwDuration, MAKELONG(iTime, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the font used by the ToolTip control.
' Parameters:
' - m_hwndTT  = Handle to the ToolTip control.
' - hFont     = Handle to the font. If this parameter is NULL, the control uses the default
'               system font to draw text.
' - fRedraw   = Specifies whether the control should be redrawn immediately upon setting the font.
'               If this parameter is TRUE, the control redraws itself.
' ========================================================================================
PRIVATE SUB CTooltip.SetFont (BYVAL hFont AS HFONT = NULL, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hwndTT, WM_SETFONT, cast(WPARAM, hFont), cast(LPARAM, fRedraw))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the top, left, bottom, and right margins for a ToolTip window. A margin is the
' distance, in pixels, between the ToolTip window border and the text contained within the
' ToolTip window.
' ========================================================================================
PRIVATE SUB CTooltip.SetMargin (BYVAL lprc AS RECT PTR)
   SendMessageW(m_hwndTT, TTM_SETMARGIN, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the maximum width for a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CTooltip.SetMaxTipWidth (BYVAL iWidth AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hwndTT, TTM_SETMAXTIPWIDTH, 0, cast(LPARAM, iWidth))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the background color in a ToolTip window.
' When visual styles are enabled, this message has no effect.
' ========================================================================================
PRIVATE SUB CTooltip.SetTipBkColor (BYVAL clr AS COLORREF)
   SendMessageW(m_hwndTT, TTM_SETTIPBKCOLOR, cast(WPARAM, clr), 0)
END SUB
' ========================================================================================

' ========================================================================================
' * Sets the text color in a ToolTip window.
' When visual styles are enabled, this message has no effect.
' ========================================================================================
PRIVATE SUB CTooltip.SetTipTextColor (BYVAL clr AS COLORREF)
   SendMessageW(m_hwndTT, TTM_SETTIPTEXTCOLOR, cast(WPARAM, clr), 0)
END SUB
' ========================================================================================

' ========================================================================================
' * Adds a standard icon and title string to a ToolTip.
' nIcon:
'   Value	Meaning
'   TTI_NONE - No icon.
'   TTI_INFO - Info icon.
'   TTI_WARNING - Warning icon
'   TTI_ERROR - Error Icon
'   TTI_INFO_LARGE - Large error Icon
'   TTI_WARNING_LARGE - Large error Icon
'   TTI_ERROR_LARGE - Large error Icon
' ========================================================================================
PRIVATE FUNCTION CTooltip.SetTitle (BYVAL nIcon AS LONG, BYVAL pwszTitle AS WSTRING PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hwndTT, TTM_SETTITLEW, nIcon, cast(LPARAM, pwszTitle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the information that a ToolTip control maintains for a tool.
' ========================================================================================
PRIVATE SUB CTooltip.SetToolInfo (BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(m_hwndTT, TTM_SETTOOLINFOW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the visual style of a ToolTip control.
' ========================================================================================
PRIVATE SUB CTooltip.SetWindowTheme (BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(m_hwndTT, TTM_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
END SUB
' ========================================================================================

' ========================================================================================
' Activates or deactivates a tracking ToolTip.
' ========================================================================================
PRIVATE SUB CTooltip.TrackActivate (BYVAL bActivate AS BOOLEAN, BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(m_hwndTT, TTM_TRACKACTIVATE, cast(WPARAM, bActivate), cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Activates or deactivates a tracking ToolTip.
' ========================================================================================
PRIVATE SUB CTooltip.TrackPosition (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   SendMessageW(m_hwndTT, TTM_TRACKPOSITION, 0, cast(LPARAM, MAKELONG(xPos, yPos)))
END SUB
' ========================================================================================

' ========================================================================================
' Activates or deactivates a tracking ToolTip.
' ========================================================================================
PRIVATE SUB CTooltip.Update
   SendMessageW(m_hwndTT, TTM_UPDATE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the ToolTip text for a tool.
' ========================================================================================
PRIVATE SUB CTooltip.UpdateTipText (BYVAL lpti AS TOOLINFOW PTR)
   SendMessageW(m_hwndTT, TTM_UPDATETIPTEXTW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Allows a subclass procedure to cause a ToolTip to display text for a window other than
' the one beneath the mouse cursor.
' ========================================================================================
PRIVATE FUNCTION CTooltip.WindowFromPoint (BYVAL lppt AS POINT PTR) AS HWND
   FUNCTION = cast(HWND, SendMessageW(m_hwndTT, TTM_WINDOWFROMPOINT, 0, cast(LPARAM, lppt)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a tooltip for a control.
' Parameters:
' - hwnd      = Handle of the window or control
' - wszText   = Tooltip text
' - bCentered = Centered (TRUE or FALSE)
' Return Value:
'   The handle of the tooltip control
' ========================================================================================
PRIVATE FUNCTION CToolTip.AddTool (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING = "", BYVAL bCentered AS BOOLEAN = FALSE) AS BOOLEAN
   IF m_hwndTT = NULL OR hwnd = NULL THEN RETURN FALSE
   IF IsWindow(hwnd) = 0 THEN RETURN FALSE
   ' // Registers the window with the tooltip control
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF this.ComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF this.ComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
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
   RETURN SendMessageW(m_hwndTT, TTM_ADDTOOLW, 0, CAST(LPARAM, @tti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets/replaces the text of a tooltip control
' Parameters:
' - hwnd     = Handle of the window or control
' - wszText  = Tooltip text
' ========================================================================================
PRIVATE FUNCTION CToolTip.SetText (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING) AS BOOLEAN
   IF m_hwndTT = NULL OR hwnd = NULL THEN RETURN FALSE
   IF IsWindow(hwnd) = 0 THEN RETURN FALSE
   IF VARPTR(wszText) = 0 THEN RETURN FALSE
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF this.ComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF this.ComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.hwnd = hwnd
      tti.uId = 0
   END IF
   ' // Retrieve the tooltip information
   IF SendMessageW(m_hwndTT, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti)) THEN
      tti.lpszText = CAST(LPWSTR, @wszText)
      SendMessageW(m_hwndTT, TTM_SETTOOLINFOW, 0, CAST(LPARAM, @tti))
      ' // Update the tooltip text
      SendMessageW(m_hwndTT, TTM_UPDATETIPTEXT, 0, CAST(LPARAM, @tti))
      ' // Check if the change has succeeded
      SendMessageW(m_hwndTT, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti))
      If tti.lpszText = @wszText THEN RETURN TRUE ELSE RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes a tool from a tooltip control.
' Parameters:
' - hwnd     = Handle of the window or control
' ========================================================================================
PRIVATE FUNCTION CToolTip.DelTool (BYVAL hwnd AS HWND) AS BOOLEAN
   IF m_hwndTT = NULL OR hwnd = NULL THEN RETURN FALSE
   IF IsWindow(hwnd) = 0 THEN RETURN FALSE
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF this.ComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF this.ComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.hwnd = hwnd
      tti.uId = 0
   END IF
   ' // Remove the tooltip
   SendMessageW(m_hwndTT, TTM_DELTOOL, 0, CAST(LPARAM, @tti))
   IF SendMessageW(m_hwndTT, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti)) = 0 THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================

END NAMESPACE
