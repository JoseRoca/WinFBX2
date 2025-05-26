' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTrackBar.inc
' Purpose:  Windows common controls - Trackbar
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
#include once "Afx2/CControl.inc"
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CTrackBar class
' ########################################################################################

TYPE CTrackBar EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE SUB ClearSel (BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB ClearTics (BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE FUNCTION GetBuddy (BYVAL fLocation AS BOOLEAN) AS HWND
   DECLARE SUB GetChannelRect (BYREF rc AS RECT)
   DECLARE FUNCTION GetChannelRect () AS RECT
   DECLARE FUNCTION GetLineSize () AS LONG
   DECLARE FUNCTION GetNumTics () AS LONG
   DECLARE FUNCTION GetPageSize () AS LONG
   DECLARE FUNCTION GetPos () AS LONG
   DECLARE FUNCTION GetPTics () AS DWORD
   DECLARE FUNCTION GetRangeMax () AS LONG
   DECLARE FUNCTION GetRangeMin () AS LONG
   DECLARE FUNCTION GetSelEnd () AS LONG
   DECLARE FUNCTION GetSelStart () AS LONG
   DECLARE FUNCTION GetThumbLength () AS LONG
   DECLARE SUB GetThumbRect (BYREF rc AS RECT)
   DECLARE FUNCTION GetThumbRect () AS RECT
   DECLARE FUNCTION GetTic (BYVAL iTic AS DWORD) AS LONG
   DECLARE FUNCTION GetTicPos (BYVAL iTic AS DWORD) AS LONG
   DECLARE FUNCTION GetTooltips () AS HWND
   DECLARE FUNCTION SetBuddy (BYVAL fLocation AS BOOLEAN, BYVAL hwndBuddy AS HWND) AS HWND
   DECLARE FUNCTION SetLineSize (BYVAL lLineSize AS LONG) AS LONG
   DECLARE FUNCTION SetPageSize (BYVAL lPageSize AS LONG) AS LONG
   DECLARE SUB SetPos (BYVAL fPosition AS BOOLEAN, BYVAL lPosition AS LONG)
   DECLARE SUB SetRange (BYVAL lMinimum AS LONG, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetRangeMax (BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetRangeMin (BYVAL lMinimum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetSel (BYVAL lMinimum AS LONG, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetSelEnd (BYVAL lEnd AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetSelStart (BYVAL lStart AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   DECLARE SUB SetThumbLength (BYVAL iLength AS LONG)
   DECLARE FUNCTION SetTic (BYVAL lPosition AS LONG) AS BOOLEAN
   DECLARE SUB SetTicFreq (BYVAL wFreq AS WORD)
   DECLARE FUNCTION SetTipSide (BYVAL fLocation AS LONG) AS LONG
   DECLARE SUB SetToolTips (BYVAL hwndTT AS HWND)

END TYPE
' ========================================================================================

' ========================================================================================
' * CTrackBar constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTrackBar (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTrackBar (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTrackBar (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTrackBar (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTrackBar destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTrackBar
END DESTRUCTOR
' ========================================================================================

' ########################################################################################
'                                     *** TRACKBAR ***
' ########################################################################################

' ========================================================================================
' Clears the current selection range in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.ClearSel (BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_CLEARSEL, fRedraw, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Removes the current tick marks from a trackbar. This message does not remove the first
' and last tick marks, which are created automatically by the trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.ClearTics (BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_CLEARTICS, fRedraw, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the handle to a trackbar control buddy window at a given location.
' Value indicating which buddy window handle will be retrieved, by relative location.
' This value can be one of the following:
' - TRUE : Retrieves the handle to the buddy to the left of the trackbar. If the trackbar
'   control uses the TBS_VERT style, the message will retrieve the buddy above the trackbar.
' - FALSE : Retrieves the handle to the buddy to the right of the trackbar. If the trackbar
'   control uses the TBS_VERT style, the message will retrieve the buddy below the trackbar.
' Returns the handle to the buddy window at the location specified by wParam, or NULL if no buddy window exists at that location.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetBuddy (BYVAL fLocation AS BOOLEAN) AS HWND
   RETURN CAST(HWND, SendMessage(m_hCtl, TBM_GETBUDDY, fLocation, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the size and position of the bounding rectangle for a trackbar's channel.
' (The channel is the area over which the slider moves. It contains the highlight when a
' range is selected.)
' ========================================================================================
PRIVATE SUB CTrackbar.GetChannelRect (BYREF rc AS RECT)
   IF VARPTR(rc) = 0 THEN ExIT SUB
   SendMessageW(m_hCtl, TBM_GETCHANNELRECT, 0, cast(LPARAM, @rc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetChannelRect () AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, TBM_GETCHANNELRECT, 0, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of logical positions the trackbar's slider moves in response to
' keyboard input from the arrow keys, such as the or keys. The logical positions are the
' integer increments in the trackbar's range of minimum to maximum slider positions.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetLineSize () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETLINESIZE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of tick marks in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetNumTics () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETNUMTICS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of logical positions the trackbar's slider moves in response to
' keyboard input, such as the or keys, or mouse input, such as clicks in the trackbar's
' channel. The logical positions are the integer increments in the trackbar's range of
' minimum to maximum slider positions.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetPageSize () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETPAGESIZE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current logical position of the slider in a trackbar. The logical positions
' are the integer values in the trackbar's range of minimum to maximum slider positions.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetPos () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the address of an array that contains the positions of the tick marks for a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetPTics () AS DWORD
   RETURN SendMessageW(m_hCtl, TBM_GETPTICS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the maximum position for the slider in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetRangeMax () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETRANGEMAX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the maximum position for the slider in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetRangeMin () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETRANGEMIN, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the ending position of the current selection range in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetSelEnd () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETSELEND, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the starting position of the current selection range in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetSelStart () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETSELSTART, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of the slider in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetThumbLength () AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETTHUMBLENGTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of the slider in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.GetThumbRect (BYREF rc AS RECT)
   IF VARPTR(rc) = 0 THEN EXIT SUB
   SendMessageW(m_hCtl, TBM_GETTHUMBRECT, 0, cast(LPARAM, @rc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetThumbRect () AS RECT
   DIM rc AS RECT
   SendMessageW(m_hCtl, TBM_GETTHUMBRECT, 0, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of the slider in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetTic (BYVAL iTic AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETTIC, iTic, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of the slider in a trackbar.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetTicPos (BYVAL iTic AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, TBM_GETTICPOS, iTic, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the ToolTip control assigned to the trackbar, if any.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.GetTooltips () AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TBM_GETTOOLTIPS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a window as the buddy window for a trackbar control. Trackbar buddy windows are
' automatically displayed in a location relative to the control's orientation (horizontal
' or vertical).
' Returns the handle to the window that was previously assigned to the control at that location.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.SetBuddy (BYVAL fLocation AS BOOLEAN, BYVAL hwndBuddy AS HWND) AS HWND
   RETURN CAST(HWND, SendMessageW(m_hCtl, TBM_SETBUDDY, fLocation, cast(LPARAM, hwndBuddy)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the number of logical positions the trackbar's slider moves in response to keyboard
' input from the arrow keys, such as the or keys. The logical positions are the integer
' increments in the trackbar's range of minimum to maximum slider positions.
' Returns a 32-bit value that specifies the previous line size.
' Remarks: The default setting for the line size is 1.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.SetLineSize (BYVAL lLineSize AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, TBM_SETLINESIZE, 0, lLineSize)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the number of logical positions the trackbar's slider moves in response to keyboard
' input, such as the or keys, or mouse input, such as clicks in the trackbar's channel.
' The logical positions are the integer increments in the trackbar's range of minimum to
' maximum slider positions.
' Returns a 32-bit value that specifies the previous page size.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.SetPageSize (BYVAL lPageSize AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, TBM_SETPAGESIZE, 0, lPageSize)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current logical position of the slider in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.SetPos (BYVAL fPosition AS BOOLEAN, BYVAL lPosition AS LONG)
   SendMessageW(m_hCtl, TBM_SETPOS, fPosition, lPosition)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the range of minimum and maximum logical positions for the slider in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.SetRange (BYVAL lMinimum AS LONG, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_SETRANGE, fRedraw, MAKELONG(lMinimum, lMaximum))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the maximum logical position for the slider in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.SetRangeMax (BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_SETRANGEMAX, fRedraw, lMaximum)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the minimum logical position for the slider in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.SetRangeMin (BYVAL lMinimum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_SETRANGEMIN, fRedraw, lMinimum)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the starting and ending positions for the available selection range in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.SetSel (BYVAL lMinimum AS LONG, BYVAL lMaximum AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_SETSEL, fRedraw, MAKELONG(lMinimum, lMaximum))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the ending logical position of the current selection range in a trackbar. This
' message is ignored if the trackbar does not have the TBS_ENABLESELRANGE style.
' ========================================================================================
PRIVATE SUB CTrackbar.SetSelEnd (BYVAL lEnd AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_SETSELEND, fRedraw, lEnd)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the starting logical position of the current selection range in a trackbar. This
' message is ignored if the trackbar does not have the BS_ENABLESELRANGE style.
' ========================================================================================
PRIVATE SUB CTrackbar.SetSelStart (BYVAL lStart AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
   SendMessageW(m_hCtl, TBM_SETSELSTART, fRedraw, lStart)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the length of the slider in a trackbar. This message is ignored if the trackbar
' does not have the TBS_FIXEDLENGTH style.
' ========================================================================================
PRIVATE SUB CTrackbar.SetThumbLength (BYVAL iLength AS LONG)
   SendMessageW(m_hCtl, TBM_SETTHUMBLENGTH, iLength, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets a tick mark in a trackbar at the specified logical position.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.SetTic (BYVAL lPosition AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, TBM_SETTIC, 0, lPosition)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the interval frequency for tick marks in a trackbar.
' ========================================================================================
PRIVATE SUB CTrackbar.SetTicFreq (BYVAL wFreq AS WORD)
   SendMessageW(m_hCtl, TBM_SETTICFREQ, wFreq, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Positions a ToolTip control used by a trackbar control. Trackbar controls that use the
' TBS_TOOLTIPS style display ToolTips.
' ========================================================================================
PRIVATE FUNCTION CTrackbar.SetTipSide (BYVAL fLocation AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, TBM_SETTIPSIDE, fLocation, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns a ToolTip control to a trackbar control.
' ========================================================================================
PRIVATE SUB CTrackbar.SetToolTips (BYVAL hwndTT AS HWND)
   SendMessageW(m_hCtl, TBM_SETTOOLTIPS, cast(WPARAM, hwndTT), 0)
END SUB
' ========================================================================================

END NAMESPACE
