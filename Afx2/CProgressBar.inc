' ########################################################################################
' Platform: Microsoft Windows
' Filename: CProgressBar.inc
' Purpose:  Windows common controls - Progress bar
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
' CProgressBar class
' ########################################################################################

TYPE CProgressBar EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION DeltaPos (BYVAL nIncrement AS LONG) AS LONG
   DECLARE FUNCTION GetHighLimit () AS INT_
   DECLARE FUNCTION GetLowLimit () AS LONG
   DECLARE FUNCTION GetMaxRange () AS LONG
   DECLARE FUNCTION GetMinRange () AS LONG
   DECLARE FUNCTION GetPos () AS DWORD
   DECLARE FUNCTION GetRange (BYVAL fWhichLimit AS BOOLEAN, BYVAL ppBRange AS PBRANGE PTR) AS LONG
   DECLARE FUNCTION IncPos (BYVAL nIncVal AS LONG) AS LONG
   DECLARE FUNCTION SetBarColor (BYVAL clrBar AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetBkColor (BYVAL clrBk AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetMarquee (BYVAL fOn AS BOOLEAN, BYVAL dwAnimationSpeed AS DWORD) AS LONG
   DECLARE FUNCTION SetMarqueeOff () AS LONG
   DECLARE FUNCTION SetMarqueeOn (BYVAL dwAnimationSpeed AS DWORD) AS LONG
   DECLARE FUNCTION SetPos (BYVAL nNewPos AS LONG) AS LONG
   DECLARE FUNCTION SetRange (BYVAL nMinRange AS WORD, BYVAL nMaxRange AS WORD) AS DWORD
   DECLARE FUNCTION SetRange32 (BYVAL nMinRange AS LONG, BYVAL nMaxRange AS LONG) AS DWORD
   DECLARE FUNCTION SetStep (BYVAL nStepInc AS LONG) AS LONG
   DECLARE FUNCTION StepIt () AS LONG
   #if _WIN32_WINNT = &h0602
   DECLARE FUNCTION GetBarColor () AS COLORREF
   DECLARE FUNCTION GetBkColor () AS COLORREF
   DECLARE FUNCTION GetState () AS LONG
   DECLARE FUNCTION SetState (BYVAL iState AS LONG) AS LONG
   #endif

END TYPE
' ========================================================================================

' ========================================================================================
' * CProgressBar constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CProgressBar (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CProgressBar (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CProgressBar (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CProgressBar (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CProgressBar destructor
' ========================================================================================
PRIVATE DESTRUCTOR CProgressBar
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                  *** PROGRESS BAR ***
' ########################################################################################

' ========================================================================================
' Advances the current position of a progress bar by a specified increment and redraws the
' bar to reflect the new position. Returns the previous position.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.DeltaPos (BYVAL nIncrement AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_DELTAPOS, nIncrement, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current high limit of a given progress bar control.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetHighLimit () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_GETRANGE, FALSE, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current low limit of a given progress bar control.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetLowLimit () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_GETRANGE, TRUE, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current maximum range of a given progress bar control.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetMaxRange () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_GETRANGE, FALSE, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current minimum range of a given progress bar control.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetMinRange () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_GETRANGE, TRUE, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current position of the progress bar.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetPos () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, PBM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the current high and low limits of a given progress bar control.
' - fWhichLimit: Flag value specifying which limit value is to be used as the message's return
'   value. This parameter can be one of the following values: TRUE (return the low limit),
'   FALSE (return the high limit).
' - ppBRange: Pointer to a PBRANGE structure that is to be filled with both limit values.
' Returns an INT that represents the limit value specified by fWhichLimit
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetRange (BYVAL fWhichLimit AS BOOLEAN, BYVAL ppBRange AS PBRANGE PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_GETRANGE, fWhichLimit, cast(LPARAM, ppBRange))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Advances the current position for a progress bar by the specified increment and redraws
' the bar to reflect the new position. nIncVaL: Signed integer that becomes the new position.
' If nIncVal is outside the range of the control, the position is set to the closest boundary.
' Do not send this message to a control that has the PBS_MARQUEE style.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.IncPos (BYVAL nIncVal AS LONG) AS LONG
   DIM nPos AS LONG = SendMessageW(m_hCtl, PBM_GETPOS, 0, 0)
   FUNCTION = SendMessageW(m_hCtl, PBM_SETPOS, nPos + nIncVal, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the color of the progress indicator bar in the progress bar control.
' - clrBar: The COLORREF value that specifies the new progress indicator bar color.
'   Specifying the CLR_DEFAULT value causes the progress bar to use its default progress
'   indicator bar color.
' Returns the previous progress indicator bar color, or CLR_DEFAULT if the progress
' indicator bar color is the default color.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetBarColor (BYVAL clrBar AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(m_hCtl, PBM_SETBARCOLOR, 0, clrBar)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color in the progress bar.
' - clrBk: COLORREF value that specifies the new background color. Specify the CLR_DEFAULT
'   value to cause the progress bar to use its default background color.
' Returns the previous background color, or CLR_DEFAULT if the background color is the default color.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetBkColor (BYVAL clrBk AS COLORREF) AS COLORREF
   FUNCTION = SendMessageW(m_hCtl, PBM_SETBKCOLOR, 0, clrBk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the progress bar to marquee mode. This causes the progress bar to move like a marquee.
' - fOn: Indicates whether to turn the marquee mode on or off.
' - dwAnimationSpeed: Time, in milliseconds, between marquee animation updates. If this
'   parameter is zero, the marquee animation is updated every 30 milliseconds.
' Return value: Always returns TRUE.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetMarquee (BYVAL fOn AS BOOLEAN, BYVAL dwAnimationSpeed AS DWORD) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_SETMARQUEE, fOn, dwAnimationSpeed)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the progress bar marquee mode to off.
' Return value: Always returns TRUE.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetMarqueeOff () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_SETMARQUEE, FALSE, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the progress bar to marquee mode. This causes the progress bar to move like a marquee.
' Return value: Always returns TRUE.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetMarqueeOn (BYVAL dwAnimationSpeed AS DWORD) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_SETMARQUEE, CTRUE, dwAnimationSpeed)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current position for a progress bar and redraws the bar to reflect the new position.
' - nNewPos: Signed integer that becomes the new position.
' Returns the previous position.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetPos (BYVAL nNewPos AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_SETPOS, nNewPos, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current position for a progress bar and redraws the bar to reflect the new position.
' The minimum range value must not be negative. By default, the minimum value is zero.
' The maximum range value must be greater than the minimum range value. By default, the
' maximum range value is 100.
' Returns the previous range values if successful, or zero otherwise. The LOWORD specifies
' the previous minimum value, and the HIWORD specifies the previous maximum value.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetRange (BYVAL nMinRange AS WORD, BYVAL nMaxRange AS WORD) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, PBM_SETRANGE, 0, MAKELONG(nMinRange, nMaxRange))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum and maximum values for a progress bar to 32-bit values, and redraws the
' bar to reflect the new range. Returns a DWORD value that holds the previous 16-bit low
' limit in its LOWORD and the previous 16-bit high limit in its HIWORD. If the previous
' ranges were 32-bit values, the return value consists of the LOWORDs of both 32-bit limits.
' To retrieve the entire high and low 32-bit values, use the PBM_GETRANGE message.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetRange32 (BYVAL nMinRange AS LONG, BYVAL nMaxRange AS LONG) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, PBM_SETRANGE32, nMinRange, nMaxRange)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the step increment for a progress bar. The step increment is the amount by
' which the progress bar increases its current position whenever it receives a PBM_STEPIT
' message. By default, the step increment is set to 10.
' Return value: Returns the previous step increment.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetStep (BYVAL nStepInc AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_SETSTEP, nStepInc, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Advances the current position for a progress bar by the step increment and redraws the
' bar to reflect the new position. An application sets the step increment by sending the
' PBM_SETSTEP message. Return value: Returns the previous position.
' Remarks: When the position exceeds the maximum range value, this message resets the
' current position so that the progress indicator starts over again from the beginning.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.StepIt () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_STEPIT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the color of the progress bar.
' ========================================================================================
#if _WIN32_WINNT = &h0602
PRIVATE FUNCTION CProgressBar.GetBarColor () AS COLORREF
   FUNCTION = SendMessageW(m_hCtl, PBM_GETBARCOLOR, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the color of the progress bar.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetBkColor () AS COLORREF
   FUNCTION = SendMessageW(m_hCtl, PBM_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current state of the progress bar. One of the following values.
' - PBST_NORMAL: In progress
' - PBST_ERROR: Error
' - PBST_PAUSED: Paused
' ========================================================================================
PRIVATE FUNCTION CProgressBar.GetState () AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_GETSTATE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of the progress bar. Can be one of the following values:
' PBST_NORMAL (in progress), PBST_ERROR (error), PBST_PAUSED (paused).
' Return value: Returns the previous state.
' ========================================================================================
PRIVATE FUNCTION CProgressBar.SetState (BYVAL iState AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, PBM_SETSTATE, iState, 0)
END FUNCTION
' ========================================================================================
#endif

END NAMESPACE
