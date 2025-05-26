' ########################################################################################
' Platform: Microsoft Windows
' Filename: CUpDown.inc
' Purpose:  Windows common controls - UpDown control
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
' CUpDown class
' ########################################################################################

TYPE CUpDown EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION GetAccel (BYVAL cAccels AS DWORD, BYVAL paAccels AS UDACCEL PTR) AS LONG
   DECLARE FUNCTION GetBase () AS LONG
   DECLARE FUNCTION GetBuddy () AS HWND
   DECLARE FUNCTION GetPos () AS LONG
   DECLARE FUNCTION GetPos32 (BYVAL pfError AS LONG PTR = NULL) AS LONG
   DECLARE FUNCTION GetRange () AS LONG
   DECLARE SUB GetRange32 (BYVAL pLow AS LONG PTR, BYVAL pHigh AS LONG PTR)
   DECLARE FUNCTION SetAccel (BYVAL cAccels AS DWORD, BYREF paAccels AS UDACCEL PTR) AS BOOLEAN
   DECLARE FUNCTION SetBase (BYVAL nBase AS LONG) AS LONG
   DECLARE FUNCTION SetBuddy (BYVAL hwndBuddy AS HWND) AS HWND
   DECLARE FUNCTION SetPos (BYVAL nPos AS SHORT) AS LONG
   DECLARE FUNCTION SetPos32 (BYVAL nPos AS LONG) AS LONG
   DECLARE SUB SetRange (BYVAL nUpper AS SHORT, BYVAL nLower AS SHORT)
   DECLARE SUB SetRange32 (BYVAL iLow AS LONG, BYVAL iHigh AS LONG)

END TYPE
' ========================================================================================

' ========================================================================================
' * CUpDown constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CUpDown (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUpDown (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUpDown (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUpDown (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CUpDown destructor
' ========================================================================================
PRIVATE DESTRUCTOR CUpDown
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                     *** UPDOWN ***
' ########################################################################################

' ========================================================================================
' Retrieves acceleration information for an up-down control.
' The return value is the number of accelerators currently set for the control.
' ========================================================================================
PRIVATE FUNCTION CUpDown.GetAccel (BYVAL cAccels AS DWORD, BYVAL paAccels AS UDACCEL PTR) AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_GETACCEL, cAccels, CAST(LPARAM, paAccels))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current radix base (that is, either base 10 or 16) for an up-down control.
' ========================================================================================
PRIVATE FUNCTION CUpDown.GetBase () AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_GETBASE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the handle to the current buddy window.
' ========================================================================================
PRIVATE FUNCTION CUpDown.GetBuddy () AS HWND
   FUNCTION = cast(HWND, SendMessageW(m_hCtl, UDM_GETBUDDY, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current position of an up-down control with 16-bit precision.
' ========================================================================================
PRIVATE FUNCTION CUpDown.GetPos () AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the 32-bit position of an up-down control.
' ========================================================================================
PRIVATE FUNCTION CUpDown.GetPos32 (BYVAL pfError AS LONG PTR = NULL) AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_GETPOS32, 0, cast(LPARAM, pfError))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the 32-bit position of an up-down control.
' ========================================================================================
PRIVATE FUNCTION CUpDown.GetRange () AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_GETRANGE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the 32-bit position of an up-down control.
' ========================================================================================
PRIVATE SUB CUpDown.GetRange32 (BYVAL pLow AS LONG PTR, BYVAL pHigh AS LONG PTR)
   SendMessageW(m_hCtl, UDM_GETRANGE32, cast(WPARAM, pLow), cast(LPARAM, pHigh))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the acceleration for an up-down control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CUpDown.SetAccel (BYVAL cAccels AS DWORD, BYREF paAccels AS UDACCEL PTR) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, UDM_SETACCEL, cAccels, cast(LPARAM, paAccels))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the acceleration for an up-down control.
' The return value is the previous base value. If an invalid base is given, the return value is zero.
' ========================================================================================
PRIVATE FUNCTION CUpDown.SetBase (BYVAL nBase AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_SETBASE, nBase, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the buddy window for an up-down control.
' The return value is the handle to the previous buddy window.
' ========================================================================================
PRIVATE FUNCTION CUpDown.SetBuddy (BYVAL hwndBuddy AS HWND) AS HWND
   FUNCTION = cast(HWND, SendMessageW(m_hCtl, UDM_SETBUDDY, cast(WPARAM, hwndBuddy), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current position for an up-down control with 16-bit precision.
' The return value is the previous position.
' ========================================================================================
PRIVATE FUNCTION CUpDown.SetPos (BYVAL nPos AS SHORT) AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_SETPOS, 0, MAKELONG(nPos, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current position for an up-down control with 32-bit precision.
' The return value is the previous position.
' ========================================================================================
PRIVATE FUNCTION CUpDown.SetPos32 (BYVAL nPos AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, UDM_SETPOS32, 0, cast(LPARAM, nPos))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum and maximum positions (range) for an up-down control.
' ========================================================================================
PRIVATE SUB CUpDown.SetRange (BYVAL nUpper AS SHORT, BYVAL nLower AS SHORT)
   SendMessageW(m_hCtl, UDM_SETRANGE, 0, MAKELONG(nUpper, nLower))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the minimum and maximum positions (range) for an up-down control.
' ========================================================================================
PRIVATE SUB CUpDown.SetRange32 (BYVAL iLow AS LONG, BYVAL iHigh AS LONG)
   SendMessageW(m_hCtl, UDM_SETRANGE32, cast(WPARAM, iLow), cast(LPARAM, iHigh))
END SUB
' ========================================================================================

END NAMESPACE
