' ########################################################################################
' Platform: Microsoft Windows
' Filename: CControl.inc
' Purpose:  Base classs for Windows controls classes
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

NAMESPACE Afx2

' ========================================================================================
' CControl class
' ========================================================================================
TYPE CControl

Private:
   m_Result AS LONG                           ' // Last result code

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS LONG
   DECLARE FUNCTION SetResult (BYVAL Result AS LONG) AS LONG
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING

   DECLARE FUNCTION DialogUnitsToPixels (BYVAL hCtl AS HWND, BYVAL dluX AS LONG, BYVAL dluY AS LONG, BYREF pixelX AS LONG, BYREF pixelY AS LONG) AS BOOLEAN
   DECLARE FUNCTION DialogUnitsToPixelsRatios (BYVAL hCtl AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   DECLARE FUNCTION DlgToPixRX (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION DlgToPixRY (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION PixelsToDialogUnits (BYVAL hCtl AS HWND, BYVAL pixelX AS LONG, BYVAL pixelY AS LONG, BYREF dluX AS LONG, BYREF dluY AS LONG) AS BOOLEAN
   DECLARE FUNCTION PixelsToDialogUnitsRatios (BYVAL hCtl AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   DECLARE FUNCTION PixToDlgRX (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION PixToDlgRY (BYVAL hCtl AS HWND) AS SINGLE


END TYPE
' ========================================================================================

' ========================================================================================
' * CControl constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CControl
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CButton destructor
' ========================================================================================
PRIVATE DESTRUCTOR CControl
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code
' ========================================================================================
PRIVATE FUNCTION CControl.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CControl.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CControl.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" & HEX(nError) & CHR(13, 10) & dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts the specified dialog box units to screen units (pixels).
' ========================================================================================
PRIVATE FUNCTION CControl.DialogUnitsToPixels (BYVAL hCtl AS HWND, BYVAL dluX AS LONG, BYVAL dluY AS LONG, BYREF pixelX AS LONG, BYREF pixelY AS LONG) AS BOOLEAN
   IF VARPTR(pixelX) = 0 OR VARPTR(pixelY) = 0 THEN RETURN FALSE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = dluX : rc.Bottom = dluY
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(pixelX) THEN pixelX = rc.Right
   IF VARPTR(pixelY) THEN pixelY = rc.Bottom
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratios from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION CControl.DialogUnitsToPixelsRatios (BYVAL hCtl AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   IF VARPTR(ratioX) = 0 OR VARPTR(ratioY) = 0 THEN RETURN FALSE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 100 : rc.Bottom = 100
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(ratioX) THEN ratioX = rc.Right / 100
   IF VARPTR(ratioY) THEN ratioY = rc.Bottom / 100
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION CControl.DlgToPixRX (BYVAL hCtl AS HWND) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN rc.Right / 4
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION CControl.DlgToPixRY (BYVAL hCtl AS HWND) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN rc.Bottom / 8
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION CControl.PixToDlgRX (BYVAL hCtl AS HWND) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN 4.0 / rc.right
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION CControl.PixToDlgRY (BYVAL hCtl AS HWND) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN 8.0 / rc.bottom
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts the specified screen units (pixels) to dialog box units.
' ========================================================================================
PRIVATE FUNCTION CControl.PixelsToDialogUnits (BYVAL hCtl AS HWND, BYVAL pixelX AS LONG, BYVAL pixelY AS LONG, BYREF dluX AS LONG, BYREF dluY AS LONG) AS BOOLEAN
   IF VARPTR(pixelX) = 0 OR VARPTR(pixelY) = 0 THEN RETURN FALSE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4
   rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(dluX) THEN dluX = pixelX * (4.0 / rc.right)
   IF VARPTR(dluY) THEN dluY = pixelY * (8.0 / rc.bottom)
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the conversion ratios from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION CControl.PixelsToDialogUnitsRatios (BYVAL hCtl AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   IF VARPTR(ratioX) = 0 OR VARPTR(ratioY) = 0 THEN RETURN FALSE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN FALSE
   IF VARPTR(ratioX) THEN ratioX = 4.0 / rc.right
   IF VARPTR(ratioY) THEN ratioY = 8.0 / rc.bottom
   RETURN TRUE
END FUNCTION
' ========================================================================================

END NAMESPACE
