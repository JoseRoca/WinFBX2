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

#if not defined(UNICODE)
   #define UNICODE
#endif
#if not defined(_WIN32_WINNT)
   #define _WIN32_WINNT &h0602
#endif

#pragma once
#include once "windows.bi"
#include once "win/commctrl.bi"
#include once "Afx2/DWSTRING.inc"
#include once "Afx2/CDialog.inc"

NAMESPACE Afx2

' ========================================================================================
' CControl class
' ========================================================================================
TYPE CControl

Private:
   m_Result AS LONG                        ' // Last result code
   m_UsePixels AS BOOLEAN                  ' // Use pixels instead of dialog units
   DIM m_rgUserData(0 TO 19) AS LONG_PTR   ' // User data

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS LONG
   DECLARE FUNCTION SetResult (BYVAL Result AS LONG) AS LONG
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING
   DECLARE FUNCTION CDialogOwnerPtr (BYVAL hwnd AS HWND) AS CDialog PTR
   DECLARE PROPERTY UsesPixels () AS BOOLEAN
   DECLARE PROPERTY UsesPixels (BYVAL bUsePixels AS BOOLEAN)
   DECLARE FUNCTION GetUser (BYVAL idx AS LONG) AS LONG_PTR
   DECLARE SUB SetUser (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)

   DECLARE FUNCTION DialogUnitsToPixels (BYVAL hCtl AS HWND, BYVAL dluX AS LONG, BYVAL dluY AS LONG, BYREF pixelX AS LONG, BYREF pixelY AS LONG) AS BOOLEAN
   DECLARE FUNCTION DialogUnitsToPixelsRatios (BYVAL hCtl AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   DECLARE FUNCTION DluToPixRX (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION DluToPixRY (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION DluToPixX (BYVAL hCtl AS HWND, BYVAL dlu AS LONG) AS SINGLE
   DECLARE FUNCTION DluToPixY (BYVAL hCtl AS HWND, BYVAL dlu AS LONG) AS SINGLE
   DECLARE FUNCTION PixelsToDialogUnits (BYVAL hCtl AS HWND, BYVAL pixelX AS LONG, BYVAL pixelY AS LONG, BYREF dluX AS LONG, BYREF dluY AS LONG) AS BOOLEAN
   DECLARE FUNCTION PixelsToDialogUnitsRatios (BYVAL hCtl AS HWND, BYREF ratioX AS SINGLE, BYREF ratioY AS SINGLE) AS BOOLEAN
   DECLARE FUNCTION PixToDluRX (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION PixToDluRY (BYVAL hCtl AS HWND) AS SINGLE
   DECLARE FUNCTION PixToDluX (BYVAL hCtl AS HWND, BYVAL pix AS LONG) AS SINGLE
   DECLARE FUNCTION PixToDluY (BYVAL hCtl AS HWND, BYVAL pix AS LONG) AS SINGLE
   DECLARE FUNCTION ScaleRX () AS SINGLE
   DECLARE FUNCTION ScaleRY () AS SINGLE
   DECLARE FUNCTION ScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DECLARE FUNCTION ScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DECLARE FUNCTION UnScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DECLARE FUNCTION UnScaleY (BYVAL cy AS SINGLE) AS SINGLE

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
' Returns true if it uses pixels
' ========================================================================================
PRIVATE PROPERTY CControl.UsesPixels () AS BOOLEAN
   RETURN m_UsePixels
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CControl.UsesPixels (BYVAL bUsePixels AS BOOLEAN)
   m_UsePixels = bUsePixels
END PROPERTY
' ========================================================================================

' =====================================================================================
' * Retrieves a value from the user data area of the control.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE FUNCTION CControl.GetUser (BYVAL idx AS LONG) AS LONG_PTR
   IF idx < 0 OR idx > 19 THEN RETURN 0
   RETURN m_rgUserData(idx)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets a value in the user data area of the control.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE SUB CControl.SetUser (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
   IF idx > -1 AND idx < 20 THEN m_rgUserData(idx) = newValue
END SUB
' =====================================================================================

' ========================================================================================
' Returns a pointer to the CDialog class given the handle of the window created with it
' or the handle of any of it's children.
' ========================================================================================
PRIVATE FUNCTION CControl.CDialogOwnerPtr (BYVAL hwnd AS HWND) AS CDialog PTR
   IF hwnd = NULL THEN RETURN NULL
   DIM hRootOwner AS ..HWND = GetAncestor(hwnd, GA_ROOTOWNER)
   IF hRootOwner = NULL THEN RETURN NULL
   RETURN CAST(CDialog PTR, .GetWindowLongPtrW(hRootOwner, DWLP_USER))
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
PRIVATE FUNCTION CControl.DluToPixRX (BYVAL hCtl AS HWND) AS SINGLE
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
PRIVATE FUNCTION CControl.DluToPixRY (BYVAL hCtl AS HWND) AS SINGLE
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
PRIVATE FUNCTION CControl.PixToDluRX (BYVAL hCtl AS HWND) AS SINGLE
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
PRIVATE FUNCTION CControl.PixToDluRY (BYVAL hCtl AS HWND) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4 : rc.Bottom = 8
   MapDialogRect(hDlg, @rc)
   RETURN 8.0 / rc.bottom
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Converts dialog uits to pixels
' ========================================================================================
PRIVATE FUNCTION CControl.DluToPixX (BYVAL hCtl AS HWND, BYVAL dlu AS LONG) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = dlu
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN 0
   RETURN rc.Right
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CControl.DluToPixY (BYVAL hCtl AS HWND, BYVAL dlu AS LONG) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Bottom = dlu
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN 0
   RETURN rc.Bottom
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
' * Converts the specified screen units (pixels) to dialog box units.
' ========================================================================================
PRIVATE FUNCTION CControl.PixToDluX (BYVAL hCtl AS HWND, BYVAL pix AS LONG) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4
   rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN 0
   RETURN pix * (4.0 / rc.right)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CControl.PixToDluY (BYVAL hCtl AS HWND, BYVAL pix AS LONG) AS SINGLE
   DIM hDlg AS HWND = GetParent(hCtl)
   DIM rc AS RECT
   rc.Right = 4
   rc.Bottom = 8
   DIM bResult AS BOOLEAN = MapDialogRect(hDlg, @rc)
   IF bResult = FALSE THEN RETURN 0
   RETURN pix * (8.0 / rc.bottom)
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

' ========================================================================================
' Retrieves the desktop horizontal scaling ratio.
' ========================================================================================
PRIVATE FUNCTION CControl.ScaleRX () AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM rxRatio AS SINGLE = (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = rxRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the desktop vertical scaling ratio.
' ========================================================================================
PRIVATE FUNCTION CControl.ScaleRY () AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM ryRatio AS SINGLE = (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = ryRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION CControl.ScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cx * (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION CControl.ScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cy * (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION CControl.UnScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cx / (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION CControl.UnScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cy / (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

END NAMESPACE
