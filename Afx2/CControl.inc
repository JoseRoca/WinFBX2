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
#include once "win/windowsx.bi"
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
   DECLARE FUNCTION GetFileVersion (BYVAL pwszFileName AS WSTRING PTR) AS LONG
   DECLARE FUNCTION ComCtlVersion () AS LONG

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
' Returns true if the dialog uses pixels
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
' idx = The index number of the user data value to retrieve, in the range 0 to 19 inclusive.
' =====================================================================================
PRIVATE FUNCTION CControl.GetUser (BYVAL idx AS LONG) AS LONG_PTR
   IF idx < 0 OR idx > 19 THEN RETURN 0
   RETURN m_rgUserData(idx)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets a value in the user data area of the control.
' idx = The index number of the user data value to set, in the range 0 to 19 inclusive.
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
' Returns the version of specified file multiplied by 100, e.g. 601 for version 6.01.
' Example: DIM ver AS LONG = AfxGetFileVersion("COMCTL32.DLL")
' ========================================================================================
PRIVATE FUNCTION CControl.GetFileVersion (BYVAL pwszFileName AS WSTRING PTR) AS LONG
   DIM pvsffi AS VS_FIXEDFILEINFO PTR, dwHandle AS DWORD
   DIM cbLen AS DWORD = GetFileVersionInfoSizeW(pwszFileName, @dwHandle)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM pVerInfo AS HANDLE = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, cbLen)
   IF pVerInfo = NULL THEN EXIT FUNCTION
   IF GetFileVersionInfoW(pwszFileName, dwHandle, cbLen, pVerInfo) THEN
      IF VerQueryValueW(pVerInfo, "\", @pvsffi, @cbLen) THEN
         DIM wMajor AS WORD = HIWORD(pvsffi->dwFileVersionMS)
         DIM wMinor AS WORD = LOWORD(pvsffi->dwFileVersionMS)
         FUNCTION = (wMajor + wMinor / 100) * 100
      END IF
   END IF
   HeapFree(GetProcessHeap, 0, pVerInfo)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version of CommCtl32.dll multiplied by 100, e.g. 600 for version 6.0.
' ========================================================================================
PRIVATE FUNCTION CControl.ComCtlVersion () AS LONG
   FUNCTION = this.GetFileVersion("COMCTL32.DLL")
END FUNCTION
' ========================================================================================

END NAMESPACE
