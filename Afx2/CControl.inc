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

END NAMESPACE
