' ########################################################################################
' Platform: Microsoft Windows
' Filename: DWSTRING.inc
' Purpose:  Implements a data type for dynamic null terminated unicode strings.
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

' // Include files
#pragma ONCE
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "/crt/string.bi"
#INCLUDE ONCE "/crt/wchar.bi"
#INCLUDE ONCE "utf_conv.bi"
#include once "win/Ole2.bi"

' ========================================================================================
' Macro for debug
' To allow debugging, define _DWSTRING_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _DWSTRING_DEBUG_
   #define _DWSTRING_DEBUG_ 0
#ENDIF
#ifndef _DWSTRING_DP_
   #define _DWSTRING_DP_ 1
   #MACRO DWSTRING_DP(st)
      #IF (_DWSTRING_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' Macro for debug
' To allow debugging, define _BSTRING_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _BSTRING_DEBUG_
   #define _BSTRING_DEBUG_ 0
#ENDIF
#ifndef _BSTRING_DP_
   #define _BSTRING_DP_ 1
   #MACRO BSTRING_DP(st)
      #IF (_BSTRING_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE Afx2

' // Forward reference declarations
TYPE DWSTRING_ AS DWSTRING
TYPE BSTRING_ AS BSTRING

' ########################################################################################
'                                *** DWSTRING CLASS ***
' ########################################################################################

TYPE DWSTRING EXTENDS WSTRING

   ' // Don't change the order of these variables
   m_pBuffer AS WSTRING PTR   ' // Pointer to the buffer
   m_BufferLen AS ssize_t     ' // Length in UTF16 characters of the buffer
   m_Capacity AS ssize_t      ' // The total size of the buffer in UTF16 characters

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   DECLARE CONSTRUCTOR (BYREF dws AS DWSTRING)
   DECLARE CONSTRUCTOR (BYREF bs AS BSTRING_)
   DECLARE CONSTRUCTOR (BYVAL nChars AS LONG, BYREF wszFill AS CONST WSTRING)
   DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
   DECLARE DESTRUCTOR
   DECLARE PROPERTY Capacity () AS UINT
   DECLARE PROPERTY Capacity (BYVAL nValue AS UINT)
   DECLARE FUNCTION Add (BYVAL pwszStr AS WSTRING PTR) AS BOOLEAN
   DECLARE FUNCTION Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   DECLARE FUNCTION Add (BYREF dws AS DWSTRING) AS BOOLEAN
   DECLARE FUNCTION ResizeBuffer (BYVAL nChars AS UINT, BYVAL bClear AS BOOLEAN = FALSE) AS WSTRING PTR
   DECLARE FUNCTION AppendBuffer (BYVAL memAddr AS ANY PTR, BYVAL nChars AS UINT) AS BOOLEAN
   DECLARE FUNCTION GetBuffer () AS WSTRING PTR
   DECLARE SUB Clear
   DECLARE OPERATOR [] (BYVAL nIndex AS UINT) BYREF AS USHORT
   DECLARE OPERATOR CAST () BYREF AS WSTRING
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR LET (BYREF ansiStr AS STRING)
   DECLARE OPERATOR LET (BYREF dws AS DWSTRING)
   DECLARE OPERATOR LET (BYREF bs AS BSTRING_)
   DECLARE OPERATOR LET (BYVAL n AS LONGINT)
   DECLARE OPERATOR LET (BYVAL n AS DOUBLE)
   DECLARE OPERATOR += (BYREF wszStr AS WSTRING)
   DECLARE OPERATOR += (BYREF dws AS DWSTRING)
   DECLARE OPERATOR += (BYREF bs AS BSTRING_)
   DECLARE OPERATOR += (BYREF ansiStr AS STRING)
   DECLARE OPERATOR += (BYVAL n AS LONGINT)
   DECLARE OPERATOR += (BYVAL n AS DOUBLE)
   DECLARE OPERATOR &= (BYREF wszStr AS WSTRING)
   DECLARE OPERATOR &= (BYREF dws AS DWSTRING)
   DECLARE OPERATOR &= (BYREF bs AS BSTRING_)
   DECLARE OPERATOR &= (BYREF ansiStr AS STRING)
   DECLARE OPERATOR &= (BYVAL n AS LONGINT)
   DECLARE OPERATOR &= (BYVAL n AS DOUBLE)
   DECLARE FUNCTION ChrW (BYVAL codepoint AS UInteger) AS DWSTRING
   DECLARE FUNCTION bstr () AS ..BSTR
   DECLARE PROPERTY utf8 () AS STRING
   DECLARE PROPERTY utf8 (BYREF utf8String AS STRING)
   DECLARE FUNCTION wchar () AS WSTRING PTR

Private:
   DECLARE FUNCTION ScanForSurrogates (BYVAL memAddr AS WSTRING PTR, BYVAL nChars AS LONG, BYVAL searchBrokenOnly AS BOOLEAN = TRUE) AS LONG
   
END TYPE
' ########################################################################################

' ########################################################################################
'                                *** BSTRING CLASS ***
' ########################################################################################
TYPE BSTRING EXTENDS WSTRING

Public:
   m_bstr AS BSTR   ' // The BSTR handle

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF bs AS BSTRING)
   DECLARE CONSTRUCTOR (BYREF dws AS DWSTRING)
'   DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR, BYVAL fAttach AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS UINT = 0)
   DECLARE CONSTRUCTOR (BYVAL nChars AS LONG, BYREF wszFill AS CONST WSTRING)
   DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
   DECLARE DESTRUCTOR
   DECLARE SUB Append (BYVAL pwszStr AS WSTRING PTR)
   DECLARE SUB Clear
   DECLARE SUB Attach (BYVAL pbstrSrc AS BSTR)
   DECLARE FUNCTION Detach () AS BSTRING
   DECLARE FUNCTION Copy () AS BSTRING
   DECLARE OPERATOR [] (BYVAL nIndex AS UINT) BYREF AS USHORT
   DECLARE OPERATOR CAST () BYREF AS CONST WSTRING
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR LET (BYREF bs AS BSTRING)
   DECLARE OPERATOR LET (BYREF dws AS DWSTRING_)
   DECLARE OPERATOR LET (BYREF s AS STRING)
'   DECLARE OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR LET (BYVAL n AS LONGINT)
   DECLARE OPERATOR LET (BYVAL n AS DOUBLE)
   DECLARE OPERATOR += (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR += (BYREF bs AS BSTRING)
   DECLARE OPERATOR += (BYREF dws AS DWSTRING_)
   DECLARE OPERATOR += (BYVAL n AS LONGINT)
   DECLARE OPERATOR += (BYVAL n AS DOUBLE)
   DECLARE OPERATOR &= (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR &= (BYREF bs AS BSTRING)
   DECLARE OPERATOR &= (BYREF dws AS DWSTRING_)
   DECLARE OPERATOR &= (BYVAL n AS LONGINT)
   DECLARE OPERATOR &= (BYVAL n AS DOUBLE)
   DECLARE PROPERTY Utf8 () AS STRING
   DECLARE PROPERTY Utf8 (BYREF utf8String AS STRING)
   DECLARE FUNCTION wchar () AS WSTRING PTR

END TYPE
' ========================================================================================

END NAMESPACE
