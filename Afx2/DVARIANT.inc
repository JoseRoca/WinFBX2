' ########################################################################################
' Platform: Microsoft Windows
' Filename: DVARIANT.inc
' Purpose:  Implements a VARIANT data type.
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
#include once "Afx2/DWSTRING.inc"

' ========================================================================================
' Macro for debug
' To allow debugging, define _DVARIANT_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _DVARIANT_DEBUG_
   #define _DVARIANT_DEBUG_ 0
#ENDIF
#ifndef _DVARIANT_DP_
   #define _DVARIANT_DP_ 1
   #MACRO DVARIANT_DP(st)
      #IF (_DVARIANT_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE Afx2

' ########################################################################################
'                                *** DVARIANT CLASS ***
' ########################################################################################
TYPE DVARIANT EXTENDS WSTRING

Public:
   vd AS VARIANT         ' // Variant data

Private:
   m_Result AS LONG      ' // Last result code
   m_dws AS DWSTRING     ' // To allow to return a WSTRING by reference

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE CONSTRUCTOR (BYREF dv AS DVARIANT)
   DECLARE CONSTRUCTOR (BYVAL v AS VARIANT)
   DECLARE CONSTRUCTOR (BYVAL pwsz AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYREF dws AS DWSTRING)
   DECLARE CONSTRUCTOR (BYREF bs AS BSTRING)
   DECLARE CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)

   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR CAST () BYREF AS WSTRING

   DECLARE OPERATOR LET (BYREF dv AS DVARIANT)
   DECLARE OPERATOR LET (BYVAL v AS VARIANT)
   DECLARE OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR LET (BYREF dws AS DWSTRING)
   DECLARE OPERATOR LET (BYREF bs AS BSTRING)
   DECLARE OPERATOR LET (BYVAL pvar AS VARIANT PTR)

   ' // Helper function
   DECLARE FUNCTION VariantToString (BYVAL pvarIn AS VARIANT PTR) AS DWSTRING

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' Usage example:
' DIM dv AS DVARIANT
' dv = "Test string"
' print dv
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT
   DVARIANT_DP("Default")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR DVARIANT
   DVARIANT_DP("")
   VariantClear(@vd)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes the DVARIANT from another DVARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYREF dv AS DVARIANT)
   DVARIANT_DP("DVARIANT")
   VariantCopy(@vd, @dv.vd)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DVARIANT from a VARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL v AS VARIANT)
   DVARIANT_DP("BYVAL VARIANT")
   VariantCopy(@vd, @v)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DVARIANT from a string.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL pwsz AS WSTRING PTR)
   DVARIANT_DP("WSTRING")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(pwsz)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYREF dws AS DWSTRING)
   DVARIANT_DP("DWSTRING")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(dws)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYREF bs AS BSTRING)
   DVARIANT_DP("BSTRING")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(bs)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a pointer to a VARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL pvar AS VARIANT PTR)
   DVARIANT_DP("VARIANT PTR")
   VariantCopy(@vd, pvar)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Assigns another DVARIANT.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYREF dv AS DVARIANT)
   DVARIANT_DP("DVARIANT")
   VariantClear(@vd) : VariantCopy(@vd, @dv.vd)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL v AS VARIANT)
   DVARIANT_DP("VARIANT")
   VariantClear(@vd) : VariantCopy(@vd, @v)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a VARIANT PTR.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL pvar AS VARIANT PTR)
   DVARIANT_DP("VARIANT PTR")
   VariantClear(@vd) : VariantCopy(@vd, pvar)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a WSTRING.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL pwszStr AS WSTRING PTR)
   DVARIANT_DP("WSTRING PTR")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(pwszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a DWSTRING.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYREF dws AS DWSTRING)
   DVARIANT_DP("DWSTRING")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(dws)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a BSTRING.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYREF bs AS BSTRING)
   DVARIANT_DP("BSTRING")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(bs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the variant data.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.CAST () AS VARIANT
   DVARIANT_DP("VARIANT")
   OPERATOR = vd
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the variant data.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.CAST () AS ANY PTR
   DVARIANT_DP("ANY PTR")
   OPERATOR = cast(ANY PTR, @vd)
END OPERATOR
' ========================================================================================

' =====================================================================================
' Extracts the contents of a VARIANT and returns them as a WSTRING.
' =====================================================================================
PRIVATE OPERATOR DVARIANT.CAST () BYREF AS WSTRING
   DVARIANT_DP("BYREF AS WSTRING")
   m_dws = this.VariantToString(@vd)
   OPERATOR = *CAST(WSTRING PTR, m_dws.GetBuffer)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Extracts the variant value of a variant structure to a string.
' Parameters:
' - pvarIn   : [in] Reference to a source variant structure.
' Return value:
' If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
' Usage: pRegExp.VariantToString(@variant_name)
' ========================================================================================
PRIVATE FUNCTION DVARIANT.VariantToString (BYVAL pvarIn AS VARIANT PTR) AS DWSTRING
   DVARIANT_DP("VARIANT PTR = " & ..WSTR(pvarin))
   DIM dws AS DWSTRING
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN RETURN ""
   DIM pVariantToStringAlloc AS FUNCTION (BYVAL pVar AS VARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS HRESULT
   pVariantToStringAlloc = DyLibSymbol(pLib, "VariantToStringAlloc")
   IF pVariantToStringAlloc = NULL THEN
      m_Result = E_FAIL
   ELSE
      DIM pwszBuf AS WSTRING PTR
      m_Result = pVariantToStringAlloc(pvarIn, @pwszBuf)
      IF pwszBuf THEN
         dws = *pwszBuf
         CoTaskMemFree pwszBuf
      END IF
   END IF
   DyLibFree(pLib)
   RETURN dws
END FUNCTION
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                         *** GLOBAL OPERATORS AND FUNCTIONS ***
' ########################################################################################

' // Outside a namespace because they are global
using Afx2

' ========================================================================================
' Returns the address of the underlying variant.
' One * returns the address of the underlying variant.
' Two ** deferences the variant data.
' Can be used to pass the variant to a BYVAL VARIANT parameter.
' ========================================================================================
PRIVATE OPERATOR * (BYREF dv AS DVARIANT) AS VARIANT PTR
   DVARIANT_DP("")
   OPERATOR = @dv.vd
END OPERATOR
' ========================================================================================
