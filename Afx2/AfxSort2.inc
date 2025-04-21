' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxSort2.inc
' Purpose:  Sorting algorithms
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
#include once "Afx2/AfxArrays2.inc"
USING Afx2

NAMESPACE Afx2

' ========================================================================================
' Sorts a one-dimensional DWSTRING array calling the C qsort function.
' Digits in the strings are considered as numerical content rather than text. This test is
' not case-sensitive.
' Parameters:
' - rgwstr : Start of target array.
' - numElm : Number of elements in the array.
' - bAscend: TRUE for sorting in ascending order; FALSE for sorting in descending order.
' Example:
' #INCLUDE ONCE "Afx/DWSTRING.inc"
' USING Afx
' DIM rg(1 TO 9) AS DWSTRING
' rg(1) = "20string"
' rg(2) = "2string"
' rg(3) = "3string"
' rg(4) = "st20ring"
' rg(5) = "st2ring"
' rg(6) = "st3ring"
' rg(7) = "string2"
' rg(8) = "string20"
' rg(9) = "string3"
' print "---- after sorting ----"
' AfxDWSTRINGLogicalSort @rg(1), 9
' FOR i AS LONG = 1 TO 9
'   print rg(i)
' NEXT
' Output:
' 2string
' 3string
' 20string
' st2ring
' st3ring
' st20ring
' string2
' string3
' string20
' Note: &h00000008 is the value of the constant SORT_DIGITSASNUMBERS.
' Remarks:
' CompareStringW returns one of the following values if successful. To maintain the C runtime
' convention of comparing strings, the value 2 can be subtracted from a nonzero return value.
' Then, the meaning of <0, ==0, and >0 is consistent with the C runtime.
' CSTR_LESS_THAN. The string indicated by a is less in lexical value than the string indicated by b.
' CSTR_EQUAL. The string indicated by a is equivalent in lexical value to the string indicated by b.
' The two strings are equivalent for sorting purposes, although not necessarily identical.
' CSTR_GREATER_THAN. The string indicated by a is greater in lexical value than the string indicated by b.
' The function returns 0 if it does not succeed. To get extended error information, the application
' can call GetLastError, which can return one of the following error codes:
' ERROR_INVALID_FLAGS. The values supplied for flags were invalid.
' ERROR_INVALID_PARAMETER. Any of the parameter values was invalid.
' ========================================================================================
' st20ring
' string2
' string3
' string20
' Note: &h00000008 is the value of the constant SORT_DIGITSASNUMBERS.
' Remarks:
' CompareStringW returns one of the following values if successful. To maintain the C runtime
' convention of comparing strings, the value 2 can be subtracted from a nonzero return value.
' Then, the meaning of <0, ==0, and >0 is consistent with the C runtime.
' CSTR_LESS_THAN. The string indicated by a is less in lexical value than the string indicated by b.
' CSTR_EQUAL. The string indicated by a is equivalent in lexical value to the string indicated by b.
' The two strings are equivalent for sorting purposes, although not necessarily identical.
' CSTR_GREATER_THAN. The string indicated by a is greater in lexical value than the string indicated by b.
' The function returns 0 if it does not succeed. To get extended error information, the application
' can call GetLastError, which can return one of the following error codes:
' ERROR_INVALID_FLAGS. The values supplied for flags were invalid.
' ERROR_INVALID_PARAMETER. Any of the parameter values was invalid.
' ========================================================================================

' ########################################################################################
'                                     DWSTRING
' ########################################################################################

' qsort DWSTRING logical comparison function
PRIVATE FUNCTION DWStringArrayLogicalCompare CDECL (BYVAL a AS DWSTRING PTR, BYVAL b AS DWSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_pBuffer), a->m_BufferLen, cast(WSTRING PTR, b->m_pBuffer), b->m_BufferLen, NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   RETURN r
END FUNCTION

' Reverse qsort logical DWSTRING comparison function
PRIVATE FUNCTION DWStringArrayReverseLogicalCompare CDECL (BYVAL a AS DWSTRING PTR, BYVAL b AS DWSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_pBuffer), a->m_BufferLen, cast(WSTRING PTR, b->m_pBuffer), b->m_BufferLen, NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

' qsort DWSTRING logical sort function
PRIVATE SUB DWStringArrayLogicalSort OVERLOAD (BYVAL rgdwstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgdwstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgdwstr, numElm, SIZEOF(DWSTRING), CPTR(ANY PTR, @DWStringArrayLogicalCompare)
   ELSE
      qsort rgdwstr, numElm, SIZEOF(DWSTRING) , CPTR(ANY PTR, @DWStringArrayReverseLogicalCompare)
   END IF
END SUB

' qsort DWSTRING logical sort function
PRIVATE SUB DWStringArrayLogicalSort OVERLOAD (rgwstr() AS DWSTRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgwstr) - LBOUND(rgwstr) + 1
   DWStringArrayLogicalSort @rgwstr(LBOUND(rgwstr)), numElm, bAscend
END SUB

' ########################################################################################
'                                      BSTRING
' ########################################################################################

' qsort BSTRING logical comparison function
PRIVATE FUNCTION BStringArrayLogicalCompare CDECL (BYVAL a AS BSTRING PTR, BYVAL b AS BSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_bstr), SysStringLen(a->m_bstr), cast(WSTRING PTR, b->m_bstr), SysStringLen(b->m_bstr), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   RETURN r
END FUNCTION

' Reverse qsort logical BSTRING comparison function
PRIVATE FUNCTION BStringArrayReverseLogicalCompare CDECL (BYVAL a AS BSTRING PTR, BYVAL b AS BSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_bstr), SysStringLen(a->m_bstr), cast(WSTRING PTR, b->m_bstr), SysStringLen(b->m_bstr), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

' qsort BSTRING logical sort function
PRIVATE SUB BStringArrayLogicalSort OVERLOAD (BYVAL rgbstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgbstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgbstr, numElm, SIZEOF(BSTRING), CPTR(ANY PTR, @BStringArrayLogicalCompare)
   ELSE
      qsort rgbstr, numElm, SIZEOF(BSTRING) , CPTR(ANY PTR, @BStringArrayReverseLogicalCompare)
   END IF
END SUB

' qsort DWSTRING logical sort function
PRIVATE SUB BStringArrayLogicalSort OVERLOAD (rgbstr() AS BSTRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgbstr) - LBOUND(rgbstr) + 1
   BStringArrayLogicalSort @rgbstr(LBOUND(rgbstr)), numElm, bAscend
END SUB

' ########################################################################################
'                                      STRING
' ########################################################################################

' qsort STRING comparison function
PRIVATE FUNCTION StringArrayCompare CDECL (BYVAL a AS STRING PTR, BYVAL b AS STRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, STRPTR(*a), LEN(*a), STRPTR(*b), LEN(*b))
   IF r <> 0 THEN r = r - 2
   RETURN r
END FUNCTION

' Reverse qsort STRING comparison function
PRIVATE FUNCTION StringArrayReverseCompare CDECL (BYVAL a AS STRING PTR, BYVAL b AS STRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, STRPTR(*a), LEN(*a), STRPTR(*b), LEN(*b))
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

' qsort STRING sort function
PRIVATE SUB StringArraySort OVERLOAD (BYVAL rgstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgstr, numElm, SIZEOF(STRING), CPTR(ANY PTR, @StringArrayCompare)
   ELSE
      qsort rgstr, numElm, SIZEOF(STRING) , CPTR(ANY PTR, @StringArrayReverseCompare)
   END IF
END SUB

' qsort STRING sort function
PRIVATE SUB StringArraySort OVERLOAD (rgstr() AS STRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgstr) - LBOUND(rgstr) + 1
   StringArraySort @rgstr(LBOUND(rgstr)), numElm, bAscend
END SUB

' ########################################################################################
'                                      ZSTRING
' ########################################################################################

' qsort ZSTRING comparison function
PRIVATE FUNCTION ZStringArrayCompare CDECL (BYVAL a AS ZSTRING PTR, BYVAL b AS ZSTRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b))
   IF r <> 0 THEN r = r - 2
   RETURN r
END FUNCTION

' Reverse qsort ZSTRING comparison function
PRIVATE FUNCTION ZStringArrayReverseCompare CDECL (BYVAL a AS ZSTRING PTR, BYVAL b AS ZSTRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b))
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

' qsort ZSTRING sort function
PRIVATE SUB ZStringArraySort OVERLOAD (BYVAL rgzstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL sizeElem AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF bAscend THEN
      qsort rgzstr, numElm, sizeElem, CPTR(ANY PTR, @ZStringArrayCompare)
   ELSE
      qsort rgzstr, numElm, sizeElem, CPTR(ANY PTR, @ZStringArrayReverseCompare)
   END IF
END SUB

' ########################################################################################
'                                      WSTRING
' ########################################################################################

' qsort WSTRING comparison function
PRIVATE FUNCTION WStringArrayCompare CDECL (BYVAL a AS WSTRING PTR, BYVAL b AS WSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   RETURN r
END FUNCTION

' Reverse qsort WSTRING comparison function
PRIVATE FUNCTION WStringArrayReverseCompare CDECL (BYVAL a AS WSTRING PTR, BYVAL b AS WSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

' qsort WSTRING sort function
PRIVATE SUB WStringArraySort OVERLOAD (BYVAL rgwstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL sizeElem AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF bAscend THEN
      qsort rgwstr, numElm, sizeElem, CPTR(ANY PTR, @WStringArrayCompare)
   ELSE
      qsort rgwstr, numElm, sizeElem, CPTR(ANY PTR, @WStringArrayReverseCompare)
   END IF
END SUB

' ########################################################################################
'                                      BYTE
' ########################################################################################

PRIVATE FUNCTION ByteArraySortCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Byte PTR, a) > *CPTR(Byte PTR, b) THEN RETURN 1
   IF *CPTR(Byte PTR, a) < *CPTR(Byte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

PRIVATE FUNCTION ByteArraySortReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Byte PTR, a) < *CPTR(Byte PTR, b) THEN RETURN 1
   IF *CPTR(Byte PTR, a) > *CPTR(Byte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
PRIVATE SUB ByteArraySort OVERLOAD (BYVAL rgbyte AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgbyte = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgbyte, numElm, SIZEOF(BYTE), CPTR(ANY PTR, @ByteArraySortCompare)
   ELSE
      qsort rgbyte, numElm, SIZEOF(BYTE), CPTR(ANY PTR, @ByteArraySortReverseCompare)
   END IF
END SUB

' qsort for numbers
PRIVATE SUB ByteArraySort OVERLOAD (rgbyte () AS BYTE, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgbyte) - LBOUND(rgbyte) + 1
   ByteArraySort @rgbyte(LBOUND(rgbyte)), numElm, bAscend
END SUB

' ########################################################################################
'                                      UBYTE
' ########################################################################################

PRIVATE FUNCTION UByteArraySortCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UByte PTR, a) > *CPTR(UByte PTR, b) THEN RETURN 1
   IF *CPTR(UByte PTR, a) < *CPTR(UByte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

PRIVATE FUNCTION UByteArraySortReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UByte PTR, a) < *CPTR(UByte PTR, b) THEN RETURN 1
   IF *CPTR(UByte PTR, a) > *CPTR(UByte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
PRIVATE SUB UByteArraySort OVERLOAD (BYVAL rgubyte AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgubyte = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgubyte, numElm, SIZEOF(UBYTE), CPTR(ANY PTR, @UByteArraySortCompare)
   ELSE
      qsort rgubyte, numElm, SIZEOF(UBYTE), CPTR(ANY PTR, @UByteArraySortReverseCompare)
   END IF
END SUB

' qsort for numbers
PRIVATE SUB UByteArraySort OVERLOAD (rgubyte () AS UBYTE, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgubyte) - LBOUND(rgubyte) + 1
   UByteArraySort @rgubyte(LBOUND(rgubyte)), numElm, bAscend
END SUB

' ########################################################################################
'                                      SHORT
' ########################################################################################

PRIVATE FUNCTION ShortArraySortCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Short PTR, a) > *CPTR(Short PTR, b) THEN RETURN 1
   IF *CPTR(Short PTR, a) < *CPTR(Short PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

PRIVATE FUNCTION ShortArraySortReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Short PTR, a) < *CPTR(Short PTR, b) THEN RETURN 1
   IF *CPTR(Short PTR, a) > *CPTR(Short PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
PRIVATE SUB ShortArraySort OVERLOAD (BYVAL rgshort AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgshort = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgshort, numElm, SIZEOF(SHORT), CPTR(ANY PTR, @ShortArraySortCompare)
   ELSE
      qsort rgshort, numElm, SIZEOF(SHORT), CPTR(ANY PTR, @ShortArraySortReverseCompare)
   END IF
END SUB

' qsort for numbers
PRIVATE SUB ShortArraySort OVERLOAD (rgshort () AS SHORT, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgshort) - LBOUND(rgshort) + 1
   ShortArraySort @rgshort(LBOUND(rgshort)), numElm, bAscend
END SUB

' ########################################################################################
'                                      USHORT
' ########################################################################################

PRIVATE FUNCTION UShortArraySortCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UShort PTR, a) > *CPTR(UShort PTR, b) THEN RETURN 1
   IF *CPTR(UShort PTR, a) < *CPTR(UShort PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

PRIVATE FUNCTION UShortArraySortReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UShort PTR, a) < *CPTR(UShort PTR, b) THEN RETURN 1
   IF *CPTR(UShort PTR, a) > *CPTR(UShort PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
PRIVATE SUB UShortArraySort OVERLOAD (BYVAL rgUShort AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgUShort = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgUShort, numElm, SIZEOF(USHORT), CPTR(ANY PTR, @UShortArraySortCompare)
   ELSE
      qsort rgUShort, numElm, SIZEOF(USHORT), CPTR(ANY PTR, @UShortArraySortReverseCompare)
   END IF
END SUB

' qsort for numbers
PRIVATE SUB UShortArraySort OVERLOAD (rgUShort () AS UShort, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgUShort) - LBOUND(rgUShort) + 1
   UShortArraySort @rgUShort(LBOUND(rgUShort)), numElm, bAscend
END SUB

END NAMESPACE

