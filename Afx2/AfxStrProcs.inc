' ########################################################################################
' Microsoft Windows
' File: AfxStrProcs.inc
' Contents: Windows High DPI wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "Afx2/CRegExp2.inc"
USING Afx2

NAMESPACE Afx2

' ========================================================================================
' * Returns a string containing only the characters contained in a specified match string.
' If wszMatchString is an empty string, strRetain returns an empty string.
' ========================================================================================
' Usage examples:
 ' DIM dws AS DWSTRING = "abacadabra"
 ' dws = strRetain(dws, "B", TRUE)
 ' print dws
' Output: "bb"
' ----------------------------------------------------------------------------------------
' DIM dws AS DWSTRING = "<p>1234567890<ak;lk;l>1234567890</p>"
' dws = strRetain(dws, "<;/p>", TRUE)
' print dws
' Output: "<p><;;></p>"
' ----------------------------------------------------------------------------------------
' DIM dws AS DWSTRING = "<p>1234567890<ak;lk;l>1234567890</p>"
' dws = strRetain(dws, "0123456789", TRUE)
' print dws
' Output: "12345678901234567890"
' ========================================================================================
FUNCTION strRetain (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
   DIM pRegExp AS CRegExp
   pRegExp.Global = TRUE
   pRegExp.IgnoreCase = IgnoreCase
   pRegExp.Pattern = "[^" & wszMatchString & "]"
   RETURN STRPTR(pRegExp.Remove(wszSourceString))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the portion of a string following the first occurrence of a character or group of characters.
' Example:
' DIM dws AS DWSTRING = "I think, therefore I am"
' dws = strRemain(dws, ",", TRUE)
' print dws
' Output: " therefore I am"
' ========================================================================================
FUNCTION strRemain OVERLOAD (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
   DIM pRegExp AS CRegExp
   pRegExp.Global = FALSE
   pRegExp.IgnoreCase = IgnoreCase
   pRegExp.Pattern = wszMatchString & "(.*)"
   IF pRegExp.Execute(wszSourceString) THEN
      iF pRegExp.SubMatchesCount THEN
         RETURN STRPTR(pRegExp.SubMatchValue(0, 0))
      END IF
   END IF
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' * nStart : Starting position to begin the search. If nStart is not specified, the search
' will begin at position 1. If nStart is zero, a nul string is returned. If nStart is
' negative, the starting position is counted from right to left: -1 for the last character,
' -2 for the second to last, etc.
' Example:
' DIM dws AS DWSTRING = "I think, therefore I am"
' dws = strRemain(5, dws, ",", TRUE)
' print dws
' Output: " therefore I am"
' ========================================================================================
FUNCTION strRemain OVERLOAD (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
   DIM pRegExp AS CRegExp
   pRegExp.Global = FALSE
   pRegExp.IgnoreCase = IgnoreCase
   pRegExp.Pattern = wszMatchString & "(.*)"
   DIM bsSrc AS BSTRING
   IF nStart < 0 THEN nStart = LEN(wszSourceString) + nStart + 1
   IF nStart < 0 OR nStart > LEN(wszSourceString) THEN RETURN ""
   bsSrc = MID(wszSourceString, nStart)
   IF pRegExp.Execute(bsSrc) THEN
      iF pRegExp.SubMatchesCount THEN
         RETURN STRPTR(pRegExp.SubMatchValue(0, 0))
      END IF
   END IF
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Count the number of occurrences of a string or a list of characters within a string.
' Usage:
'    DIM dws AS DWSTRING = "abacadabra"
'    DIM nCount AS LONG = strTally(dw, "bac", TRUE)
'    PRINT nCount
' Output: Returns 1, counting the string "bac"
' ----------------------------------------------------------------------------------------
'    DIM dws AS DWSTRING = "abacadabra"
'    DIM nCount AS LONG = strTally(dws, "b|a|c", TRUE)
'    PRINT nCount
' Output: Returns 8, counting all "b", "a", and "c" characters.
' The | is the OR operator in regular expressions. It means "match b OR a OR c".
' ========================================================================================
PRIVATE FUNCTION strTally (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
   DIM pRegExp AS CRegExp
   pRegExp.Global = TRUE
   pRegExp.IgnoreCase = IgnoreCase
   pRegExp.Pattern = wszMatchString
   IF pRegExp.Execute(wszSourceString) THEN
      RETURN pRegExp.MatchesCount
   END IF
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Find the position of the first character in the string that does not match in the match string.
' strVerify returns zero if each character in wszSourceString is present in wszMatchString.
' If not, it returns the position of the first non-matching character in wszSourceString.
' Example:
'    DIM dws AS DWSTRING = "123.65,22.5"
'    DIM nPos AS LONG = strVerify(5, dws, "0123456789", TRUE)
'    PRINT nPos
' Output: 7.
' Returns 7 since 5 starts it past the first non-digit ("." at position 4)
' ========================================================================================
PRIVATE FUNCTION strVerify (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
   IF nStart <= 0 OR nStart > LEN(wszSourceString) - 1 THEN RETURN 0
   DIM pRegExp AS CRegExp
   pRegExp.Global = TRUE
   pRegExp.IgnoreCase = IgnoreCase
   pRegExp.Pattern = "[^" & wszMatchString & "]"
   DIM bsSrc AS BSTRING = MID(wszSourceString, nStart)
   IF pRegExp.Execute(bsSrc) THEN
      RETURN pRegExp.MatchPos + nStart - 1
   END IF
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed from the left, right or mid side of the string.
' - wszSide: "LEFT", "MID" or "RIGHT".
' - nCount: Number or characters to remove.
' - nStart: (only MID) Starting position.
' Example: DIM dws AS DWSTRING = strClip("LEFT", "1234567890", 3)     ' Output: "4567890"
' Example: DIM dws AS DWSTRING = strClip("RIGHT", "1234567890", 3)    ' Output: "1234567"
' Example: DIM dws AS DWSTRING = strClip("MID", "1234567890", 3, 4)   ' Output: =1237890"
' ========================================================================================
PRIVATE FUNCTION strClip (BYREF wszSide AS CONST WSTRING, BYREF wszSourceString AS CONST WSTRING, BYVAL nCount AS LONG, BYVAL nStart AS LONG = 0) AS DWSTRING
   DIM dws AS DWSTRING = wszSourceString
   IF nCount <= 0 THEN RETURN dws
   DIM nLen AS LONG = LEN(dws)
   SELECT CASE UCASE(wszSide)
      CASE "LEFT"
         nCount = IIF(nLen < nCount, nLen, nCount)
         dws = MID(wszSourceString, nCount + 1)
         RETURN dws
      CASE "RIGHT"
         nCount = nLen - nCount
         nCount = IIF(nLen < nCount, nLen, nCount)
         dws = LEFT(wszSourceString, nCount)
         RETURN dws
      CASE "MID"
         IF nStart <= 0 THEN RETURN dws
         dws = LEFT(wszSourceString, nStart - 1) + MID(wszSourceString, nStart + nCount)
         RETURN dws
   END SELECT
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified number of characters from a string expression.
' Returns a string based on wszMainStr but with nCount characters deleted
' starting at position nStart. The first character in the string is position 1, etc.
' Usage example:
' DIM dws AS DWSTRING = strDelete("1234567890", 4, 3)   ' Output: 1237890
' ========================================================================================
PRIVATE FUNCTION strDelete (BYREF wszSourceString AS CONST WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS DWSTRING
   DIM nLen AS LONG = LEN(wszSourceString)
   IF nLen = 0 OR nStart <= 0 OR nCount <= 0 OR nStart > nLen THEN RETURN wszSourceString
   RETURN LEFT(wszSourceString, nStart - 1) + MID(wszSourceString, nStart + nCount)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retuns TRUE if the passed string is muneric.
' Example: PRINT strIsNumeric("1.2345678901234567e+029")   ' Output: true
' ========================================================================================
PRIVATE FUNCTION strIsNumeric (BYREF wszSourcestring AS CONST WSTRING) AS BOOLEAN
   DIM pRegExp AS CRegExp
   pRegExp.Pattern = $"^[\+\-]?\d*\.?\d+(?:[Ee][\+\-]?\d+)?$"
   RETURN pRegExp.Test(wszSourcestring)
END FUNCTION
' ========================================================================================

' ========================================================================================
' ' * Shrinks a string to use a consistent single character delimiter.
' The purpose of this function is to create a string with consecutive data items (words)
' separated by a consistent single character. This makes it very straightforward to parse
' the results as needed.
' If wszMask is not defined then all leading spaces and trailing spaces are removed entirely.
' All occurrences of two or more spaces are changed to a single space. Therefore, the new
' string returned consists of zero or more words, each separated by a single space character.
' If wszMask is specified, it defines one or more delimiter characters to shrink. All leading
' and trailing mask characters are removed entirely. All occurrences of one or more mask
' characters are replaced with the first character of wszMask The new string returned consists
' of zero or more words, each separated by the character found in the first position of wszMask.
' Example:
'    DIM dws AS BSTRING = strShrink(",,, one , two     three, four,", " ,")
' Output: "one two three four".
' ========================================================================================
PRIVATE FUNCTION strShrink(BYREF wszSourceString AS CONST WSTRING, BYREF wszMask AS CONST WSTRING = "", BYVAL IgnoreCase AS LONG = TRUE) AS DWSTRING
   DIM dwsText AS DWSTRING = wszSourceString
   DIM pRegExp AS CRegExp
   pRegExp.Global = TRUE
   pRegExp.IgnoreCase = IgnoreCase
   IF wszMask = "" THEN
      ' // Change all occurrences of two or more spaces to a single space
      pRegExp.Pattern = "[ ]+"
      dwsText = STRPTR(pRegExp.Replace(dwsText, " "))
   ELSE
      ' // Remove the passed characters
      pRegExp.Pattern = "[" & wszMask & "]+"
      dwsText = STRPTR(pRegExp.Replace(dwsText, LEFT(wszMask, 1)))
   END IF
   ' // Remove all leading and trailing spaces
   pRegExp.Pattern = $"^\s+|\s+$"
   dwsText = STRPTR(pRegExp.Replace(dwsText, ""))
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Reverses the contents of a string expression.
' Usage example: DIM dws AS DWSTRING = strReverse("garden")   ' Output: "nedrag"
' ========================================================================================
PRIVATE FUNCTION strReverse (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   DIM dwsSourceString AS DWSTRING = wszSourceString
   DIM wszChar AS WSTRING * 2
   DIM nLen AS LONG = LEN(wszSourceString)
   FOR i AS LONG = 1 TO nLen \ 2
      wszChar = MID(dwsSourceString, i, 1)
      MID(dwsSourceString, i, 1) = MID(dwsSourceString, nLen - i + 1, 1)
      MID(dwsSourceString, nLen - i + 1, 1) = wszChar
   NEXT
   RETURN dwsSourceString
END FUNCTION
' ========================================================================================

' ========================================================================================
' Splits a string into tokens, which are sequences of contiguous characters separated by
' any of the characters that are part of delimiters. Each token is added to a DWSTRING
' (my own dynamic unicode string data type for Free Basic) and delimited by a carriage
' return and line feed. The returned string will be parsed later to get the individual tokens.
' Instead of DWSTRING (which is not an intrinsic Free Basic data type) a buffer could be used.
' wcstok is a C function.
' Parameters:
' - wszStr = The string to split.
' - wszDelimiters = The delimiter characters.
' Return value: A CRLF separated list of tokens to be parsed with strParse.
' Usage:
' DIM wsz AS WSTRING * 260 = "- This, a sample string."
' DIM dwsTokens AS DWSTRING = strSplit(wsz, " ,.-")
' print "len dwsTokens: ", len(dwsTokens)
' print dwsTokens
' ========================================================================================
FUNCTION strSplit (BYREF wszStr AS WSTRING, BYREF wszDelimiters AS WSTRING) AS DWSTRING
   DIM dws AS DWSTRING = wszStr
   DIM dwsRes AS DWSTRING
   DIM pwsz AS WSTRING PTR = wcstok(dws, @wszDelimiters)
   WHILE pwsz <> NULL
      dwsRes += *pwsz
      dwsRes += CHR(13, 10)
      pwsz = wcstok(NULL, @wszDelimiters)
   WEND
   ' // Remove the last CRLF
   RETURN LEFT(dwsRes, LEN(dwsRes) - 2)
END FUNCTION
' ========================================================================================

END NAMESPACE

