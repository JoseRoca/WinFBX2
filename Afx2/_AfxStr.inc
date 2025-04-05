' ========================================================================================
' Converts an array of bytes into a formatted string.
' Note: Wrapped because, at the time of writing, it is not supported by the provided
' FreeBasic crypt32 import library.
' ========================================================================================
PRIVATE FUNCTION AfxCryptBinaryToStringA (BYVAL pbBinary AS CONST UBYTE PTR, BYVAL cbBinary AS DWORD, _
BYVAL dwFlags AS DWORD, BYVAL pszString AS LPSTR, BYVAL pcchString AS DWORD PTR) AS WINBOOL
   DIM AS ANY PTR pLib = DyLibLoad("crypt32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCryptBinaryToStringA AS FUNCTION (BYVAL pbBinary AS CONST UBYTE PTR, BYVAL cbBinary AS DWORD, _
       BYVAL dwFlags AS DWORD, BYVAL pszString AS LPSTR, BYVAL pcchString AS DWORD PTR) AS WINBOOL
   pCryptBinaryToStringA = DyLibSymbol(pLib, "CryptBinaryToStringA")
   IF pCryptBinaryToStringA THEN FUNCTION = pCryptBinaryToStringA(pbBinary, cbBinary, dwFlags, pszString, pcchString)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxCryptBinaryToStringW (BYVAL pbBinary AS CONST UBYTE PTR, BYVAL cbBinary AS DWORD, _
BYVAL dwFlags AS DWORD, BYVAL pszString AS LPWSTR, BYVAL pcchString AS DWORD PTR) AS WINBOOL
   DIM AS ANY PTR pLib = DyLibLoad("crypt32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCryptBinaryToStringW AS FUNCTION (BYVAL pbBinary AS CONST UBYTE PTR, BYVAL cbBinary AS DWORD, _
       BYVAL dwFlags AS DWORD, BYVAL pszString AS LPWSTR, BYVAL pcchString AS DWORD PTR) AS WINBOOL
   pCryptBinaryToStringW = DyLibSymbol(pLib, "CryptBinaryToStringW")
   IF pCryptBinaryToStringW THEN FUNCTION = pCryptBinaryToStringW(pbBinary, cbBinary, dwFlags, pszString, pcchString)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
#ifndef UNICODE
   #define AfxCryptBinaryToString(p1, p2, p3, p4, p5) AfxCryptBinaryToStringA(p1, p2, p3, p4, p5)
#else
   #define AfxCryptBinaryToString(p1, p2, p3, p4, p5) AfxCryptBinaryToStringW(p1, p2, p3, p4, p5)
#endif

' ========================================================================================
' Converts a formatted string into an array of bytes.
' ========================================================================================
PRIVATE FUNCTION AfxCryptStringToBinaryA (BYVAL pszString AS LPCSTR, BYVAL cchString AS DWORD, _
BYVAL dwFlags AS DWORD, BYVAL pbBinary AS UBYTE PTR, BYVAL pcbBinary AS DWORD PTR, _
BYVAL pdwSkip AS DWORD PTR, BYVAL pdwFlags AS DWORD PTR) AS WINBOOL
   DIM AS ANY PTR pLib = DyLibLoad("crypt32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCryptStringToBinaryA AS FUNCTION (BYVAL pszString AS LPCSTR, BYVAL cchString AS DWORD, _
       BYVAL dwFlags AS DWORD, BYVAL pbBinary AS UBYTE PTR, BYVAL pcbBinary AS DWORD PTR, _
       BYVAL pdwSkip AS DWORD PTR, BYVAL pdwFlags AS DWORD PTR) AS WINBOOL
   pCryptStringToBinaryA = DyLibSymbol(pLib, "CryptStringToBinaryA")
   IF pCryptStringToBinaryA THEN FUNCTION = pCryptStringToBinaryA(pszString, cchString, dwFlags, pbBinary, pcbBinary, pdwSkip, pdwFLags)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxCryptStringToBinaryW (BYVAL pwszString AS LPCWSTR, BYVAL cchString AS DWORD, _
BYVAL dwFlags AS DWORD, BYVAL pbBinary AS UBYTE PTR, BYVAL pcbBinary AS DWORD PTR, _
BYVAL pdwSkip AS DWORD PTR, BYVAL pdwFlags AS DWORD PTR) AS WINBOOL
   DIM AS ANY PTR pLib = DyLibLoad("crypt32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCryptStringToBinaryW AS FUNCTION (BYVAL pwszString AS LPCWSTR, BYVAL cchString AS DWORD, _
       BYVAL dwFlags AS DWORD, BYVAL pbBinary AS UBYTE PTR, BYVAL pcbBinary AS DWORD PTR, _
       BYVAL pdwSkip AS DWORD PTR, BYVAL pdwFlags AS DWORD PTR) AS WINBOOL
   pCryptStringToBinaryW = DyLibSymbol(pLib, "CryptStringToBinaryW")
   IF pCryptStringToBinaryW THEN FUNCTION = pCryptStringToBinaryW(pwszString, cchString, dwFlags, pbBinary, pcbBinary, pdwSkip, pdwFLags)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
#ifndef UNICODE
   #define AfxCryptStringToBinary(p1, p2, p3, p4, p5, p6, p7) AfxCryptStringToBinaryA(p1, p2, p3, p4, p5, p6, p7)
#else
   #define AfxCryptStringToBinary(p1, p2, p3, p4, p5, p6, p7) AfxCryptStringToBinaryW(p1, p2, p3, p4, p5, p6, p7)
#endif

' ========================================================================================
PRIVATE FUNCTION AfxBase64EncodeA (BYREF strData AS STRING) AS STRING
   DIM cchStr AS DWORD, strOut AS STRING
   DIM bRes AS LONG = AfxCryptBinaryToStringA(STRPTR(strData), LEN(strData), CRYPT_STRING_BASE64, NULL, @cchStr)
   IF bRes = 0 OR cchStr = 0 THEN EXIT FUNCTION
   strOut = SPACE(cchStr)
   bRes = AfxCryptBinaryToStringA(STRPTR(strData), LEN(strData), CRYPT_STRING_BASE64, STRPTR(strOut), @cchStr)
   FUNCTION = strOut
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxBase64EncodeW (BYREF dwsData AS DWSTRING) AS DWSTRING
   DIM cchStr AS DWORD, dwsOut AS DWSTRING
   DIM bRes AS LONG = AfxCryptBinaryToStringW(cwsData, LEN(cwsData) * 2, CRYPT_STRING_BASE64, NULL, @cchStr)
   IF bRes = 0 OR cchStr = 0 THEN EXIT FUNCTION
   dwsOut = WSPACE(cchStr)
   bRes = AfxCryptBinaryToStringW(cwsData, LEN(cwsData) * 2, CRYPT_STRING_BASE64, dwsOut, @cchStr)
   FUNCTION = dwsOut
END FUNCTION
' ========================================================================================
#ifndef UNICODE
   #define AfxBase64Encode(p) AfxBase64EncodeA(p)
#else
   #define AfxBase64Encode(p) AfxBase64EncodeW(p)
#endif

' ========================================================================================
PRIVATE FUNCTION AfxBase64DecodeA (BYREF strData AS STRING) AS STRING
   DIM cbBinary AS DWORD, strOut AS STRING
   DIM bRes AS LONG = AfxCryptStringToBinaryA(STRPTR(strData), LEN(strData), CRYPT_STRING_BASE64, NULL, @cbBinary, NULL, NULL)
   IF bRes = 0 OR cbBinary = 0 THEN EXIT FUNCTION
   strOut = SPACE(cbBinary)
   bRes = AfxCryptStringToBinaryA(STRPTR(strData), LEN(strData), CRYPT_STRING_BASE64, STRPTR(strOut), @cbBinary, NULL, NULL)
   FUNCTION = strOut
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxBase64DecodeW (BYREF dwsData AS DWSTRING) AS DWSTRING
   DIM cbBinary AS DWORD, dwsOut AS DWSTRING
   DIM bRes AS LONG = AfxCryptStringToBinaryW(dwsData, LEN(dwsData) * 2, CRYPT_STRING_BASE64, NULL, @cbBinary, NULL, NULL)
   IF bRes = 0 OR cbBinary = 0 THEN EXIT FUNCTION
   dwsOut = WSPACE(cbBinary \ 2)
   bRes = AfxCryptStringToBinaryW(cwsData, LEN(cwsData) * 2, CRYPT_STRING_BASE64, dwsOut, @cbBinary, NULL, NULL)
   FUNCTION = dwsOut
END FUNCTION
' ========================================================================================
#ifndef UNICODE
   #define AfxBase64Decode(p) AfxBase64DecodeA(p)
#else
   #define AfxBase64Decode(p) AfxBase64DecodeW(p)
#endif

' ========================================================================================
' Note: Although we can't write variadic functions with FreeBasic 64 bit (we can with 32
' bit only), we can call external variadic functions written in C, e.g.
'    DIM wszOut AS WSTRING * 260
'    DIM wszFmt AS WSTRING * 260 = "%s %d + %d = %d."
'    DIM wszText AS WSTRING * 260 = "The answer is"
'    DIM hr AS HRESULT = StringCbPrintfW(@wszOut, SIZEOF(wszOut), @wszFmt, @wszText, 1, 2, 3)
'    print wszOut
' Output: "The answer is 1 + 2 = 3."
' StringCbPrintf function:
' https://msdn.microsoft.com/en-us/library/windows/desktop/ms647510(v=vs.85).aspx
' StringCbPrintf_l (A/W) is similar to StringCbPrintf but includes a parameter for locale
' information.
' StringCbPrintfEx (A/W) adds to the functionality of StringCbPrintf by returning a pointer
' to the end of the destination string as well as the number of bytes left unused in that
' string. Flags may also be passed to the function for additional control.
' StringCbPrintf_lEx (A/W) is similar to StringCbPrintfEx but includes a parameter for
' locale information.
' They can be very useful to do string formatting.
' ========================================================================================

' The below procedures are placed here because AfxSetEnviron uses the AfxStrExtract procedure.

' ========================================================================================
' Retrieves the contents of the specified variable from the environment block of the
' calling process.
' - pwszName : The name of the environment variable.
' Return value: The contents of the specified environment variable.
' Example: DIM dws AS DWSTRING = AfxEnviron("path")
' ========================================================================================
PRIVATE FUNCTION AfxEnviron (BYVAL pwszName AS LPCWSTR) AS DWSTRING
   DIM wszBuffer AS WSTRING * 32767
   DIM cb AS DWORD = GetEnvironmentVariableW(pwszName, @wszBuffer, 32767)
   RETURN LEFT(wszBuffer, cb)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the contents of the specified environment variable for the current process.
' - pwszName : The name of the environment variable.
'              The operating system creates the environment variable if it does not exist
'              and pwszValue is not NULL.
' - pszValue : The contents of the environment variable.
'              The maximum size of a user-defined environment variable is 32,767 characters.
' Return value:
'   If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE.
'   To get extended error information, call GetLastError.
' Example: AfxSetEnviron("path", "c:")
' ========================================================================================
PRIVATE FUNCTION AfxSetEnviron OVERLOAD (BYVAL pwszName AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS BOOLEAN
   RETURN SetEnvironmentVariableW(pwszName, pwszValue)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the contents of the specified environment variable for the current process.
' - varexpr = Name and setting of an environment variable in the following (or equivalent)
'   form: varname=varstring. (varname being the name of the environment variable, and
'   varstring being its text value to set).
' Returns 0 on success, or -1 on failure.
' Example: AfxSetEnviron "path=c:"
' ========================================================================================
PRIVATE FUNCTION AfxSetEnviron OVERLOAD (BYREF varexp AS WSTRING) AS BOOLEAN
   DIM cwsName AS DWSTRING = AfxStrExtract(1, varexp, "=")
   IF LEN(cwsName) = 0 THEN RETURN TRUE
   dwsName = TRIM(dwsName)
   DIM dwsValue AS DWSTRING
   DIM p AS LONG = INSTR(varexp, "=")
   IF p = 0 THEN RETURN TRUE
   dwsValue = MID(varexp, p + 1)
   dwsValue = TRIM(dwsValue)
   IF LEN(dwsValue) = 0 THEN RETURN TRUE
   RETURN NOT SetEnvironmentVariable(dwsName, dwsValue)
END FUNCTION
' ========================================================================================
