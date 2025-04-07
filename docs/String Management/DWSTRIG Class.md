# DWSTRING Class

The `DWSTRING` class implements a dynamic unicode null terminated string. Free Basic has a dynamic string data type (STRING) and a fixed length unicode data type (WSTRING), but it lacks a dynamic unicode string. `DWSTRING` behaves as if it was a native data type, working transparently with the intrinsic Free Basic string functions and operators.

**Include file**: DWSTRING.INC.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructors) | Initialize the class with the specified value. |
| [Operator \*](#Operator*) | One * returns the address of the `DWsTRING` buffer.<br> Two ** returns the address of the start of the string data. |
| [Operator &](#Operator&) | Concatenates strings. |
| [Operator +=](#Operator+=) | Appends a string to the CWSTR. |
| [Operator &=](#Operator&=) | Appends a string to the CWSTR. |
| [Operator []](#Operator[]) | Gets or sets the corresponding unicode integer representation of the character at the specified position. |
| [Operator Let](#OperatorLet) | Assigns a string to the CWSTR. It implements the = operator. |
| [Operator Cast](#OperatorCast) | Returns a pointer to the CWSTR buffer or the string data.<br>Casting is automatic. You don't have to call this operator. |
| [bstr](#bstr) | Returns the contents of the CWSTR as a BSTR. |
| [cbstr](#cbstr) | Returns the contents of the CWSTR as a CBSTR. |
| [wchar](#wchar) | Returns the string data as a new unicode string allocated with CoTaskMemAlloc. |
| [Utf8](#Utf8) | Converts from UTF8 to Unicode and from Unicode to UTF8. |
| [Capacity](#Capacity) | Gets/sets the size of the internal buffer. |

# <a name="constructors"></a>Constructors

```
CONSTRUCTOR DWSTRING
```
Creates an empty Unicode string buffer with an initial null-terminated string.
```
CONSTRUCTOR DWSTRING (BYVAL pwszStr AS WSTRING PTR)
```
Initializes a `DWSTRING` instance from a wide string pointer.
Example: DIM dws AS DWSTRING
```
CONSTRUCTOR DWSTRING (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
```
Initializes the string from an ANSI or UTF-8 encoded string, with optional code page support.
Example: DIM dws AS DWSTRING = "Hello, world"
Example: DIM dws AS DWSTRING = DWSTRING("Hello, utf", CP_UTF8)
For a list of code pages see: [Code Page Identifiers](https://msdn.microsoft.com/en-us/library/windows/desktop/dd317756(v=vs.85).aspx)
```
CONSTRUCTOR DWSTRING (BYREF dws AS DWSTRING)
```
Creates a copy of an existing DWSTRING.
Example: DIM dws1 AS DWSTRING = "Test string" : DIM dws2 AS DWSTRING = dws1
```
CONSTRUCTOR DWSTRING (BYVAL nChars AS LONG, BYREF wszFill AS CONST WSTRING)
```
Creates a DWSTRING with a fixed-length buffer, initialized with a fill character.
Example: DIM dws AS DWSTRING = DWSTRING(260, " ")
```
CONSTRUCTOR DWSTRING (BYVAL n AS LONGINT)
CONSTRUCTOR (BYVAL n AS DOUBLE)
```
Initializes a DWSTRING with a numeric value, allowing automatic conversion.
Example: DIM dwsNum AS DWSTRING= 12345
Example: DIM dwsFloat AS DWSTRING = 3.1415

#### Remarks

`DWSTRING` works transparently with literals and Free Basic native strings, e.g.

```
DIM dws AS DWSTRING = "One"
DIM s AS STRING = "Three"
dws = dws & " Two " & s
PRINT dws
```

It can be used with Windows API functions, e.g.

```
DIM nLen AS LONG = SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
DIM dwsText AS DWSTRING = WSPACE(nLen + 1)
SendMessageW(hwnd, WM_GETTEXT, nLen + 1, cast(LPARAM, *cwsText))
dwsText = LEFT(dwsText, LEN(dwsText) - 1)
PRINT dwsText
```

We can use arrays of `DWSTRING` strings transparently, e.g.

```
DIM rg(1 TO 10) AS DWSTRING
FOR i AS LONG = 1 TO 10
   rg(i) = "string " & i
NEXT

FOR i AS LONG = 1 TO 10
   PRINT rg(i)
NEXT
```

A two-dimensional array

```
DIM rg2 (1 TO 2, 1 TO 2) AS DWSTRING
rg2(1, 1) = "string 1 1"
rg2(1, 2) = "string 1 2"
rg2(2, 1) = "string 2 1"
rg2(2, 2) = "string 2 2"
PRINT rg2(2, 1)
```

REDIM PRESERVE / ERASE

```
REDIM rg(0) AS DWSTRING
rg(0) = "string 0"
REDIM PRESERVE rg(0 TO 2) AS DWSTRING
rg(1) = "string 1"
rg(2) = "string 2"
PRINT rg(0)
PRINT rg(1)
PRINT rg(2)
ERASE rg
```

You can also use it with files:

Using FreeBasic intrinsic statements:

```
DIM dws AS DWSTRING = "Дмитрий Дмитриевич Шостакович"
DIM f AS LONG = FREEFILE
OPEN "тест.txt" FOR OUTPUT ENCODING "utf16" AS #f
PRINT #f, cws
CLOSE #f
```

Using the Windows API:

```
' // Writing to a file
DIM dwsFilename AS DWSTRING = "тест.txt"
DIM dwsText AS DWSTRING = "Дмитрий Дмитриевич Шостакович"
DIM hFile AS HANDLE = CreateFileW(cwsFilename, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)
IF hFile THEN
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS LONG = WriteFile(hFile, cwsText, LEN(dwsText) * 2, @dwBytesWritten, NULL)
   CloseHandle(hFile)
END IF
```

```
' // Read the file
hFile = CreateFileW(dwsFilename, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
IF hFile THEN
   DIM dwFileSize AS DWORD = GetFileSize(hFile, NULL)
   IF dwFileSize THEN
      DIM dwsOut AS DWSTRING = WSPACE(dwFileSize \ 2)
      DIM bSuccess AS LONG = ReadFile(hFile, *dwsOut, dwFileSize, NULL, NULL)
      CloseHandle(hFile)
      PRINT dwsOut
   END IF
END IF
```

Notice that, contrarily to **CreateFileW**, FreeBasic's OPEN statement doesn't allow to use unicode for the file name.

# Operators

## <a name="Operator*"></a>Operator *

Dereferences the CWSTR.<br>One * returns the address of the CWSTR buffer.<br>Two ** returns the address of the start of the string data.

```
OPERATOR * (BYREF cws AS CWSTR) AS WSTRING PTR
```
## <a name="Operator&"></a>Operator &

Concatenates strings.

```
OPERATOR & (BYREF cws1 AS CWSTR, BYREF cws2 AS CWSTR) AS CWSTR
```

## <a name="Operator+="></a>Operator +=

Appends a string to the CWSTR.

```
OPERATOR += (BYREF wszStr AS CONST WSTRING)
OPERATOR += (BYREF cws AS CWStr)
OPERATOR += (BYREF cbs AS CBStr)
OPERATOR += (BYREF ansiStr AS STRING)
```

## <a name="Operator&="></a>Operator &=

Appends a string to the CWSTR.

```
OPERATOR &= (BYREF wszStr AS CONST WSTRING)
OPERATOR &= (BYREF cws AS CWStr)
OPERATOR &= (BYREF cbs AS CBStr)
OPERATOR &= (BYREF ansiStr AS STRING)
```

## <a name="Operator[]"></a>Operator []

Gets or sets the corresponding unicode integer representation of the character at the specified position. The index parameter is zero based (0 for the first character, 1 for the second, etc.). This operator must not be used in case of empty string because reference is undefined (inducing runtime error). Otherwise, the user must ensure that the index does not exceed the range "\[0, Len(cws) - 1]". Outside this range, results are undefined.

```
OPERATOR [] (BYVAL nIndex AS LONG) BYREF AS USHORT
```
#### Example
```
DIM cws as CWSTR = "1234567890"
print cws[0]
cws[0] = ASC("X")
print cws
```

## <a name="OperatorLet"></a>Operator Let

Assigns a string to the CWSTR.

```
OPERATOR LET (BYREF wszStr AS CONST WSTRING)
OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
OPERATOR LET (BYREF cws AS CWStr)
OPERATOR LET (BYREF cbs AS CBStr)
OPERATOR LET (BYREF ansiStr AS STRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | A WSTRING. |
| *pwszStr* | A pointer to a WSTRING. |
| *cws* | A CWSTR. |
| *cbs* | A CBSTR. |
| *ansiStr* | An ansi string or string literal. |

# Casting and Conversions

## <a name="OperatorCast"></a>Operator Cast

```
OPERATOR CAST () BYREF AS CONST WSTRING
OPERATOR CAST () AS ANY PTR
```

Returns a pointer to the CWSTR buffer or the string data. These operators aren't called directly.

## <a name="bstr"></a>bstr

Returns the contents of the CWSTR as a BSTR.

```
FUNCTION bstr () AS AFX_BSTR
```

## <a name="cbstr"></a>cbstr

Returns the contents of the CWSTR as a CBSTR.

```
FUNCTION cbstr () AS CBStr
```

## <a name="wchar"></a>wchar

Returns the string data as a new unicode string allocated with CoTaskMemAlloc.

```
FUNCTION wchar () AS WSTRING PTR
```

Useful when we need to pass a pointer to a null terminated wide string to a function or method that will release it. If we pass a WSTRING it will GPF. If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and returns a valid pointer to that item. If there is insufficient memory available, CoTaskMemAlloc returns NULL.

## <a name="Utf8"></a>Utf8

Converts from UTF8 to Unicode and from Unicode to UTF8.

```
PROPERTY Utf8() AS STRING
PROPERTY Utf8 (BYREF utf8String AS STRING)
```

# Methods

## <a name="Capacity"></a>Capacity

Gets/sets the size of the internal buffer. The size is the number of bytes which can be stored without further expansion.

```
PROPERTY Capacity() AS UINT
PROPERTY Capacity (BYVAL nValue AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nValue* | The new capacity value, in bytes. If the new capacity is equal to the current capacity, no operation is performed; is it is smaller, the buffer is shortened and the contents that exceed the new capacity are lost. If you pass an odd number, 1 is added to the value to make it even. |
