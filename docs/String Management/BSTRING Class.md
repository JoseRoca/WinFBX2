# BSTRING Class

The `BSTRING` and `DWSTRING` classes implement a dynamic unicode null terminated string. Free Basic has a dynamic string data type (STRING) and a fixed length unicode data type (WSTRING), but it lacks a dynamic unicode string. `BSTRING` and `DWSTRING` behave as if they were native data types, working directly with the intrinsic FreeBasic string functions and operators.

While `DWSTRING`does its own memory allocations, `BSTRING`is a wrapper on top of of the OLE string (aka BSTR) data type and the memory management is done by the COM library. It is better to use `DWSTRING`for general purposes, since it is faster, reserving the use of `BSTRING`for COM programming.

**Include file**: BSTRING.INC.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructors) | Initialize the class with the specified value. |
| [Operator \*](#operator*) | One * returns the address of the `BSTRING` buffer.<br> Two ** returns the address of the start of the string data. |
| [Operator Cast](#operatorcast) | Returns a pointer to the `BSTRING` buffer or the string data.<br>Casting is automatic. You don't have to call this operator. |
| [Attach](#attach) | Attaches a BSTR to the BSTRING class. |
| [Detach](#detach) | Detaches the underlying BSTR from the BSTRING class and returns it as the result of the function. |
| [Utf8](#utf8) | Converts from UTF8 to Unicode and from Unicode to UTF8. |
| [wchar](#wchar) | Returns the string data as a new unicode string allocated with CoTaskMemAlloc. |

# <a name="constructors"></a>Constructors

Initialize the class with the specified value.

```
CONSTRUCTOR BSTRING
CONSTRUCTOR BSTRING (BYREF bs AS BSTRING)
CONSTRUCTOR BSTRING (BYREF dws AS DWSTRING)
CONSTRUCTOR BSTRING (BYVAL pwszStr AS WSTRING PTR)
CONSTRUCTOR BSTRING (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
CONSTRUCTOR (BYVAL nChars AS LONG, BYREF wszFill AS CONST WSTRING)
CONSTRUCTOR (BYVAL n AS LONGINT)
CONSTRUCTOR (BYVAL n AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | A WSTRING. |
| *dws* | A DWSTRING. |
| *bs* | A BSTRING. |
| *ansiStr* | An ansi string or string literal. |
| *nCodePage* | The code page to be used for ansi to unicode conversions. |
| *n* | A number. |

For a list of code pages see: [Code Page Identifiers](https://msdn.microsoft.com/en-us/library/windows/desktop/dd317756(v=vs.85).aspx)

# Operators

## <a name="operator*"></a>Operator *

Deferences the `BSTRING`.<br>One * returns the address of the underlying `BSTR` handle (same as STRPTR).<br>Two ** returns the address of the start of the string data (same as *STRPTR).

```
OPERATOR * (BYREF bs AS BSTRING) AS WSTRING PTR
```

# Casting and Conversions

## <a name="operatorcast"></a>Operator Cast

Returns a pointer to the underlying `BSTR` or the string data. These operators aren't called directly.

```
OPERATOR CAST () BYREF AS CONST WSTRING
OPERATOR CAST () AS ANY PTR
```

## <a name="wchar"></a>wchar

Returns the string data as a new unicode string allocated with **CoTaskMemAlloc**.

Useful when we need to pass a pointer to a null terminated wide string to a function or method that will release it. If we pass a WSTRING it will GPF. If the length of the input string is 0, **CoTaskMemAlloc** allocates a zero-length item and returns a valid pointer to that item. If there is insufficient memory available, **CoTaskMemAlloc** returns NULL.

```
FUNCTION wchar () AS WSTRING PTR
```

## <a name="utf8"></a>Utf8

Converts from UTF8 to Unicode and from Unicode to UTF8.

```
PROPERTY Utf8() AS STRING
PROPERTY Utf8 (BYREF utf8String AS STRING)
```

# Methods

## <a name="attach"></a>Attach

Attaches a `BSTR` to the `BSTRING` class.

```
SUB Attach (BYVAL pbstrSrc AS BSTR)
```

## <a name="detach"></a>Detach

Detaches the underlying `BSTR` from the `BSTRING` class and returns it as the result of the function. The returned pointer must be freed by **SysFreeString**.

```
FUNCTION Detach () AS BSTR
```

This method frees the *m_bstr* member of the `BSTRING` class and returns it as the result of the function. Because it no longer belongs to the class, it must be freed by **SysFreeString**.
