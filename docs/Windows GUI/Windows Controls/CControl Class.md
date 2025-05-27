# CControl Class

`CControl` is a base class for all the other classes that deal with Windows controls. It has methods to set and get the last result code. It has been implemented because the Windows API functions **GetLastError** and **SetLastError** can't be used reliably with wrapper classes or procedures that return another class (the construction of the returned class may also use these functions and wipe the error code).

**Include file**: CControl.inc.

| Name       | Description |
| ---------- | ----------- |
| [CDialogOwnerPtr](#cdialogownerptr) | Returns a pointer to the CDialog class given the handle of the window created with it or the handle of any of it's children |
| [ComCtlVersion](#comctlversion) | Returns the version of CommCtl32.dll multiplied by 100, e.g. 600 for version 6.0. |
| [GetFileVersion](#getfileversion) | Returns the version of specified file multiplied by 100, e.g. 601 for version 6.01. |
| [GetUser](#getuser) | Retrieves a value from the user data area of the control. |
| [SetUser](#setuser) | Sets a value in the user data area of the control. |
| [UsesPixels](#usespixels) | Returns true if the dialog uses pixels |

**Error functions**

| Name       | Description |
| ---------- | ----------- |
| [GetErrorInfo](#geterrorinfo) | Returns a localized description of the specified error code. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [SetResult](#setresult) | Sets the last result code. |

---

### <a name="geterrorinfo"></a>GetErrorInfo

Returns a localized description of the specified error code. If the error is omited, it will return the value returned by the Windows API function **GetLastError**.
```
PRIVATE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
```
---

### <a name="getlastresult"></a>GetLastResult

Returns the last result code
```
FUNCTION GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
```
---

### <a name="setresult"></a>SetResult

Sets the last result code.
```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

### <a name="cdialogownerptr"></a>CDialogOwnerPtr

Returns a pointer to the CDialog class given the handle of the window created with it or the handle of any of it's children.
```
FUNCTION CDialogOwnerPtr (BYVAL hwnd AS HWND) AS CDialog PTR
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | The handle of the window |

```
---

### <a name="getuser"></a>GetUser

Retrieves a value from the user data area of the control.
```
FUNCTION GetUser (BYVAL idx AS LONG) AS LONG_PTR
```
| Parameter | Description |
| --------- | ----------- |
| *idx* | The index number of the user data value to retrieve, in the range 0 to 19 inclusive. |

```
---

### <a name="setuser"></a>SetUser

Sets a value in the user data area of the control.
```
FUNCTION GetUser (BYVAL idx AS LONG) AS LONG_PTR
```
| Parameter | Description |
| --------- | ----------- |
| *idx* | The index number of the user data value to set, in the range 0 to 19 inclusive. |

```
---

### <a name="usespixels"></a>UsesPixels

Gets/sets a flag indicating if the dialog uses pixels (TRUE). If this flag is false, then the dialog uses dialog units.
```
PROPERTY UsesPixels () AS BOOLEAN
PROPERTY UsesPixels (BYVAL bUsePixels AS BOOLEAN)
```
---
