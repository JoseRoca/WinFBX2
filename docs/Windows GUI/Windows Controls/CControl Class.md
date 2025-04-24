#CControl Class

`CControl` is a base class for all the other classes that deal with Windows controls. It has methods to set and get the last result code. It has been implemented because the Windows API functions **GetLastError** and **SetLastError** can't be used reliably with wrapper classes or procedures that return another class (the construction of the returned class may also use these functions and wipe the error code).

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


