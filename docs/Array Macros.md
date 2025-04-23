# Array Macros

Macros to manipulate one-dimensional dynamic arrays.

**Include file**: AfxArrays2.inc.

| Name       | Description |
| ---------- | ----------- |
| [AppendElementToArray](#appendelementtoarray) | Appends a new element to the array. |
| [AppendArrayToArray](#appendarraytoarray) | Appends one array to another array. |
| [InsertElementIntoArray](#insertelementintoarray) | Inserts a new element into an array. |
| [InsertArrayIntoArray](#insertarrayintoarray) | Inserts an array into another array. |
| [RemoveElementFromArray](#removeelementfromarray) | Removes the specified element from an array. |
| [RemoveFirstElementFromArray](#removefirstelementfromarray) | Removes the first element from an array. |
| [RemoveLastElementFromArray](#removelastelementfromarray) | Removes the last element from an array. |
| [RemoveElementsFromArray](#removeelementsfromarray) | Removes the specified elements from an array. |

---

# Sorting Macros

**Include file**: AfxSort2.inc.

| Name       | Description |
| ---------- | ----------- |
| [AfxSortNumericArray](#afxsortnumericarray) | Sorts one-dimensional numeric arrays of any type. |
| [AfxSortStringArray](#afxsortstringarray) | Sorts one-dimensional string arrays of any type. |

---

### <a name="appendelementtoarray"></a>AppendElementToArray

Appends a new element to a dynamic one-dimensional array.
```
#macro AppendElementToArray(rg, elem, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *elem* | The element to append. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array and the element to append can be of any type, but they have to be of the same type between them.

#### Usage example

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
' Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="appendarraytoarray"></a>AppendArrayToArray

Appends a one-dimensional array to another dynamic one-dimensional array.
```
#macro AppendArrayToArray(rg, rg2, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *rg2* | The array to append. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array and the array to append can be of any type, but they have to be of the same type between them.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM rg2(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
' Fill the array to append
xStr = "String 2 - "
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, xStr & WSTR(i), res)
NEXT
' // Append the array to the first array
AppendArrayToArray(rg, rg2, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM rg2(ANY) AS LONG
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong, res)
   nLong += 1
NEXT
' Fill the array to insert
nLong = 12345
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, nLong, res)
   nLong += 1
NEXT
AppendArrayToArray(rg, rg2, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="insertelementintoarray"></a>InsertElementIntoArray

Inserts a new element before the specified position into a dynamic one-dimensional array.
```
#macro InsertElementIntoArray(rg, pos, elem, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *pos* | The position in the array where the new element will be added. This position is relative to the lower bound of the array. |
| *elem* | The element to insert. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array and the array to append can be of any type, but they have to be of the same type between them.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
InsertElementIntoArray(rg, 5, xStr & WSTR(11), res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM rg2(ANY) AS LONG
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong, res)
   nLong += 1
NEXT
' Fill the array to insert
nLong = 12345
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, nLong, res)
   nLong += 1
NEXT
InsertElementIntoArray(rg, 5, nLong, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="insertarrayintoarray"></a>InsertArrayIntoArray

Inserts a one-dimensional array before the specified position in another dynamic one-dimensional array.
```
#macro InsertArrayIntoArray(rg, pos, rg2, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *pos* | The position in the array where the new element will be added. This position is relative to the lower bound of the array. |
| *rg2* | The array to insert. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array and the array to insert can be of any type, but they have to be of the same type between them.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM rg2(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
' Fill the array to insert
xStr = "String 2 - "
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, xStr & WSTR(i), res)
NEXT
' // Insert the array to the first array
InsertArrayIntoArray(rg, 5, rg2, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM rg2(ANY) AS LONG
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong, res)
   nLong += 1
NEXT
' Fill the array to insert
nLong = 12345
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, nLong, res)
   nLong += 1
NEXT
' // Insert the array to the first array
InsertArrayIntoArray(rg, 5, rg2, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="removeelementfromarray"></a>RemoveElementFromArray

Removes the specified element of a dynamic one-dimensional array.
```
#macro RemoveElementFromArray(rg, pos, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *pos* | The position in the array of the element to remove. This position is relative to the lower bound of the array. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array can be of any type.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
' // Remove the fifth element
RemoveElementFromArray(rg, 5, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   REDIM PRESERVE rg(UBOUND(rg) + 1)
   rg(i - 1) = nLong
   nLong += 1
NEXT
' // Remove the specified element
RemoveElementFromArray(rg, 5, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="removefirstelementfromarray"></a>RemoveFirstElementFromArray

Removes the first element of a dynamic one-dimensional array.
```
#macro RemoveFirstElementFromArray(rg, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array can be of any type.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
' // Remove the frst element
RemoveFirstElementFromArray(rg, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 12345
FOR i AS LONG = 1 TO 10
   REDIM PRESERVE rg(UBOUND(rg) + 1)
   rg(i - 1) = nLong
   nLong += 1
NEXT
' // Remove the first element
RemoveFirstElementFromArray(rg, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="removelastelementfromarray"></a>RemoveLastElementFromArray

Removes the last element of a dynamic one-dimensional array.
```
#macro RemoveLastElementFromArray(rg, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array can be of any type.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i), res)
NEXT
' // Remove the last element
RemoveLastElementFromArray(rg, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG   ' // or DWORD, SINGLE, DOUBLE, etc.
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 12345
FOR i AS LONG = 1 TO 10
   REDIM PRESERVE rg(UBOUND(rg) + 1)
   rg(i - 1) = nLong
   nLong += 1
NEXT
' // Remove the last element
RemoveLastElementFromArray(rg, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```

### <a name="Removeelementsfromarray"></a>RemoveElementsFromArray

Removes the specified elements of a dynamic one-dimensional array.
```
#macro RemoveElementsFromArray(rg, rgElements, res)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *rgElements* | Array of long values indicating the 0-based positions of the elements to remove. |
| *res* | The result code. A boolean true of false value. |

#### Remarks

The array can be of any type.

#### Usage example

```
DIM rg(ANY) AS LONG
DIM res AS BOOLEAN
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong, res)
   nLong += 1
NEXT
' Fill the array of elements to delete
DIM rgElements(0 TO 3) AS LONG => {2, 4, 6, 8}
RemoveElementsFromArray(rg, rgElements, res)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="afxsortnumericarray"></a>AfxSortNumericArray

Sorts one-dimensional numeric array of any type.
```
#macro AfxSortNumericArray(rg, ascending)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array to sort. |
| *ascending* | If true, sorts the array in ascending order. I false, sorts the array in descending order. |

#### Return code

It does not return a result.

#### Usage example
```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rgstr(ANY) AS XSTRING
DIM rgstr2(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
DIM res AS BOOLEAN
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rgstr, xStr & WSTR(i), res)
NEXT
' Fill the array to append
xStr = "String 2 - "
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rgstr2, xStr & WSTR(i), res)
NEXT
' // Append the array to the first array
AppendArrayToArray(rgstr, rgstr2, res)
' // Display the array
FOR i AS LONG = LBOUND(rgstr) TO UBOUND(rgstr)
   print rgstr(i)
NEXT
' // Sort the array
AfxSortStringArray(rgstr, true)
' // Display the sorted array
FOR i AS LONG = LBOUND(rgstr) TO UBOUND(rgstr)
   print rgstr(i)
NEXT
```
---

### <a name="afxsortstringarray"></a>AfxSortStringArray

Sorts one-dimensional string array of any type.
```
#macro AfxSortStringArray(rg, ascending)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array to sort. |
| *ascending* | If true, sorts the array in ascending order. I false, sorts the array in descending order. |

#### Return code

It does not return a result.

#### Usage example
```
#define XNUMBER DOUBLE ' // or SHORT, INTEGER, LONG, LONGINT, etc.
DIM rgNum(ANY) AS XNUMBER
DIM rgNum2(ANY) AS XNUMBER
DIM bNumRes AS BOOLEAN
DIM xNum AS XNUMBER = 1234567.89
' // Fill the array
FOR i AS LONG = 1 TO 10
   xNum += 0.01
   AppendElementToArray(rgNum, xNum, bNumRes)
NEXT
' Fill the array to insert
xNum = 111.01
FOR i AS LONG = 1 TO 5
   xNum += 1
   AppendElementToArray(rgNum, xNum, bNumRes)
NEXT
' // Insert the array to the first array
InsertArrayIntoArray(rgNum, 5, rgNum2, bNumRes)
' // Display the array
FOR i AS LONG = LBOUND(rgNum) TO UBOUND(rgNum)
   print rgNum(i)
NEXT
' // Sort the numeric array
AfxSortNumericArray(rgNum, true)
' // Display the sorted array
FOR i AS LONG = LBOUND(rgNum) TO UBOUND(rgNum)
   print rgNum(i)
NEXT
```
---
