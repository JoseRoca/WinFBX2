# Array Macros

Macros to manipulate one-dimensional arrays of variable length.

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

### <a name="appendelementtoarray"></a>AppendElementToArray

Appends a new element to a varible length one-dimensional array.
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

Appends a one-dimensional array to another one-dimensional array.
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
