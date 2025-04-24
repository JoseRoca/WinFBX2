# CButton Class

Wrapper class on top of all the Windows `Button` messages and functions.

A button control is a small, rectangular child window that can be clicked on and off. Buttons can be used alone or in groups and can either be labeled or appear without text. A button typically changes appearance when the user clicks it.

Typical buttons are the check box, radio button, and pushbutton. A `CButton` object can become any of these, according to the button style specified at its initialization by the **AddControl** method of the `CDialog`class.

**Include file**: CButton.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructors) | Create instances of the `CButton` class. |
| [Check](#check) | Sets the button state to checked. |
| [CheckDlgButton](#checkdlgbutton) | Changes the check state of a button control. |
| [CheckRadioButton](#checkradiobutton) | Adds a check mark to (checks) a specified radio button in a group and removes a check mark from (clears) all other radio buttons in the group. |
| [Click](#click) | Simulates the user clicking a button. |
| [DeleteBitmap](#deletebitmap) | Deletes a bitmap associated with a button. |
| [DeleteIcon](#deleteicon) | Deletes an icon associated with a button. |
| [DeleteImage](#deleteimage) | Deletes an image (icon or bitmap) associated a button. |
| [Disable](#disable) | Disables a button. |
| [Enable](#enable) | Enables a button. |
| [GetCheck](#getcheck) | Gets the check state of a radio button or check box.  |
| [GetIdealSize](#getidealsize) | Gets the size of the button that best fits its text and image, if an image list is present. |
| [GetImage](#getimage) | Retrieves a handle to the image (icon or bitmap) associated with the button. |
| [GetImageList](#getimagelist) | Gets the BUTTON_IMAGELIST structure that describes the image list assigned to a button control.  |
| [GetNote](#getnote) | Gets the text of the note associated with the Command Link button. |
| [GetNoteLength](#getnotelength) | Gets the length of the note text that may be displayed in the description for a command link. |
| [GetSplitInfo](#getsplitinfo) | Gets information for a specified split button control. |
| [GetState](#getstate) | Retrieves the state of a button or check box. |
| [GetStyle](#getstyle) | Retrieves the style of button. |
| [GetText](#gettext) | Retrieves the text in a button control. |
| [GetTextLength](#gettextlength) | Retrieves the length of the text in a button control. |
| [GetTextMargin](#gettextmargin) | Retrieves the margins used to draw text in a button control. |
| [Gray](#gray) | Sets the button state to grayed, indicating an indeterminate state. |
| [IsDlgButtonChecked](#isdlgbuttonchecked) | Dtermines whether a button control is checked or whether a three-state button control is checked, unchecked, or indeterminate. |
| [SetBitmap](#setbitmap) | Associates a new bitmap with a button. |
| [SetCheck](#setcheck) | Sets the check state of a radio button or check box. |
| [SetDontClick](#setdontclick) | Sets a flag on a radio button that controls the generation of BN_CLICKED messages when the button receives focus. |
| [SetDropdownState](#setdropdownstate) | Sets the drop down state for a specified button with style of BS_SPLITBUTTON. |
| [SetElevationRequiredState](#setelevationrequiredstate) | Sets the elevation required state for a specified button or command link to display an elevated icon. |
| [SetIcon](#seticon) | Associates a new icon with the button. |
| [SetImage](#setimage) | Associates a new image (icon or bitmap) with the button. |
| [SetImageList](#setimagelist) | Assigns an image list to a button control. |
| [SetNote](#setnote) | Sets the text of the note associated with a specified Command Link button. |
| [SetSplitInfo](#setsplitinfo) | Sets information for a specified split button control. |
| [SetState](#setstate) | Sets the highlight state of a button. |
| [SetStyle](#setstyle) | Sets the style of a button. |
| [SetText](#settext) | Sets the text of a button. |
| [SetTextMargin](#settextmargin) | Sets the margins for drawing text in a button control. |
| [Uncheck](#uncheck) | Sets the button state to cleared. |

---

# <a name="constructors"></a>Constructors

Creates instances of the `CButton` class.

```
CONSTRUCTOR (BYVAL hCtl AS HWND)
CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL hCtl AS HWND)
CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *hCtl* | Handle of the button control. |
| *hParent* | Handle of the parent window of the button control. |
| *cID* | Control identifier of the button control. |
| *pDlg* | Pointer to an instance of the `CDialog` class. |

---

## <a name="check"></a>Check

Sets the button state to checked.

```
SUB Check ()
```
---

## <a name="checkdlgbutton"></a>CheckDlgButton

Changes the check state of a button control.

```
FUNCTION CheckDlgButton (BYVAL cIDButton AS LONG, BYVAL uCheck AS UINT) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *cIDButton* | The identifier of the button to modify. |
| *uCheck* | The check state of the button. This parameter can be one of the following values. |

| Value  | Meaning |
| ------ | ------- |
| **BST_CHECKED** | Sets the button state to checked. |
| **BST_INDETERMINATE** | Sets the button state to grayed, indicating an indeterminate state. Use this value only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style. |
| **BST_UNCHECKED** | Sets the button state to cleared |

#### Return value

If the function succeeds, the return value is true. If the function fails, the return value is false.

#### Remarks

The **CheckDlgButton** function sends a **BM_SETCHECK** message to the specified button control in the specified dialog box.

---

## <a name="checkradiobutton"></a>CheckRadioButton

Adds a check mark to (checks) a specified radio button in a group and removes a check mark from (clears) all other radio buttons in the group.

```
FUNCTION CheckRadioButton (BYVAL cIDFirstButton AS LONG, BYVAL cIDLastButton AS LONG, _
   BYVAL cIDCheckButton AS LONG) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *cIDFirstButton* | The identifier of the first radio button in the group. |
| *cIDLastButton* | The identifier of the last radio button in the group. |
| *cIDCheckButton* | The identifier of the radio button to select. |

#### Return value

If the function succeeds, the return value is true.
If the function fails, the return value is false. 

#### Remarks
The **CheckRadioButton** function sends a **BM_SETCHECK** message to each of the radio buttons in the indicated group.

The *cIDFirstButton* and *cIDLastButton* parameters specify a range of button identifiers (normally the resource IDs of the buttons). The position of buttons in the tab order is irrelevant; if a button forms part of a group, but has an ID outside the specified range, it is not affected by this call.

---

## <a name="click"></a>Click

Simulates the user clicking a button. This message causes the button to receive the **WM_LBUTTONDOWN** and **WM_LBUTTONUP** messages, and the button's parent window to receive a **BN_CLICKED** notification code.

```
SUb Click ()
```

#### Return value

This message does not return a value.

#### Remarks

If the button is in a dialog box and the dialog box is not active, the **BM_CLICK** message might fail. To ensure success in this situation, call the **SetActiveWindow** function to activate the dialog box before sending the BM_CLICK message to the button.

---

## <a name="deletebitmap"></a>DeleteBitmap

Deletes a bitmap associated with a button. Returns TRUE or FALSE.

```
FUNCTION DeleteBitmap () AS BOOLEAN
```
---

## <a name="deleteicon"></a>DeleteIcon

Deletes a icon associated with a button. Returns TRUE or FALSE.

```
FUNCTION DeleteIcon () AS BOOLEAN
```
---

## <a name="deleteimage"></a>DeleteImage

Deletes an image (icon or bitmap) associated a button. Returns TRUE or FALSE.

```
FUNCTION DeleteImage () AS BOOLEAN
```
---

## <a name="disable"></a>Disable

Disables a button. Returns FALSE if the windows was previous disabled; otherwise TRUE.

```
FUNCTION Disable () AS BOOLEAN
```
---

## <a name="enable"></a>Enable

Enables a button. Returns FALSE if the windows was previous enabled; otherwise TRUE.

```
FUNCTION Enable () AS BOOLEAN
```
---

## <a name="getcheck"></a>GetCheck

Gets the check state of a radio button or check box. 

```
FUNCTION GetCheck () AS DWORD
```
#### Return value

The return value from a button created with the **BS_AUTOCHECKBOX**, **BS_AUTORADIOBUTTON**, **BS_AUTO3STATE**, **BS_CHECKBOX**, **BS_RADIOBUTTON**, or **BS_3STATE** style can be one of the following.

| Return code  | Description |
| ------------ | ----------- |
| **BST_CHECKED** | Button is checked. |
| **BST_INDETERMINATE** | Button is grayed, indicating an indeterminate state (applies only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style). |
| **BST_UNCHECKED** | Button is cleared. |

---

