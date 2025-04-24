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
| [GetBitmap](#getbitmap) | Retrieves a handle to the bitmap associated with the button. |
| [GetIcon](#geticon) | Retrieves a handle to the icon associated with the button. |
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

## <a name="getidealsize"></a>GetIdealSize

Gets the size of the button that best fits its text and image, if an image list is present.

```
FUNCTION GetIdealSize (BYREF tSize AS SIZE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *tsize* | A pointer to a **SIZE** structure that receives the desired size of the button, including text and image list, if present. The calling application is responsible for allocating this structure. Set the **cx** and **cy** members to zero to have the ideal height and width returned in the **SIZE** structure. To specify a button width, set the **cx** member to the desired button width. The system will calculate the ideal height for this width and return it in the **cy** member. |

#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

#### Remarks

If no special button width is desired, then you must set both members of **SIZE** to zero to calculate and return the ideal height and width. If the value of the **cx** member is greater than zero, then this value is considered the desired button width, and the ideal height for this width is calculated and returned in the **cy** member.

This message is most applicable to PushButtons. When sent to a PushButton, the message retrieves the bounding rectangle required to display the button's text. Additionally, if the PushButton has an image list, the bounding rectangle is also sized to include the button's image.

When sent to a button of any other type, the size of the control's window rectangle is retrieved.

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0.

---

## <a name="getbitmap"></a>GetBitmap

Retrieves a handle to the bitmap associated with the button.

#### Return value

The return value is a handle to the bitmap, if any; otherwise, it is NULL.

---

## <a name="geticon"></a>GetIcon

Retrieves a handle to the icon associated with the button.

#### Return value

The return value is a handle to the icon, if any; otherwise, it is NULL.

---

## <a name="getimage"></a>GetImage

Retrieves a handle to the image (icon or bitmap) associated with the button.

```
FUNCTION GetImage (BYVAL imageType aS LONG) AS HANDLE
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The type of image to associate with the button. This parameter can be one of the following values. |

| Value | Meaning |
| ------------ | ----------- |
| **IMAGE_BITMAP** | A bitmap. |
| **IMAGE_ICON** | An icon. |

#### Return value

The return value is a handle to the image, if any; otherwise, it is NULL.

---

## <a name="getimagelist"></a>GetImageList

Gets the **BUTTON_IMAGELIST** structure that describes the image list assigned to a button control. 

```
FUNCTION GetImageList () AS BUTTON_IMAGELIST PTR
```

#### Return value

A pointer to a **BUTTON_IMAGELIST** structure that contains image list information.

---

## <a name="getnote"></a>GetNote

Gets the text of the note associated with the Command Link button.

```
FUNCTION GetNote () AS DWSTRING
```

#### Remarks

The **BCM_GETNOTE** message works only with buttons that have the **BS_COMMANDLINK** or **BS_DEFCOMMANDLINK** button style.

**GetLastResul** will contain:

**ERROR_NOT_SUPPORTED**, if the button does not have the **BS_DEFCOMMANDLINK** or **BS_COMMANDLINK** style.

---

## <a name="getnotelength"></a>GetNoteLength

Gets the length of the note text that may be displayed in the description for a command link.

```
FUNCTION GetNoteLength () AS LONG
```

#### Return value

Returns the length of the note text in Unicode characters, not including any terminating NULL, or zero if there is no note text.

---

## <a name="getsplitindo"></a>GetSplitInfo

Gets information for a split button control. 

```
FUNCTION GetSplitInfo (BYREF pInfo AS BUTTON_SPLITINFO) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *pInfo* | A pointer to a **BUTTON_SPLITINFO** structure to receive information on the button. The caller is responsible for allocating the memory for the structure. Set the **mask** member of this structure to determine what information to receive. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Use this message only with the **BS_SPLITBUTTON** and **BS_DEFSPLITBUTTON** button styles.

---

## <a name="getstate"></a>GetState

Retrieves the state of a button or check box. 

```
FUNCTION GetState () AS DWORD
```

#### Return value

The return value specifies the current state of the button. It is a combination of the following values.

| Return code  | Description |
| ------------ | ----------- |
| **BST_CHECKED** | The button is checked. |
| **BST_DROPDOWNPUSHED** | The button is in the drop-down state. Applies only if the button has the **TBSTYLE_DROPDOWN** style. |
| **BST_FOCUS** | The button has the keyboard focus. |
| **BST_HOT** | The button is hot; that is, the mouse is hovering over it. |
| **BST_INDETERMINATE** | The state of the button is indeterminate. Applies only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style. |
| **BST_PUSHED** | The button is being shown in the pushed state. |
| **BST_UNCHECKED** | No special state. Equivalent to zero. |

---

## <a name="getstyle"></a>GetStyle

Retrieves the style of button.

```
FUNCTION GetStyle () AS DWORD
```
---

## <a name="gettext"></a>GetText

Retrieves the text in a button control.

```
FUNCTION GetText () AS DWSTRING
```
---

## <a name="gettextlength"></a>GetTextLength

Determines the length, in characters, of the text associated with a button.

```
FUNCTION GetTextLength () AS LONG
```

#### Return value

The return value is the length of the text in characters, not including the terminating null character.

---

## <a name="gettextmargin"></a>GetTextMargin

Retrieves the margins used to draw text in a button control.

```
FUNCTION GetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *tMargin* | A pointer to a **RECT** structure that contains the margins to use for drawing text. |

#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

#### Remarks

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0. For more information on manifests, see Enabling Visual Styles.

---

## <a name="gray"></a>Gray

Sets the button state to grayed, indicating an indeterminate state. Use this value only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style.

```
SUB Gray ()
```

## <a name="isdlgbuttonchecked"></a>IsDlgButtonChecked

Determines whether a button control is checked or whether a three-state button control is checked, unchecked, or indeterminate.

```
FUNCTION IsDlgButtonChecked (BYVAL cIDButton AS LONG) AS UINT
```
| Parameter | Description |
| --------- | ----------- |
| *cIDButton* | The identifier of the button control. |

#### Return value

The return value from a button created with the **BS_AUTOCHECKBOX**, **BS_AUTORADIOBUTTON**, **BS_AUTO3STATE**, **BS_CHECKBOX**, **BS_RADIOBUTTON**, or **BS_3STATE** styles can be one of the values in the following table. If the button has any other style, the return value is zero.

| Returm code | Description |
| ----------- | ----------- |
| **BST_CHECKED** | The button is checked. |
| **BST_INDETERMINATE** | The button is in an indeterminate state (applies only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style). |
| **BST_UNCHECKED** | The button is not checked. |

#### Remarks

The **IsDlgButtonChecked** method sends a **BM_GETCHECK** message to the specified button control.

