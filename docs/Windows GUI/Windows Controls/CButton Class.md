# CButton Class

Wrapper class on top of all the Windows `Button` messages. A *button* is a control the user can click to provide input to an application.

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
