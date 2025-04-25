# CEdit Class

`CEdit` is a class on top of the Edit control. An *edit control* is a rectangular control window typically used in a dialog box to permit the user to enter and edit text by typing on the keyboard. The system automatically processes all user-initiated text operations and notifies the application when the operations are completed.

**Include file**: CEdit.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructors) | Create instances of the `CEdit` class. |
| [CanUndo](#canundo) | Determines whether there are any actions in an edit control's undo queue. |
| [CharFromPos](#charfrompos) | Gets information about the character closest to a specified point in the client area of an edit control. |
| [Clear](#clear) | Deletes (clears) the current selection, if any, from the edit control. |
| [Copy](#copy) | Copies the current selection to the clipboard in CF_TEXT format. |
| [Cut](#cut) | Deletes (cuts) the current selection, if any, in the edit control and copy the deletedtext to the clipboard in CF_TEXT format. |
| [Disable](#disable) | Disables a button |
| [EmptyUndoBuffer](#emptyundobuffer) | Resets the undo flag of an edit control. |
| [Enable](#enable) | Enables a button. |
| [FmtLines](#fmtlines) | Sets a flag that determines whether a multiline edit control includes soft line-break characters. |
| [GetCueBannerText](#getcuebannertext) | Gets the text that is displayed as the textual cue, or tip, in an edit control. |
| [GetFirstVisibleLine](#getfirstvisibleline) | Gets the zero-based index of the uppermost visible line in a multiline edit control. |
| [GetHandle](#gethandle) | Gets a handle of the memory currently allocated for a multiline edit control's text. |
| [GetIMEStatus](#getimestatus) | Retrieves a set of status flags that indicate how the edit control interacts with the Input Method Editor (IME). |
| [GetLimitText](#getlimittext) | Gets the current text limit for an edit control. The return value is the text limit. |
| [GetLine](#getline) | Copies a line of text from an edit control. |
| [GetLineCount](#getlinecount) | Gets the number of lines in a multiline edit control. |
| [GetMargins](#getmargins) | Returns the width of the left margin in the LOWORD, and the width of the right margin in the HIWORD. |
| [GetLeftMargin](#getleftmargin) | Returns the width of the left margin for an edit control. |
| [GetRightMargin](#getrightmargin) | Returns the width of the right margin for an edit control. |
| [GetModify](#getmodify) | Gets the state of an edit control's modification flag. |
| [GetPasswordChar](#getpasswordchar) | Gets the password character that an edit control displays when the user enters text. |
| [GetRect](#getrect) | Gets the formatting rectangle of an edit control. |
| [GetSel](#getsel) | Gets the starting and ending character positions of the current selection in an edit control. |
| [GetSelStart](#getselstart) | Gets the starting character position of the current selection in an edit control. |
| [GetSelEnd](#getselend) | Gets the ending character position of the current selection in an edit control. |
| [GetText](#gettext) | Retrieves the text from an edit control. |
| [GetTextLength](#gettextlength) | Retrieves the text length from an edit control. |
| [GetThumb](#getthumb) | Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline edit control. |
| [GetWordBreakProc](#GetWordBreakProc) | Gets the address of the currently registered word-break procedure. |
| [HideBalloonTip](#hideballoontip) | Hides any balloon tip associated with an edit control. |
| [LimitText](#limittext) | Sets the text limit of an edit control. |
| [LineFromChar](#linefromchar) | Gets the index of the line that contains the specified character index in a multiline edit control. |
| [LineLength](#linelength) | Retrieves the length, in characters, of a line in an edit control. |
| [LineScroll](#linescoll) | Scrolls the text in a multiline edit control. |
| [Paste](#paste) | Copies the current content of the clipboard to the edit control at the current caretposition. Data is inserted only if the clipboard contains data in CF_TEXT format. |
| [PosFromChar](#posfromchar) | Retrieves the client area coordinates of a specified character in an edit control. |
| [ReplaceSel](#replacesel) | Replaces the current selection in an edit control with the specified text. |
| [Scroll](#scroll) | Scrolls the text vertically in a multiline edit control. |
| [ScrollCaret](#scrollcaret) | Scrolls the caret into view in an edit control. |
| [SetCueBannerText](#setcuebannertext) | Sets the textual cue, or tip, that is displayed by the edit control to prompt the user for information. |
| [SetCueBannerTextFocused](#setcuebannertextfocused) | Sets the text that is displayed as the textual cue, or tip, for an edit control. |
| [SetHandle](#sethandle) | Sets the handle of the memory that will be used by a multiline edit control. |
| [SetIMEStatus](#setimestatus) | Sets the status flags that determine how an edit control interacts with the Input Method Editor (IME). |
| [SetLimitText](#setlimittext) | Sets the text limit of an edit control. |
| [SetMargins](#setmargins) | Sets the widths of the left and right margins for an edit control. |
| [SetLeftMargin](#setleftmargin) | Sets the width of the left margin for an edit control. |
| [SetRightMargin](#setrightmargin) | Sets the width of the left margin for an edit control. |
| [SetModify](#setmodify) | Sets or clears the modification flag for an edit control. |
| [SetPasswordChar](#setpasswordchar) | Sets or removes the password character for an edit control. |
| [SetReadOnly](#setreadonly) | Sets or removes the read-only style (ES_READONLY) of an edit control. |
| [SetRect](#setrect) | Sets the formatting rectangle of a multiline edit control. |
| [SetRectNoPaint](#setrectnopaint) | Sets the formatting rectangle of a multiline edit control, but itdoes not redraw the edit control window. |
| [SetSel](#setsel) | Selects a range of characters in an edit control. |
| [SetTabStops](#settabstops) | Sets the tab stops in a multiline edit control. |
| [SetText](#settext) | Sets the text of an edit control. |
| [SetWordBreakProc](#setwordbreakproc) | ' Replaces an edit control's default Wordwrap function with an application-defined Wordwrap function. |
| [ShowBalloonTip](#showballoontip) | Displays a balloon tip associated with an edit control. |
| [Undo](#undo) | Undoes the last edit control operation in the control's undo queue. |

---

# <a name="constructors"></a>Constructors

Creates instances of the `CEdit` class.

```
CONSTRUCTOR (BYVAL hCtl AS HWND)
CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
```
| Parameter  | Description |
| ---------- | ----------- |
| *hCtl* | Handle of the edit control. |
| *hParent* | Handle of the parent window of the edit control. |
| *cID* | Control identifier of the edit control. |
| *pDlg* | Pointer to an instance of the `CDialog` class. |

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM hCtl AS HWND = pDlg->ControlHandle(103)
DIM pEdit AS CEdit = hCtl
pEdit.SetText "Text string"
```
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.SetText "Text string"
```
```
DIM hCtl AS HWND = pDlg->ControlHandle(103)
Cedit(hCtl).SetText("Text string")
```
```
Cedit(pDlg, 103).SetText("Text string")
```
---

### <a name="canundo"></a>CanUndo

Determines whether there are any actions in an edit control's undo queue.
```
FUNCTION CanUndo () AS BOOLEAN
```
#### Return value

If there are actions in the control's undo queue, the return value is true. If the undo queue is empty, the return value is false.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
DIM b AS BOOLEAN = pEdit.CanUndo
```
```
CEdit(pDlg, 103).CanUndo
```
---

### <a name="charfrompos"></a>CharFromPos

Gets information about the character closest to a specified point in the client area of an edit control.
```
FUNCTION CharFromPos (BYVAL x AS SHORT, BYVAL y AS SHORT) AS DWORD
```
| Parameter | Description |
| --------- | ----------- |
| x | The x-coordinate of a point in the control's client area. |
| y | The y-coordinate of a point in the control's client area. |

#### Remarks

The coordinates are in screen units and are relative to the upper-left corner of the control's client area.
 
#### Return value

The **LOWORD** specifies the zero-based index of the character nearest the specified point. This index is relative to the beginning of the control, not the beginning of the line. If the specified point is beyond the last character in the edit control, the return value indicates the last character in the control. The **HIWORD** specifies the zero-based index of the line that contains the character. For single-line edit controls, this value is zero. The index indicates the line delimiter if the specified point is beyond the last visible character in a line.

---

### <a name="clear"></a>Clear

Sends a **WM_CLEAR** message to an edit control to delete (clear) the current selection, if any, from the edit control.
```
SUB Clear ()
```
#### Return value

This message does not return a value.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.Clear
```
```
CEdit(pDlg, 103).Clear
```
---

### <a name="copy"></a>Copy

Sends the **WM_COPY** message to an edit control to copy the current selection to the clipboard in CF_TEXT format.
```
SUB Copy ()
```
#### Return value

This message does not return a value.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.Copy
```
```
CEdit(pDlg, 103).Copy
```
---

### <a name="cut"></a>Cut

Sends a **WM_CUT** message to an edit control to delete (cut) the current selection, if any, in the edit control and copy the deleted text to the clipboard in CF_TEXT format.
```
SUB Cut ()
```
#### Return value

This message does not return a value.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.Cut
```
```
CEdit(pDlg, 103).Cut
```
---

### <a name="disable"></a>Disable

Disables mouse and keyboard input to the specified window or control. When input is disabled, the window does not receive input such as mouse clicks and key presses.
```
SUB Disable ()
```
#### Return value

Returns FALSE if the control was previously disabled; otherwise TRUE.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.Disable
```
```
CEdit(pDlg, 103).Disable
```
#### Remarks

If the window is being disabled, the system sends a **WM_CANCELMODE** message. If the enabled state of a window is changing, the system sends a **WM_ENABLE** message after the **WM_CANCELMODE** message. If a window is already disabled, its child windows are implicitly disabled, although they are not sent a **WM_ENABLE** message.

A window must be enabled before it can be activated. For example, if an application is displaying a modeless dialog box and has disabled its main window, the application must enable the main window before destroying the dialog box. Otherwise, another window will receive the keyboard focus and be activated. If a child window is disabled, it is ignored when the system tries to determine which window should receive mouse messages.

By default, a window is enabled when it is created. To create a window that is initially disabled, an application can specify the **WS_DISABLED** style in the **CreateWindow** or **CreateWindowEx** function. After a window has been created, an application can use **EnableWindow** API function to enable or disable the window.

An application can use this function to disable a control in a dialog box. A disabled control cannot receive the keyboard focus, nor can a user gain access to it.

---

### <a name="enable"></a>Enable

Enables mouse and keyboard input to the specified window or control. When input is enabled, the window receives all input.
```
SUB Enable ()
```
#### Return value

Returns FALSE if the control was previously enabled; otherwise TRUE.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.Enable
```
```
CEdit(pDlg, 103).Enable
```

#### Remarks

If the window is being disabled, the system sends a **WM_CANCELMODE** message. If the enabled state of a window is changing, the system sends a **WM_ENABLE** message after the **WM_CANCELMODE** message. If a window is already disabled, its child windows are implicitly disabled, although they are not sent a **WM_ENABLE** message.

A window must be enabled before it can be activated. For example, if an application is displaying a modeless dialog box and has disabled its main window, the application must enable the main window before destroying the dialog box. Otherwise, another window will receive the keyboard focus and be activated. If a child window is disabled, it is ignored when the system tries to determine which window should receive mouse messages.

By default, a window is enabled when it is created. To create a window that is initially disabled, an application can specify the **WS_DISABLED** style in the **CreateWindow** or **CreateWindowEx** function. After a window has been created, an application can use the **EnableWindow** API function to enable or disable the window.

An application can use this method to enable a control in a dialog box.

---

### <a name="emptyundobuffer"></a>EmptyUndoBuffer

Resets the undo flag of an edit control. The undo flag is set whenever an operation within the edit control can be undone. 
```
SUB EmptyUndoBuffer ()
```
#### Return value

This method does not return a value.

#### Remarks

This method empties all undo and redo buffers. The undo flag is automatically reset whenever the edit control receives a **WM_SETTEXT** or **EM_SETHANDLE** message.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.EmptyUndoBuffer
```
```
CEdit(pDlg, 103).EmptyUndoBuffer
```
