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
| [LineIndex](#lineindex) | Gets the character index of the first character of a specified line in a multiline edit control. |
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
---

### <a name="fmtlines"></a>FmtLines

Sets a flag that determines whether a multiline edit control includes soft line-break characters. A soft line break consists of two carriage returns and a line feed and is inserted at the end of a line that is broken because of wordwrapping.
```
SUB FmtLines (BYVAL AddEolFlag AS BOOLEAN) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *AddEolFlag* | A value of TRUE inserts the characters; a value of FALSE removes them. |

#### Return value

The return value is identical to the *AddEolFlag* parameter.

#### Remarks

This message affects only the buffer returned by the **EM_GETHANDLE** message and the text returned by the **WM_GETTEXT** message. It has no effect on the display of the text within the edit control.

The **EM_FMTLINES** message does not affect a line that ends with a hard line break. A hard line break consists of one carriage return and a line feed.

The size of the text changes when this message is processed.

#### Usage examples

Note: 103 is the identifier of the edit control. Change it to the real one.
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.FmtLines(TRUE)
```
```
CEdit(pDlg, 103).FmtLines(TRUE)
```
---

### <a name="getcuebannertext"></a>GetCueBannerText

Gets the text that is displayed as the textual cue, or tip, in an edit control.
```
FUNCTION GetCueBannerText (BYVAL pwszText AS WSTRING PTR, BYVAL cchText AS LONG) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pwszText* | A pointer to a Unicode buffer that receives the text set as the textual cue. The caller is responsible for allocating the buffer. |
| *cchText* | The size of the buffer pointed to by *pwszText* in characters, including the terminating NULL. |

#### Return value

Returns TRUE if successful or FALSE otherwise.

#### Remarks

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0.

---

### <a name="getfirstvisibleline"></a>GetFirstVisibleLine

Gets the zero-based index of the uppermost visible line in a multiline edit control. You can send this message to either an edit control or a rich edit control.
```
FUNCTION GetFirstVisibleLine () AS LONG
```

#### Return value

The return value is the zero-based index of the uppermost visible line in a multiline edit control.

For single-line edit controls, the return value is the zero-based index of the first visible character.

#### Remarks

The number of lines and the length of the lines in an edit control depend on the width of the control and the current Wordwrap setting.

---

### <a name="gethandle"></a>GetHandle

Gets a handle of the memory currently allocated for a multiline edit control's text.
```
FUNCTION GetHandle () AS HLOCAL
```
#### Return value

The return value is a memory handle identifying the buffer that holds the content of the edit control. If an error occurs, such as sending the message to a single-line edit control, the return value is zero.

#### Remarks

If the function succeeds, the application can access the contents of the edit control by casting the return value to **HLOCAL** and passing it to **LocalLock**. **LocalLock** returns a pointer to a buffer that is a null-terminated array of characters. For example, The application may not change the contents of the buffer. To unlock the buffer, the application calls **LocalUnlock** before allowing the edit control to receive new messages.

If your application cannot abide by the restrictions imposed by **EM_GETHANDLE**, use the **GetWindowTextLength** and **GetWindowText** functions to copy the contents of the edit control into an application-provided buffer.

---

### <a name="getimestatus"></a>GetIMEStatus

Gets a set of status flags that indicate how the edit control interacts with the Input Method Editor (IME).
```
FUNCTION GetIMEStatus (BYVAL nStatusType AS LONG) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *nStatusType* | The type of status to retrieve. This parameter can be the following value:<br>**EMSIS_COMPOSITIONSTRING**: Sets behavior for handling the composition string. |

#### Return value

Data specific to the type of status to retrieve. With the **EMSIS_COMPOSITIONSTRING** value for *nStatusType*, this return value is one or more of the following values.

| Return code | Description |
| ----------- | ----------- |
| **EIMES_GETCOMPSTRATONCE** | If this flag is set, the edit control hooks the **WM_IME_COMPOSITION** message with fFlags set to **GCS_RESULTSTR** and returns the result string immediately. If this flag is not set, the edit control passes the **WM_IME_COMPOSITION** message to the default window procedure and processes the result string from the **WM_CHAR** message; this is the default behavior of the edit control. |
| **EIMES_CANCELCOMPSTRINFOCUS** | If this flag is set, the edit control cancels the composition string when it receives the **WM_SETFOCUS** message. If this flag is not set, the edit control does not cancel the composition string; this is the default behavior of the edit control. |
| **EIMES_COMPLETECOMPSTRKILLFOCUS** | If this flag is set, the edit control completes the composition string upon receiving the **WM_KILLFOCUS** message. If this flag is not set, the edit control does not complete the composition string; this is the default behavior of the edit control. |

---

### <a name="getlimittext"></a>GetLimitText

Gets a set of status flags that indicate how the edit control interacts with the Input Method Editor (IME).
```
FUNCTION GetLimitText () AS DWORD
```
#### Return value

The return value is the text limit.

#### Remarks

The text limit is the maximum amount of text, in characters, that the control can contain.

---


### <a name="getline"></a>GetLine

Copies a line of text from an edit control and places it in a specified buffer. You can send this message to either an edit control or a rich edit control.
```
FUNCTION GetLine (BYVAL which AS DWORD) AS DWSTRING
```
| Parameter | Description |
| --------- | ----------- |
| *which* | The zero-based index of the line to retrieve from a multiline edit control. A value of zero specifies the topmost line. This parameter is ignored by a single-line edit control. |

#### Return value

The return value is the number of characters copied. The return value is zero if the line number specified by the *which* parameter is greater than the number of lines in the edit control.

#### Remarks

The copied line does not contain a terminating null character.

---

### <a name="getlinecount"></a>GetLineCount

Gets the number of lines in a multiline edit control.
```
FUNCTION GetLineCount () AS LONG
```

#### Return value

The return value is an integer specifying the total number of text lines in the multiline edit control or rich edit control. If the control has no text, the return value is 1. The return value will never be less than 1.

#### Remarks

Retrieves the total number of text lines, not just the number of lines that are currently visible.

If the Wordwrap feature is enabled, the number of lines can change when the dimensions of the editing window change.

---

### <a name="getmargins"></a>GetMargins

Gets the widths of the left and right margins for an edit control.
```
FUNCTION GetMargins () AS DWORD
```

#### Return value

Returns the width of the left margin in the LOWORD, and the width of the right margin in the HIWORD.

---

### <a name="getleftmargin"></a>GetLeftMargin

Gets the width of the left margin for an edit control.
```
FUNCTION GetLeftMargin () AS WORD
```

#### Return value

Returns the width of the left margin.

---

### <a name="getrightmargin"></a>GetRigthMargin

Gets the width of the right margin for an edit control.
```
FUNCTION GetRightMargin () AS WORD
```

#### Return value

Returns the width of the right margin.

---

### <a name="getpasswordchar"></a>GetPasswordChar

Gets the password character that an edit control displays when the user enters text. 
```
FUNCTION GetPasswordChar () AS LONG
```

#### Return value

The return value specifies the character to be displayed in place of any characters typed by the user. If the return value is NULL, there is no password character, and the control displays the characters typed by the user.

####Remarks

If an edit control is created with the **ES_PASSWORD** style, the default password character is set to an asterisk (*). If an edit control is created without the **ES_PASSWORD** style, there is no password character. To change the password character, send the **EM_SETPASSWORDCHAR** message.

Multiline edit controls do not support the password style or messages.

---

### <a name="getrect"></a>GetRect

Gets the password character that an edit control displays when the user enters text. 
```
SUB GetRect (BYREF rc AS RECT)
FUNCTION GetRect () AS RECT
```
| Parameter | Description |
| --------- | ----------- |
| *rc* | A pointer to a **RECT** structure that receives the formatting rectangle. |

#### Return value

Gets the formatting rectangle of an edit control. The formatting rectangle is the limiting rectangle into which the control draws the text. The limiting rectangle is independent of the size of the edit-control window.

---

### <a name="getsel"></a>GetSel

Gets the starting and ending character positions (in characters) of the current selection in an edit control.
```
FUNCTION GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
FUNCTION GetSel () AS POINTL
```
| Parameter | Description |
| --------- | ----------- |
| *dwStartPos* | A pointer to a DWORD value that receives the starting position of the selection. This parameter can be NULL. |
| *dwEndPos* | A pointer to a DWORD value that receives the position of the first unselected character after the end of the selection. This parameter can be NULL. |

#### Return value

The starting and ending character positions (in characters) of the current selection in an edit control.

#### Remarks

If there is no selection, the starting and ending values are both the position of the caret.

---

### <a name="getselstart"></a>GetSelStart

Gets the starting character position of the current selection in an edit control.
```
FUNCTION GetSelStart () AS DWORD
```

#### Return value

The starting character position of the current selection in an edit control.

#### Remarks

If there is no selection, the starting and ending values are both the position of the caret.

---

### <a name="getselend"></a>GetSelEnd

Gets the ending character position of the current selection in an edit control.
```
FUNCTION GetSelEnd () AS DWORD
```
#### Return value

The ending character position of the current selection in an edit control.

#### Remarks

If there is no selection, the starting and ending values are both the position of the caret.

---

### <a name="gettext"></a>GetText

Retrieves the text in a edit control.
```
FUNCTION GetText () AS DWSTRING
```
#### Return value

The retrieved text.

#### Usage examples

```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.Gettext
```
```
CEdit(pDlg, 103).GetText
```
---

### <a name="gettextlength"></a>GetTextLength

Retrieves the length of the text in a edit control.
```
FUNCTION GetTextLength () AS DWSTRING
```
#### Return value

The length of the text in a edit control.

#### Usage examples

```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
DIM nLen AS LONG = pEdit.GetTextLength
```
```
DIM nLen AS LONG = CEdit(pDlg, 103).GetTextLength
```
---

### <a name="getthumb"></a>GetThumb

Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline edit control.
```
FUNCTION GetThumb () AS LONG
```
#### Return value

The position of the scroll box.

#### Usage examples

```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
DIM nPos AS LONG = pEdit.GetThumb
```
```
DIM nPos AS LONG = CEdit(pDlg, 103).GetThumb
```
---

### <a name="getwordbreakproc"></a>GetWordBreakProc

Gets the address of the current Wordwrap function.
```
FUNCTION GetWordBreakProc () AS LONG_PTR
```
#### Return value

The address of the current Wordwrap function.

---

### <a name="hideballoontip"></a>HideBalloonTip

Hides any balloon tip associated with an edit control.
```
FUNCTION HideBalloonTip () AS BOOLEAN
```
#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0.

---

### <a name="limittext"></a>LimitText

Sets the text limit of an edit control. The text limit is the maximum amount of text, in characters, that the user can type into the edit control.
```
SUB LimitText (BYVAL chMax AS DWORD)
```
#### Return value

This method does not return a value.

#### Usage examples
```
DIM pEdit AS CEdit = CEdit(pDlg, 103)
pEdit.LimitText(50)
```
```
CEdit(pDlg, 103).LimitText(50)
```
---

### <a name="linefromchar"></a>LineFromChar

Gets the index of the line that contains the specified character index in a multiline edit control. A character index is the zero-based index of the character from the beginning of the edit control. 
```
FUNCTION LineFromChar (BYVAL index AS LONG) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *index* | The character index of the character contained in the line whose number is to be retrieved. If this parameter is -1, it retrieves either the line number of the current line (the line containing the caret) or, if there is a selection, the line number of the line containing the beginning of the selection. |

#### Return value

The return value is the zero-based line number of the line containing the character index specified by *index*.

---

### <a name="lineindex"></a>LineIndex

Gets the character index of the first character of a specified line in a multiline edit control. A character index is the zero-based index of the character from the beginning of the edit control.
```
FUNCTION LineIndex (BYVAL nLine AS LONG) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *nLine* | The zero-based line number. A value of -1 specifies the current line number (the line that contains the caret). |

#### Return value

The return value is the character index of the line specified in the *nLine* parameter, or it is -1 if the specified line number is greater than the number of lines in the edit control.

---

### <a name="linelength"></a>LineLength

Retrieves the length, in characters, of a line in an edit control.
```
FUNCTION LineLength (BYVAL index AS LONG) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *index* | The character index of a character in the line whose length is to be retrieved. If this parameter is greater than the number of characters in the control, the return value is zero.This parameter can be -1. In this case, the message returns the number of unselected characters on lines containing selected characters. For example, if the selection extended from the fourth character of one line through the eighth character from the end of the next line, the return value would be 10 (three characters on the first line and seven on the next). |

#### Return value

The length, in characters, of a line in an edit control.

---
