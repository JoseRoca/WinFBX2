' ########################################################################################
' Platform: Microsoft Windows
' Filename: CEdit.inc
' Purpose:  Windows Edit control
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#if not defined(UNICODE)
   #define UNICODE
#endif
#if not defined(_WIN32_WINNT)
   #define _WIN32_WINNT &h0602
#endif

#pragma once
#include once "windows.bi"
#include once "win/commctrl.bi"
#include once "Afx2/CControl.inc"
#include once "Afx2/CDialog.inc"
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CEdit class
' ########################################################################################

TYPE CEdit EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION DluToPixRX () AS SINGLE
   DECLARE FUNCTION DluToPixRY () AS SINGLE
   DECLARE FUNCTION PixToDluRX () AS SINGLE
   DECLARE FUNCTION PixToDluRY () AS SINGLE
   DECLARE FUNCTION DluToPixX (BYVAL dlu AS LONG) AS SINGLE
   DECLARE FUNCTION DluToPixY (BYVAL dlu AS LONG) AS SINGLE
   DECLARE FUNCTION PixToDluX (BYVAL pix AS LONG) AS SINGLE
   DECLARE FUNCTION PixToDluY (BYVAL pix AS LONG) AS SINGLE

   DECLARE FUNCTION CanUndo () AS BOOLEAN
   DECLARE FUNCTION CharFromPos (BYVAL x AS SHORT, BYVAL y AS SHORT) AS DWORD
   DECLARE SUB Clear ()
   DECLARE SUB Copy ()
   DECLARE SUB Cut ()
   DECLARE SUB EmptyUndoBuffer ()
   DECLARE FUNCTION Disable () AS BOOLEAN
   DECLARE FUNCTION Enable () AS BOOLEAN
   DECLARE FUNCTION FmtLines (BYVAL AddEolFlag AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION GetCueBannerText (BYVAL pwszText AS WSTRING PTR, BYVAL cchText AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetFirstVisibleLine () AS LONG
   DECLARE FUNCTION GetIMEStatus (BYVAL nStatusType AS LONG) AS LONG
   DECLARE FUNCTION GetHandle () AS ..HLOCAL
   DECLARE FUNCTION GetLimitText () AS DWORD
   DECLARE FUNCTION GetLine (BYVAL which AS DWORD) AS DWSTRING
   DECLARE FUNCTION GetLineCount () AS LONG
   DECLARE FUNCTION GetMargins () AS DWORD
   DECLARE FUNCTION GetLeftMargin () AS WORD
   DECLARE FUNCTION GetRightMargin () AS WORD
   DECLARE FUNCTION GetModify () AS BOOLEAN
   DECLARE FUNCTION GetPasswordChar () AS LONG
   DECLARE SUB GetRect (BYREF rc AS RECT)
   DECLARE FUNCTION GetRect () AS RECT
   DECLARE FUNCTION GetSel () AS ..POINTL
   DECLARE FUNCTION GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
   DECLARE FUNCTION GetSelStart () AS DWORD
   DECLARE FUNCTION GetSelEnd () AS DWORD
   DECLARE FUNCTION GetTextLength () AS LONG
   DECLARE FUNCTION GetText () AS DWSTRING
   DECLARE FUNCTION GetThumb () AS LONG
   DECLARE FUNCTION GetWordBreakProc () AS LONG_PTR
   DECLARE FUNCTION HideBalloonTip () AS BOOLEAN
   DECLARE SUB LimitText (BYVAL chMax AS DWORD)
   DECLARE FUNCTION LineFromChar (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION LineLength (BYVAL index AS LONG) AS LONG
   DECLARE FUNCTION LineScroll (BYVAL x AS LONG, BYVAL y AS LONG) AS BOOLEAN
   DECLARE SUB Paste ()
   DECLARE FUNCTION PosFromChar (BYVAL index AS LONG) AS LONG
   DECLARE SUB ReplaceSel (BYVAL bCanBeUndone AS BOOLEAN, BYVAL pwszText AS WSTRING PTR)
   DECLARE FUNCTION Scroll (BYVAL nAction AS LONG) AS DWORD
   DECLARE SUB ScrollCaret ()
   DECLARE FUNCTION SetCueBannerText (BYVAL pwszText AS WSTRING PTR, BYVAL bDrawFocused AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION SetCueBannerTextFocused (BYVAL pwszText AS WSTRING PTR, BYVAL bDrawFocused AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION SetIMEStatus (BYVAL nStatusType AS LONG, BYVAL flags AS DWORD) AS DWORD
   DECLARE SUB SetHandle (BYVAL hLocal AS ..HLOCAL)
   DECLARE SUB SetLimitText (BYVAL chMax AS DWORD)
   DECLARE SUB SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
   DECLARE SUB SetLeftMargin (BYVAL nWidth AS WORD)
   DECLARE SUB SetRightMargin (BYVAL nWidth AS WORD)
   DECLARE SUB SetModify (BYVAL bModify AS BOOLEAN)
   DECLARE SUB SetPasswordChar (BYVAL dwchar AS DWORD)
   DECLARE FUNCTION SetReadOnly (BYVAL bReadOnly AS BOOLEAN) AS LONG
   DECLARE SUB SetRect (BYREF prect AS RECT)
   DECLARE SUB SetRectNoPaint (BYREF prect AS RECT)
   DECLARE SUB SetSel (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
   DECLARE FUNCTION SetTabStops (BYVAL cTabs AS LONG, BYVAL prgTabStops AS LONG PTR) AS BOOLEAN
   DECLARE FUNCTION SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   DECLARE SUB SetWordBreakProc (BYVAL pfn AS LONG_PTR)
   DECLARE FUNCTION ShowBalloonTip (BYREF peditballoontip AS EDITBALLOONTIP) AS BOOLEAN
   DECLARE FUNCTION Undo () AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CEdit constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CEdit (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CEdit (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CEdit (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CEdit (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CEdit destructor
' ========================================================================================
PRIVATE DESTRUCTOR CEdit
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Retrieves the conversion ratio from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION CEdit.DluToPixRX () AS SINGLE
   RETURN Base.DluToPixRX(m_hCtl)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION CEdit.DluToPixRY () AS SINGLE
   RETURN Base.DluToPixRY(m_hCtl)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION CEdit.PixToDluRX () AS SINGLE
   RETURN Base.PixToDluRX(m_hCtl)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the conversion ratio from pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION CEdit.PixToDluRY () AS SINGLE
   RETURN Base.PixToDluRY(m_hCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts dialog units to pixels
' ========================================================================================
PRIVATE FUNCTION CEdit.DluToPixX (BYVAL dlu AS LONG) AS SINGLE
   RETURN Base.DluToPixX(m_hCtl, dlu)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CEdit.DluToPixY (BYVAL dlu AS LONG) AS SINGLE
   RETURN Base.DluToPixY(m_hCtl, dlu)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Converts pixels to dialog units
' ========================================================================================
PRIVATE FUNCTION CEdit.PixToDluX (BYVAL pix AS LONG) AS SINGLE
   RETURN Base.PixToDluX(m_hCtl, pix)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CEdit.PixToDluY (BYVAL pix AS LONG) AS SINGLE
   RETURN Base.PixToDluY(m_hCtl, pix)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                      *** EDIT ***
' ########################################################################################

' ========================================================================================
' * Determines whether there are any actions in an edit control's undo queue. You can send
' this message to either an edit control or a rich edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.CanUndo () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_CANUNDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about the character closest to a specified point in the client area of
' an edit control. You can send this message to either an edit control or a rich edit control.
' Return value: The LOWORD specifies the zero-based index of the character nearest the
' specified point. This index is relative to the beginning of the control, not the beginning
' of the line. If the specified point is beyond the last character in the edit control, the
' return value indicates the last character in the control. The HIWORD specifies the
' zero-based index of the line that contains the character. For single-line edit controls,
' this value is zero. The index indicates the line delimiter if the specified point is
' beyond the last visible character in a line.
' ========================================================================================
PRIVATE FUNCTION CEdit.CharFromPos (BYVAL x AS SHORT, BYVAL y AS SHORT) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, EM_CHARFROMPOS, 0, cast(LPARAM, MAKELONG(x, y)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes (clears) the current selection, if any, from the edit control.
' ========================================================================================
PRIVATE SUB CEdit.Clear ()
   SendMessageW(m_hCtl, WM_CLEAR, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' * Copies the current selection to the clipboard in CF_TEXT format.
' ========================================================================================
PRIVATE SUB CEdit.Copy ()
   SendMessageW(m_hCtl, WM_COPY, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' * Deletes (cuts) the current selection, if any, in the edit control and copy the deleted
' text to the clipboard in CF_TEXT format.
' ========================================================================================
PRIVATE SUB CEdit.Cut ()
   SendMessageW(m_hCtl, WM_CUT, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Resets the undo flag of an edit control. The undo flag is set whenever an operation
' within the edit control can be undone.
' ========================================================================================
SUB CEdit.EmptyUndoBuffer ()
   SendMessageW m_hCtl, EM_EMPTYUNDOBUFFER, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' * Enables a button.
' Return value: False if the windows was previous enabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CEdit.Enable () AS BOOLEAN
   RETURN EnableWindow(m_hCtl, TRUE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Disables a button.
' Return value: False if the windows was previous disabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CEdit.Disable () AS BOOLEAN
   DIM res AS LONG = EnableWindow(m_hCtl, FALSE)
   DIM bEnable AS BOOLEAN
   IF res = 0 THEN bEnable = TRUE ELSE bEnable = FALSE
   RETURN bEnable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a flag that determines whether a multiline edit control includes soft line-break
' characters. A soft line break consists of two carriage returns and a line feed and is
' inserted at the end of a line that is broken because of wordwrapping.
' ========================================================================================
PRIVATE FUNCTION CEdit.FmtLines (BYVAL AddEolFlag AS BOOLEAN) AS BOOLEAN
   SendMessageW(m_hCtl, EM_FMTLINES, CAST(WPARAM, AddEolFlag), 0)
   RETURN AddEolFlag
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text that is displayed as the textual cue, or tip, in an edit control.
' To use this method, you must provide a manifest specifying Comctl32.dll version 6.0.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetCueBannerText (BYVAL pwszText AS WSTRING PTR, BYVAL cchText AS LONG) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_GETCUEBANNER, CAsT(WPARAM, pwszText), cchText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the zero-based index of the uppermost visible line in a multiline edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetFirstVisibleLine () AS LONG
   RETURN SendMessageW(m_hCtl, EM_GETFIRSTVISIBLELINE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a handle of the memory currently allocated for a multiline edit control's text.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetHandle () AS ..HLOCAL
   RETURN CAST(..HLOCAL, SendMessageW(m_hCtl, EM_GETHANDLE, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE FUNCTION CEdit.GetHilite () AS LONG
'   RETURN SendMessageW(m_hCtl, EM_GETHILITE, 0, 0)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a set of status flags that indicate how the edit control interacts with the
' Input Method Editor (IME).
' ========================================================================================
PRIVATE FUNCTION CEdit.GetIMEStatus (BYVAL nStatusType AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, EM_GETIMESTATUS, cast(WPARAM, nStatusType), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text limit for an edit control. The return value is the text limit.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetLimitText () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, EM_GETLIMITTEXT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a line of text from an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetLine (BYVAL which AS DWORD) AS DWSTRING
   DIM buffer AS DWSTRING
   DIM n AS LONG
   buffer = MKI(32765) + WSTRING(32765, 0)
   n = SendMessageW(m_hCtl, EM_GETLINE, which, CAST(LPARAM, STRPTR(buffer)))
   RETURN LEFT(buffer, n)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of lines in a multiline edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetLineCount () AS LONG
   RETURN SendMessageW(m_hCtl, EM_GETLINECOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width of the left margin in the LOWORD, and the width of the right margin in the HIWORD.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetMargins () AS DWORD
   FUNCTION = SendMessageW(m_hCtl, EM_GETMARGINS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width of the left margin for an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetLeftMargin () AS WORD
   FUNCTION = LOWORD(SendMessageW(m_hCtl, EM_GETMARGINS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the right margin for an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetRightMargin () AS WORD
   FUNCTION = HIWORD(SendMessageW(m_hCtl, EM_GETMARGINS, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the state of an edit control's modification flag. The flag indicates whether the
' contents of the edit control have been modified.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetModify () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_GETMODIFY, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the password character that an edit control displays when the user enters text.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetPasswordChar () AS LONG
   RETURN SendMessageW(m_hCtl, EM_GETPASSWORDCHAR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the formatting rectangle of an edit control. The formatting rectangle is the
' limiting rectangle into which the control draws the text. The limiting rectangle is
' independent of the size of the edit-control window.
' ========================================================================================
PRIVATE SUB CEdit.GetRect (BYREF rc AS RECT)
   SendMessageW m_hCtl, EM_GETRECT, 0, CAST(LPARAM, @rc)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CEdit.GetRect () AS RECT
   DIM rc AS RECT
   SendMessageW m_hCtl, EM_GETRECT, 0, CAST(LPARAM, @rc)
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the starting and ending character positions of the current selection in an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetSel (BYREF dwStartPos AS DWORD, BYREF dwEndPos AS DWORD) AS LONG
   RETURN SendMessageW(m_hCtl, EM_GETSEL, CAST(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CEdit.GetSel () AS ..POINTL
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD, pt AS .POINTL
   SendMessageW(m_hCtl, EM_GETSEL, CAST(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
   pt.x = dwStartPos : pt.y = dwEndPos
   RETURN pt
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the starting character position of the current selection in an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetSelStart () AS DWORD
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD
   SendMessageW(m_hCtl, EM_GETSEL, CAST(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
   RETURN dwStartPos
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the ending character position of the current selection in an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetSelEnd () AS DWORD
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD
   SendMessageW(m_hCtl, EM_GETSEL, CAST(WPARAM, @dwStartPos), cast(LPARAM, @dwEndPos))
   RETURN dwEndPos
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the text length from an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetTextLength () AS LONG
   RETURN GetWindowTextLengthW(m_hCtl)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Retrieves the text in a button control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetText () AS DWSTRING
   DIM nLen AS LONG = GetWindowTextLengthW(m_hCtl)
   DIM dwsText AS DWSTRING = WSPACE(nLen + 1)
   ..GetWindowTextW(m_hCtl, dwsText, nLen +  1)
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline
' edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetThumb () AS LONG
   FUNCTION = SendMessageW(m_hCtl, EM_GETTHUMB, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the address of the currently registered word-break procedure.
' ========================================================================================
PRIVATE FUNCTION CEdit.GetWordBreakProc () AS LONG_PTR
   RETURN SendMessageW(m_hCtl, EM_GETWORDBREAKPROC, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides any balloon tip associated with an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.HideBalloonTip () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_HIDEBALLOONTIP, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text limit of an edit control. The text limit is the maximum amount of text, in
' bytes, that the user can type into the edit control.
' ========================================================================================
PRIVATE SUB CEdit.LimitText (BYVAL chMax AS DWORD)
   SendMessageW m_hCtl, EM_LIMITTEXT, chMax, 0
END SUB
' ========================================================================================

' ========================================================================================
' Gets the index of the line that contains the specified character index in a multiline
' edit control. A character index is the zero-based index of the character from the
' beginning of the edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.LineFromChar (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, EM_LINEFROMCHAR, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length, in characters, of a line in an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.LineLength (BYVAL index AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, EM_LINELENGTH, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the text in a multiline edit control.
' If the message is sent to a multiline edit control, the return value is TRUE.
' If the message is sent to a single-line edit control, the return value is FALSE.
' ========================================================================================
PRIVATE FUNCTION CEdit.LineScroll (BYVAL x AS LONG, BYVAL y AS LONG) AS BOOLEAN
   FUNCTION = SendMessageW(m_hCtl, EM_LINESCROLL, cast(WPARAM, x), cast(LPARAM, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the current content of the clipboard to the edit control at the current caret
' position. Data is inserted only if the clipboard contains data in CF_TEXT format.
' ========================================================================================
PRIVATE SUB CEdit.Paste ()
   SendMessageW(m_hCtl, WM_PASTE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the client area coordinates of a specified character in an edit control.
' The return value contains the client area coordinates of the character. The LOWORD
' contains the horizontal coordinate and the HIWORD contains the vertical coordinate.
' ========================================================================================
PRIVATE FUNCTION CEdit.PosFromChar (BYVAL index AS LONG) AS LONG
   FUNCTION = SendMessageW(m_hCtl, EM_POSFROMCHAR, cast(WPARAM, index), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the current selection in an edit control with the specified text.
' ========================================================================================
PRIVATE SUB CEdit.ReplaceSel (BYVAL bCanBeUndone AS BOOLEAN, BYVAL pwszText AS WSTRING PTR)
   SendMessageW m_hCtl, EM_REPLACESEL, CAST(WPARAM, bCanBeUndone), CAST(LPARAM, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Scrolls the text vertically in a multiline edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.Scroll (BYVAL nAction AS LONG) AS DWORD
   RETURN SendMessageW(m_hCtl, EM_SCROLL, nAction, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the caret into view in an edit control.
' ========================================================================================
PRIVATE SUB CEdit.ScrollCaret ()
   SendMessageW m_hCtl, EM_SCROLLCARET, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the text that is displayed as the textual cue, or tip, for an edit control.
' - pwszText: A pointer to a Unicode string that contains the text to display as the textual cue.
' - bDrawFocused: TRUE if the cue banner should show even when the edit control has focus;
'   otherwise, FALSE. FALSE is the default behavior—the cue banner disappears when the user
'   clicks in the control.
' If the message succeeds, it returns TRUE. Otherwise it returns FALSE.
' ========================================================================================
PRIVATE FUNCTION CEdit.SetCueBannerText (BYVAL pwszText AS WSTRING PTR, BYVAL bDrawFocused AS BOOLEAN = FALSE) AS LONG
   RETURN SendMessageW(m_hCtl, EM_SETCUEBANNER, CAST(WPARAM, bDrawFocused), CAST(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the handle of the memory that will be used by a multiline edit control.
' ========================================================================================
PRIVATE SUB CEdit.SetHandle (BYVAL hLocal AS ..HLOCAL)
   SendMessageW m_hCtl, EM_SETHANDLE, CAST(WPARAM, hLocal), 0
END SUB
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE FUNCTION CEdit.SetHilite (BYVAL ichStart AS LONG, BYVAL ichEnd AS LONG) AS LONG
'   RETURN SendMessageW(m_hCtl, EM_SETHILITE, ichStart, ichEnd)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the status flags that determine how an edit control interacts with the Input Method
' Editor (IME).
' - nStatusType: The type of status to set. This parameter can be EMSIS_COMPOSITIONSTRING.
' - flags: Data specific to the status type. If wParam is EMSIS_COMPOSITIONSTRING, this
'   parameter can be one or more of the following values: EIMES_GETCOMPSTRATONCE,
'   EIMES_CANCELCOMPSTRINFOCUS, EIMES_COMPLETECOMPSTRKILLFOCUS.
' Returns the previous value of the flags parameter.
' ========================================================================================
PRIVATE FUNCTION CEdit.SetIMEStatus (BYVAL nStatusType AS LONG, BYVAL flags AS DWORD) AS DWORD
   FUNCTION = SendMessageW(m_hCtl, EM_SETIMESTATUS, CAST(WPARAM, nStatusType), CAST(LPARAM, flags))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text limit of an edit control. The text limit is the maximum amount of text, in
' TCHARs, that the user can type into the edit control.
' ========================================================================================
PRIVATE SUB CEdit.SetLimitText (BYVAL chMax AS DWORD)
   SendMessageW m_hCtl, EM_LIMITTEXT, chMax, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the widths of the left and right margins for an edit control. The message redraws
' the control to reflect the new margins.
' - nMargins: One or more of the following values:
'   EC_LEFTMARGIN: Sets the left margin.
'   EC_RIGHTMARGIN: Sets the right margin.
' - nWidth: The low-order word specifies the new width of the left margin, in pixels.
'   The high-order word specifies the new width of the right margin, in pixels.
' ========================================================================================
PRIVATE SUB CEdit.SetMargins (BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
   SendMessageW(m_hCtl, EM_SETMARGINS, nMargins, nWidth)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the width of the left margin for an edit control. The message redraws the control
' to reflect the new margin.
' ========================================================================================
PRIVATE SUB CEdit.SetLeftMargin (BYVAL nWidth AS WORD)
   SendMessageW(m_hCtl, EM_SETMARGINS, EC_LEFTMARGIN, MAKELONG(nWidth, 0))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the width of the right margin for an edit control. The message redraws the control
' to reflect the new margin.
' ========================================================================================
PRIVATE SUB CEdit.SetRightMargin (BYVAL nWidth AS WORD)
   SendMessageW(m_hCtl, EM_SETMARGINS, EC_RIGHTMARGIN, MAKELONG(0, nWidth))
END SUB
' ========================================================================================

' ========================================================================================
' Sets or clears the modification flag for an edit control. The modification flag
' indicates whether the text within the edit control has been modified.
' ========================================================================================
PRIVATE SUB CEdit.SetModify (BYVAL bModify AS BOOLEAN)
   SendMessageW m_hCtl, EM_SETMODIFY, CAST(WPARAM, bModify), 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets or removes the password character for an edit control. When a password character is
' set, that character is displayed in place of the characters typed by the user.
' ========================================================================================
PRIVATE SUB CEdit.SetPasswordChar (BYVAL dwchar AS DWORD)
   SendMessageW m_hCtl, EM_SETPASSWORDCHAR, dwchar, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets or removes the read-only style (ES_READONLY) of an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.SetReadOnly (BYVAL bReadOnly AS BOOLEAN) AS LONG
   RETURN SendMessageW(m_hCtl, EM_SETREADONLY, CAST(WPARAM, bReadOnly), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the formatting rectangle of a multiline edit control. The formatting rectangle is
' the limiting rectangle into which the control draws the text. The limiting rectangle is
' independent of the size of the edit control window. This message is processed only by
' multiline edit controls.
' ========================================================================================
PRIVATE SUB CEdit.SetRect (BYREF rc AS RECT)
   SendMessageW m_hCtl, EM_SETRECT, 0, CAST(LPARAM, @rc)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the formatting rectangle of a multiline edit control. The EM_SETRECTNP message is
' identical to the EM_SETRECT message, except that EM_SETRECTNP does not redraw the edit
' control window. The formatting rectangle is the limiting rectangle into which the
' control draws the text. The limiting rectangle is independent of the size of the edit
' control window. This message is processed only by multiline edit controls.
' ========================================================================================
PRIVATE SUB CEdit.SetRectNoPaint (BYREF rc AS RECT)
   SendMessageW m_hCtl, EM_SETRECTNP, 0, CAST(LPARAM, @rc)
END SUB
' ========================================================================================

' ========================================================================================
' Selects a range of characters in an edit control.
' ========================================================================================
PRIVATE SUB CEdit.SetSel (BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
   SendMessageW m_hCtl, EM_SETSEL, nStart, nEnd
END SUB
' ========================================================================================

' ========================================================================================
' Sets the tab stops in a multiline edit control. When text is copied to the control, any
' tab character in the text causes space to be generated up to the next tab stop. This
' message is processed only by multiline edit controls.
' ========================================================================================
PRIVATE FUNCTION CEdit.SetTabStops (BYVAL cTabs AS LONG, BYVAL prgTabStops AS LONG PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_SETTABSTOPS, cTabs, CAST(LPARAM, prgTabStops))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of an edit control.
' Note: SetWindowText cannot change the text of a control in another application.
' ========================================================================================
PRIVATE FUNCTION CEdit.SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SetWindowText(m_hCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces an edit control's default Wordwrap function with an application-defined
' Wordwrap function.
' ========================================================================================
PRIVATE SUB CEdit.SetWordBreakProc (BYVAL pfn AS LONG_PTR)
   SendMessageW m_hCtl, EM_SETWORDBREAKPROC, 0, pfn
END SUB
' ========================================================================================

' ========================================================================================
' Displays a balloon tip associated with an edit control.
' ========================================================================================
PRIVATE FUNCTION CEdit.ShowBalloonTip (BYREF peditballoontip AS EDITBALLOONTIP) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_SHOWBALLOONTIP, 0, CAST(LPARAM, @peditballoontip))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Undoes the last edit control operation in the control's undo queue.
' ========================================================================================
PRIVATE FUNCTION CEdit.Undo () AS BOOLEAN
   RETURN SendMessageW(m_hCtl, EM_UNDO, 0, 0)
END FUNCTION
' ========================================================================================

END NAMESPACE
