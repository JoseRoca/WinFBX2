' ########################################################################################
' Platform: Microsoft Windows
' Filename: CStatic.inc
' Purpose:  Windows common controls - Static controls
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

#pragma once
#include once "windows.bi"
#include once "win/commctrl.bi"
#include once "Afx2/CControl.inc"
using Afx2
' // conflicting define in windowsx.bi
#undef DeleteBitmap

NAMESPACE Afx2

' ########################################################################################
' CStatic class
' ########################################################################################

TYPE CStatic EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION Enable () AS BOOLEAN
   DECLARE FUNCTION Disable () AS BOOLEAN
   DECLARE FUNCTION DeleteBitmap () AS BOOLEAN
   DECLARE FUNCTION DeleteCursor () AS BOOLEAN
   DECLARE FUNCTION DeleteEnhancedMetafile () AS BOOLEAN
   DECLARE FUNCTION DeleteIcon () AS BOOLEAN
   DECLARE FUNCTION DeleteImage (BYVAL nType AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetIcon () AS HICON
   DECLARE FUNCTION GetImage (BYVAL nType AS LONG) AS HANDLE
   DECLARE FUNCTION SetImage (BYVAL nType AS LONG, BYVAL hImage AS HANDLE) AS HANDLE
   DECLARE FUNCTION SetBitmap (BYVAL hBmp AS HBITMAP) AS HBITMAP
   DECLARE FUNCTION SetCursor (BYVAL hCur AS HCURSOR) AS HCURSOR
   DECLARE FUNCTION SetEnhancedMetafile (BYVAL _hEnhMetafile AS HENHMETAFILE) AS HENHMETAFILE
   DECLARE FUNCTION SetIcon (BYVAL hIcon AS HICON) AS HICON
   DECLARE FUNCTION GetTextLength () AS LONG
   DECLARE FUNCTION GetText () AS DWSTRING
   DECLARE FUNCTION SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' * CStatic constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CStatic (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CStatic (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CStatic (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CStatic (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CStatic destructor
' ========================================================================================
PRIVATE DESTRUCTOR CStatic
END DESTRUCTOR
' ========================================================================================


' ########################################################################################
'                                     *** STATIC ***
' ########################################################################################

' ========================================================================================
' * Enables the control.
' Return value: False if the control was previous enabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CStatic.Enable () AS BOOLEAN
   RETURN EnableWindow(m_hCtl, TRUE)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Disables the control.
' Return value: False if the control was previous disabled; otherwise TRUE
' ========================================================================================
PRIVATE FUNCTION CStatic.Disable () AS BOOLEAN
   DIM res AS LONG = EnableWindow(m_hCtl, FALSE)
   DIM bEnable AS BOOLEAN
   IF res = 0 THEN bEnable = TRUE ELSE bEnable = FALSE
   RETURN bEnable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a bitmap associated with a static control. Returns TRUE o FALSE.
' ========================================================================================
PRIVATE FUNCTION CStatic.DeleteBitmap () AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, IMAGE_BITMAP), 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a cursor associated with a static control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CStatic.DeleteCursor () AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, IMAGE_CURSOR), 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a cursor associated with a static control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CStatic.DeleteEnhancedMetafile () AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, IMAGE_ENHMETAFILE), 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes an icon associated with a static control. Returns TRUE o FALSE.
' ========================================================================================
PRIVATE FUNCTION CStatic.DeleteIcon () AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, STM_SETICON, 0, 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes an image associated with a static control. Returns TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION CStatic.DeleteImage (BYVAL nType AS LONG) AS BOOLEAN
   RETURN DeleteObject(cast(HGDIOBJ, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, nType), 0)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a handle to the icon associated with a static control that has the SS_ICON style.
' ========================================================================================
PRIVATE FUNCTION CStatic.GetIcon () AS HICON
   RETURN CAST(HICON, SendMessageW(m_hCtl, STM_GETICON, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a handle to the image (icon or bitmap) associated with a static control.
' The return value is a handle to the image, if any; otherwise, it is NULL.
' ========================================================================================
PRIVATE FUNCTION CStatic.GetImage (BYVAL nType AS LONG) AS HANDLE
   RETURN cast(HANDLE, SendMessageW(m_hCtl, STM_GETIMAGE, nType, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates a bitmap with a static control.
' The return value is a handle to the image previously associated with the static control,
' if any; otherwise, it is NULL.
' ========================================================================================
PRIVATE FUNCTION CStatic.SetBitmap (BYVAL hBmp AS HBITMAP) AS HBITMAP
   RETURN cast(HBITMAP, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, IMAGE_BITMAP), cast(LPARAM, hBmp)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates a cursor with a static control.
' The return value is a handle to the cursor previously associated with the static control,
' ========================================================================================
PRIVATE FUNCTION CStatic.SetCursor (BYVAL hCur AS HCURSOR) AS HCURSOR
   RETURN cast(HCURSOR, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, IMAGE_CURSOR), cast(LPARAM, hCur)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates an enhanced metafile with a static control.
' The return value is a handle to the image previously associated with the static control,
' if any; otherwise, it is NULL.
' ========================================================================================
PRIVATE FUNCTION CStatic.SetEnhancedMetafile (BYVAL _hEnhMetafile AS HENHMETAFILE) AS HENHMETAFILE
   RETURN cast(HENHMETAFILE, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, IMAGE_ENHMETAFILE), cast(LPARAM, _hEnhMetafile)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates an icon with an static control.
' ========================================================================================
PRIVATE FUNCTION CStatic.SetIcon (BYVAL hIcon AS HICON) AS HICON
   RETURN CAST(..HICON, SendMessageW(m_hCtl, STM_SETICON, CAST(WPARAM, hIcon), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Associates a new image with a static control.
' The return value is a handle to the image previously associated with the static control,
' if any; otherwise, it is NULL.
' To associate an image with a static control, the control must have the proper style. The
' following table shows the style needed for each image type.
' Image type          Static control style
'   IMAGE_BITMAP      SS_BITMAP
'   IMAGE_CURSOR      SS_ICON
'   IMAGE_ENHMETAFILE SS_ENHMETAFILE
'   IMAGE_ICON        SS_ICON
' ========================================================================================
PRIVATE FUNCTION CStatic.SetImage (BYVAL nType AS LONG, BYVAL hImage AS HANDLE) AS HANDLE
   RETURN cast(HANDLE, SendMessageW(m_hCtl, STM_SETIMAGE, cast(WPARAM, nType), cast(LPARAM, hImage)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the length of the text of a static control.
' ========================================================================================
PRIVATE FUNCTION CStatic.GetTextLength () AS LONG
   RETURN GetWindowTextLength(m_hCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of a static control.
' ========================================================================================
PRIVATE FUNCTION CStatic.GetText () AS DWSTRING
   DIM nLen AS LONG, buffer AS DWSTRING
   nLen = GetWindowTextLength(m_hCtl)
   buffer = WSPACE(nLen + 1)
   GetWindowTextW m_hCtl, STRPTR(buffer), nLen + 1
   RETURN LEFT(buffer, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of a static control.
' ========================================================================================
PRIVATE FUNCTION CStatic.SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
   RETURN SetWindowTextW(m_hCtl, pwszText)
END FUNCTION
' ========================================================================================

END NAMESPACE
