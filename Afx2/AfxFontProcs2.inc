' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxFontProcs2.inc
' Purpose:  Windows font procedures
' Compiler: FreeBASIC 32 & 64 bit
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
#include once "Afx2/AfxWin2.inc"
USING Afx2

NAMESPACE Afx2

' ########################################################################################
'                                      *** FONTS ***
' ########################################################################################

' ========================================================================================
' Font enum
' ========================================================================================
enum
   AFX_FONT_CAPTION = 1    ' // Caption font
   AFX_FONT_SMALLCAPTION   ' // Small caption font
   AFX_FONT_MENU           ' // Font used in menu bars
   AFX_FONT_STATUS         ' // Font used in status bars and tooltips
   AFX_FONT_MESSAGE        ' // Font used in message boxes
   ' // Font settings
   AFX_FONT_HEIGHT         ' // Font height
   AFX_FONT_WEIGHT         ' // Font weight
   AFX_FONT_ITALIC         ' // Font italic
   AFX_FONT_UNDERLINE      ' // Font underline
   AFX_FONT_STRIKEOUT      ' // Font strikeout
   AFX_FONT_CHARSET        ' // Font charset
end enum
' ========================================================================================

' ========================================================================================
' Gets the font with which the window is currently drawing its text.
' Note: You can also use the GetWindowFont macro (windowsx.bi).
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowFont (BYVAL hwnd AS HWND) AS HFONT
   FUNCTION = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font with which the window is currently drawing its text.
' hwnd
'   Handle to the window.
' hFont
'   Handle to the font. If this parameter is NULL, the control uses the default system
'   font to draw text.
' fRedraw
'   Specifies whether the control should be redrawn immediately upon setting the font. If
'   this parameter is TRUE, the control redraws itself.
' ========================================================================================
PRIVATE SUB AfxSetWindowFont (BYVAL hwnd AS HWND, BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = CTRUE)
   SendMessageW(hwnd, WM_SETFONT, CAST(WPARAM, hFont), fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves information about the font being used by a window or control
' Parameter:
' - hwnd = Handle of the window or control.
' Return value:
' - A LOGFONTW structure.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowFontInfo (BYVAL hwnd AS HWND) AS LOGFONTW
   ' // Get the handle of the font used by the window or control
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
   ' // Get the LOGFONTW structure
   DIM lfw AS LOGFONTW
   IF hCurFont THEN GetObjectW(hCurFont, SIZEOF(lfw), @lfw)
   RETURN lfw
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information about the fonts used by Windows.
' Parameters:
' - nType = The type of font:
'   AFX_FONT_CAPTION, AFX_FONT_SMALLCAPTION, AFX_FONT_MENU, AFX_FONT_STATUS, AFX_FONT_MESSAGE
' - plfw = Pointer to a LOGFONTW structure that receives the font information.
' Return value: TRUE on succes or FALSE on failure.
' To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowsFontInfo (BYVAL nType AS LONG, BYVAL plfw AS LOGFONTW PTR) AS BOOLEAN
   DIM ncm AS NONCLIENTMETRICSW
   IF plfw = NULL THEN EXIT FUNCTION
   IF AfxWindowsVersion >= 6 THEN ncm.cbSize = 504 ELSE ncm.cbSize = 500
   IF SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, SIZEOF(ncm), @ncm, 0) = 0 THEN EXIT FUNCTION
   SELECT CASE nType
      CASE AFX_FONT_CAPTION      : *plfw = ncm.lfCaptionFont
      CASE AFX_FONT_SMALLCAPTION : *plfw = ncm.lfSmCaptionFont
      CASE AFX_FONT_MENU         : *plfw = ncm.lfMenuFont
      CASE AFX_FONT_STATUS       : *plfw = ncm.lfStatusFont
      CASE AFX_FONT_MESSAGE      : *plfw = ncm.lfMessageFont
      CASE ELSE
         RETURN FALSE
   END SELECT
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the point size of the fonts used by Windows.
' Parameters:
' - nType = The type of font:
'   AFX_FONT_CAPTION, AFX_FONT_SMALLCAPTION, AFX_FONT_MENU, AFX_FONT_STATUS, AFX_FONT_MESSAGE
' Return value: The point size.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowsFontPointSize (BYVAL nType AS LONG) AS LONG
   DIM ncm AS NONCLIENTMETRICSW
   IF AfxWindowsVersion >= 6 THEN ncm.cbSize = 504 ELSE ncm.cbSize = 500
   IF SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, SIZEOF(ncm), @ncm, 0) = 0 THEN EXIT FUNCTION
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN EXIT FUNCTION
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC hDC
   DIM nPointSize AS LONG
   SELECT CASE nType
      CASE AFX_FONT_CAPTION      : nPointSize = MulDiv(ncm.lfCaptionFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_SMALLCAPTION : nPointSize = MulDiv(ncm.lfSmCaptionFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_MENU         : nPointSize = MulDiv(ncm.lfMenuFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_STATUS       : nPointSize = MulDiv(ncm.lfStatusFont.lfHeight, 72, cyPixelsPerInch)
      CASE AFX_FONT_MESSAGE      : nPointSize = MulDiv(ncm.lfMessageFont.lfHeight, 72, cyPixelsPerInch)
   END SELECT
   IF nPointSize < 0 THEN nPointSize = -nPointSize
   FUNCTION = nPointSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the point size of a font given its logical height
' ========================================================================================
PRIVATE FUNCTION AfxGetFontPointSize (BYVAL nHeight AS LONG) AS LONG
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN EXIT FUNCTION
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC HDC
   DIM nPointSize AS LONG = MulDiv(nHeight, 72, cyPixelsPerInch)
   IF nPointSize < 0 THEN nPointSize = -nPointSize
   FUNCTION = nPointSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the logical height of a font given its point size
' ========================================================================================
PRIVATE FUNCTION AfxGetFontHeight (BYVAL nPointSize AS LONG) AS LONG
   DIM hDC AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
   IF hDC = NULL THEN EXIT FUNCTION
   DIM cyPixelsPerInch AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   DeleteDC HDC
   DIM nHeight AS LONG = -MulDiv(nPointSize, cyPixelsPerInch, 72)
   FUNCTION = nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies the face name of the font of a window or control.
' Parameters:
' - hwnd = Handle of the window or control.
' - wszNewFaceName = The new face name of the font.
' Return value: The handle of the new font on success, or NULL on failure.
' To get extended error information, call GetLastError.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION AfxModifyFontFaceName (BYVAL hwnd AS HWND, BYREF wszNewFaceName AS WSTRING) AS HFONT
   DIM lfw AS LOGFONTW
   IF hwnd = NULL OR VARPTR(wszNewFaceName) = NULL THEN EXIT FUNCTION
   ' // Get the handle of the font used by the window
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
   IF hCurFont = 0 THEN EXIT FUNCTION
   ' // Get the LOGFONTW structure
   IF GetObject(hCurFont, SIZEOF(lfw), @lfw) = 0 THEN EXIT FUNCTION
   ' // Change the face name
   lfw.lfFaceName = wszNewFaceName
   ' // Create a new font
   DIM hNewFont AS HFONT = CreateFontIndirectW(@lfw)
   IF hNewFont = 0 THEN EXIT FUNCTION
   ' // Select the new font and delete the old one
   DIM hDC AS HDC = GetDC(hwnd)
   DeleteObject(SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
   ReleaseDC(hwnd, hDC)
   SendMessageW(hwnd, WM_SETFONT, CAST(WPARAM, hNewFont), CTRUE)
   FUNCTION = hNewFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies settings of the font used by a window of control.
' Parameters:
' - hwnd = Handle of the window or control.
' - nSetting : One of the AFX_FONT_xxx constants.
' - nValue: Depends of the nSetting value
'   AFX_FONT_HEIGHT
'      The base is 100. To increase the font a 20% pass 120; to reduce it a 20% pass 80%.
'   AFX_FONT_WEIGHT
'      The weight of the font in the range 0 through 1000. For example, 400 is normal and
'      700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
'   AFX_FONT_ITALIC : TRUE or FALSE.
'   AFX_FONT_UNDERLINE : TRUE or FALSE.
'   AFX_FONT_STRIKEOUT : TRUE or FALSE.
'   AFX_FONT_CHARSET
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the new font on success, or NULL on failure.
' To get extended error information, call GetLastError.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION AfxModifyFontSettings (BYVAL hwnd AS HWND, BYVAL nSetting AS LONG, BYVAL nValue AS LONG) AS HFONT
   DIM lfw AS LOGFONTW
   IF IsWindow(hwnd) = 0 THEN EXIT FUNCTION
   ' // Get the handle of the font used by the window
   DIM hCurFont AS HFONT = CAST(HFONT, SendMessageW(hwnd, WM_GETFONT, 0, 0))
   IF hCurFont = NULL THEN EXIT FUNCTION
   ' // Get the LOGFONTW structure
   IF GetObject(hCurFont, SIZEOF(lfw), @lfw) = 0 THEN EXIT FUNCTION
   ' // Change the specified setting
   SELECT CASE nSetting
      CASE AFX_FONT_HEIGHT
         ' // Change the point size
         DIM lPointSize AS LONG = AfxGetFontPointSize(lfw.lfHeight)
         lPointSize = lPointSize * (nValue / 100)
         lfw.lfHeight = -MulDiv(lPointSize, AfxLogPixelsY, 72)
      CASE AFX_FONT_WEIGHT
         ' // Change the font weight
         lfw.lfWeight = nValue
      CASE AFX_FONT_ITALIC
         ' // Change the italic flag
         lfw.lfItalic = CUBYTE(nValue)
      CASE AFX_FONT_UNDERLINE
         ' // Change the underline flag
         lfw.lfUnderline = CUBYTE(nValue)
      CASE AFX_FONT_STRIKEOUT
         ' // Change the strikeout flag
         lfw.lfStrikeOut = CUBYTE(nValue)
      CASE AFX_FONT_CHARSET
         ' // Change the charset
         lfw.lfCharset = CUBYTE(nValue)
      CASE ELSE
         EXIT FUNCTION
   END SELECT
   ' // Create a new font
   DIM hNewFont AS HFONT = CreateFontIndirectW(@lfw)
   IF hNewFont = NULL THEN EXIT FUNCTION
   ' // Select the new font and delete the old one
   DIM hDC AS HDC = GetDC(hwnd)
   DeleteObject(SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
   ReleaseDC(hwnd, hDC)
   SendMessageW(hwnd, WM_SETFONT, CAST(WPARAM, hNewFont), CTRUE)
   FUNCTION = hNewFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Modifies the height of the font used by a window of control.
' Parameters:
' - hwnd = Handle of the window or control.
' - nValue: The base is 100. To increase the font a 20% pass 120; to reduce it a 20% pass 80%.
' Return value: The handle of the new font on success, or NULL on failure.
' To get extended error information, call GetLastError.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' ========================================================================================
PRIVATE FUNCTION AfxModifyFontHeight (BYVAL hwnd AS HWND, BYVAL nValue AS LONG) AS HFONT
   FUNCTION = AfxModifyFontSettings(hwnd, AFX_FONT_HEIGHT, nValue)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a logical font.
' Parameters:
' - wszFaceName = The typeface name.
' - lPointSize = The point size.
' - DPI = Dots per inch to calculate scaling. Default value = 96 (no scaling). If you pass -1
'   and the application is DPI aware, the DPI value used by the operating system will be used.
' - lWeight = The weight of the font in the range 0 through 1000. For example, 400 is normal
'      and 700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the font or NULL on failure.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' Usage examples:
'   hFont = AfxCreateFont("MS Sans Serif", 8, , FW_NORMAL, , , , DEFAULT_CHARSET)
'   hFont = AfxCreateFont("Courier New", 10, -1 , FW_BOLD, , , , DEFAULT_CHARSET)
'   hFont = AfxCreateFont("Marlett", 8, -1, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION AfxCreateFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL DPI AS LONG = 96, _
   BYVAL lWeight AS LONG = 0, BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, _
   BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT

   DIM tlfw AS LOGFONTW
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)

   ' // Font scaling
   IF DPI = -1 THEN DPI = GetDeviceCaps(hDC, LOGPIXELSX)
   IF DPI > 0 THEN lPointSize = (lPointSize * DPI) \ GetDeviceCaps(hDC, LOGPIXELSY)

   tlfw.lfHeight         = -MulDiv(lPointSize, .GetDeviceCaps(hDC, LOGPIXELSY), 72)  ' logical font height
   tlfw.lfWidth          =  0                                                        ' average character width
   tlfw.lfEscapement     =  0                                                        ' escapement
   tlfw.lfOrientation    =  0                                                        ' orientation angles
   tlfw.lfWeight         =  lWeight                                                  ' font weight
   tlfw.lfItalic         =  bItalic                                                  ' italic(CTRUE/FALSE)
   tlfw.lfUnderline      =  bUnderline                                               ' underline(CTRUE/FALSE)
   tlfw.lfStrikeOut      =  bStrikeOut                                               ' strikeout(CTRUE/FALSE)
   tlfw.lfCharSet        =  bCharset                                                 ' character set
   tlfw.lfOutPrecision   =  OUT_TT_PRECIS                                            ' output precision
   tlfw.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                      ' clipping precision
   tlfw.lfQuality        =  DEFAULT_QUALITY                                          ' output quality
   tlfw.lfPitchAndFamily =  FF_DONTCARE                                              ' pitch and family
   tlfw.lfFaceName       =  wszFaceName                                              ' typeface name

   ReleaseDC(HWND_DESKTOP, hDC)
   FUNCTION = CreateFontIndirectW(@tlfw)

END FUNCTION
' ========================================================================================

END NAMESPACE
