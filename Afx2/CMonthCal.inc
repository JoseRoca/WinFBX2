' ########################################################################################
' Platform: Microsoft Windows
' Filename: CMonthCal.inc
' Purpose:  Windows common controls
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
using Afx2

NAMESPACE Afx2

' ########################################################################################
' CMonthCal class
' ########################################################################################

TYPE CMonthCal EXTENDS CControl

Private:
   m_hCtl AS HWND   ' // Control handle

Public:
   DECLARE CONSTRUCTOR (BYVAL hCtl AS HWND)
   DECLARE CONSTRUCTOR (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   DECLARE CONSTRUCTOR (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   DECLARE DESTRUCTOR

   DECLARE FUNCTION GetCalendarBorder () AS DWORD
   DECLARE FUNCTION GetCalendarCount () AS DWORD
   DECLARE FUNCTION GetCalendarGridInfo (BYREF _mcGridInfo AS MCGRIDINFO) AS BOOLEan
   DECLARE FUNCTION GetCALID () AS CALID
   DECLARE FUNCTION GetColor (BYVAL iColor AS LONG) AS COLORREF
   DECLARE FUNCTION GetCurrentView () AS DWORD
   DECLARE FUNCTION GetCurSel (BYREF st AS SYSTEMTIME) AS BOOLEAN
   DECLARE FUNCTION GetFirstDayOfWeek () AS DWORD
   DECLARE FUNCTION GetMaxSelCount () AS DWORD
   DECLARE FUNCTION GetMaxTodayWidth () AS DWORD
   DECLARE FUNCTION GetMinReqRect (BYREF rectInfo AS RECT) AS LONG
   DECLARE FUNCTION GetMonthDelta () AS LONG
   DECLARE FUNCTION GetMonthRange (BYVAL gmr AS DWORD, BYVaL rgst AS SYSTEMTIME PTR) AS LONG
   DECLARE FUNCTION GetRange (BYVAL rgst AS SYSTEMTIME PTR) AS DWORD
   DECLARE FUNCTION GetSelRange (BYVAL rgst AS SYSTEMTIME PTR) AS LONG
   DECLARE FUNCTION GetToday (BYREF st AS SYSTEMTIME) AS BOOLEAN
   DECLARE FUNCTION HitTest (BYREF info AS MCHITTESTINFO) AS DWORD
   DECLARE FUNCTION SetCalendarBorder (BYVAL fset AS BOOLEAN, BYVAL xyborder AS LONG) AS LONG
   DECLARE SUB SetCALID (BYVAL calid AS UINT)
   DECLARE FUNCTION SetColor (BYVAL iColor AS LONG, BYVAL clr AS COLORREF) AS COLORREF
   DECLARE FUNCTION SetCurrentView (BYVAL dwNewView AS DWORD) AS BOOLEAN
   DECLARE FUNCTION SetCurSel (BYREF st AS SYSTEMTIME) AS LONG
   DECLARE FUNCTION SetDayState (BYVAL iMonths AS LONG, BYVAL rgds AS MONTHDAYSTATE PTR) AS BOOLEAN
   DECLARE FUNCTION SetFirstDayOfWeek (BYVAL iDay AS LONG) AS DWORD
   DECLARE FUNCTION SetMaxSelCount (BYVAL iMax AS UINT) AS BOOLEAN
   DECLARE FUNCTION SetMonthDelta (BYVAL iDelta AS LONG) AS LONG
   DECLARE FUNCTION SetRange (BYVAL fWhichLimit AS DWORD, BYVAL rgst AS SYSTEMTIME PTR) AS BOOLEAN
   DECLARE FUNCTION SetSelRange (BYVAL rgst AS SYSTEMTIME PTR) AS BOOLEAN
   DECLARE SUB SetToday (BYREF st AS SYSTEMTIME)
   DECLARE FUNCTION SizeRectToMin (BYREF rc AS RECT) AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' * CMonthCal constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CMonthCal (BYVAL hCtl AS HWND)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMonthCal (BYVAL hParent AS HWND, BYVAL cID AS LONG)
   DIM hCtl AS HWND = GetDlgItem(hParent, cID)
   IF IsChild(hParent, hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   DIM pDlg AS CDialog PTR = this.CDialogOwnerPtr(hCtl)
   IF pDlg THEN this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMonthCal (BYREF pDlg AS CDialog, BYVAL cID AS LONG)
   IF VARPTR(pDlg) = 0 THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg.hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg.UsesPixels
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMonthCal (BYVAL pDlg AS CDialog PTR, BYVAL cID AS LONG)
   IF pDlg = NULL THEN EXIT CONSTRUCTOR
   DIM hCtl AS HWND = GetDlgItem(pDlg->hDialog, cID)
   IF IsWindow(hCtl) THEN m_hCtl = hCtl ELSE EXIT CONSTRUCTOR
   this.UsesPixels = pDlg->UsesPixels
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CMonthCal destructor
' ========================================================================================
PRIVATE DESTRUCTOR CMonthCal
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Gets the border size, in pixels, of a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetCalendarBorder () AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETCALENDARBORDER, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of calendars currently displayed in the calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetCalendarCount () AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETCALENDARCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about a calendar grid.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetCalendarGridInfo (BYREF info AS MCGRIDINFO) AS BOOLEAN
   IF VARPTR(info) = 0 THEN info.cbSize = SIZEOF(info)
   RETURN SendMessageW(m_hCtl, MCM_GETCALENDARGRIDINFO, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current calendar ID for the given calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetCALID () AS CALID
   RETURN SendMessageW(m_hCtl, MCM_GETCALID, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the color for a given portion of a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetColor (BYVAL iColor AS LONG) AS COLORREF
   RETURN SendMessageW(m_hCtl, MCM_GETCOLOR, iColor, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the view for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetCurrentView () AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETCURRENTVIEW, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the currently selected date.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetCurSel (BYREF st AS SYSTEMTIME) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, MCM_GETCURSEL, 0, CAsT(LPARaM, @st))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the first day of the week for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetFirstDayOfWeek () AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETFIRSTDAYOFWEEK, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the maximum date range that can be selected in a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetMaxSelCount () AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETMAXSELCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the maximum width of the "today" string in a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetMaxTodayWidth () AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETMAXTODAYWIDTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the maximum width of the "today" string in a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetMinReqRect (BYREF rectInfo AS RECT) AS LONG
   RETURN SendMessageW(m_hCtl, MCM_GETMINREQRECT, 0, CasT(LPARAM, @rectInfo))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the scroll rate for a month calendar control. The scroll rate is the number of
' months that the control moves its display when the user clicks a scroll button.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetMonthDelta () AS LONG
   RETURN SendMessageW(m_hCtl, MCM_GETMONTHDELTA, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves date information (using SYSTEMTIME structures) that represents the high and
' low limits of a month calendar control's display.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetMonthRange (BYVAL gmr AS DWORD, BYVAL rgst AS SYSTEMTIME PTR) AS LONG
   RETURN SendMessageW(m_hCtl, MCM_GETMONTHRANGE, gmr, CAST(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the minimum and maximum allowable dates set for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetRange (BYVAL rgst AS SYSTEMTIME PTR) AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_GETRANGE, 0, CAST(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves date information that represents the upper and lower limits of the date range
' currently selected by the user.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetSelRange (BYVAL rgst AS SYSTEMTIME PTR) AS LONG
   RETURN SendMessageW(m_hCtl, MCM_GETSELRANGE, 0, CAST(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the date information for the date specified as "today" for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.GetToday (BYREF st AS SYSTEMTIME) AS BOOLEAN
   IF VARPTR(st) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, MCM_GETTODAY, 0, CAST(LPARAM, @st))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines which portion of a month calendar control is at a given point on the screen.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.HitTest (BYREF info AS MCHITTESTINFO) AS DWORD
   IF VARPTR(info) = 0 THEN info.cbSize = SIZEOF(info)
   RETURN SendMessageW(m_hCtl, MCM_HITTEST, 0, CAST(LPARAM, @info))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the border size, in pixels, of a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetCalendarBorder (BYVAL fset AS BOOLEAN, BYVAL xyborder AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, MCM_SETCALENDARBORDER, fset, xyborder)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the calendar ID for the given calendar control.
' ========================================================================================
PRIVATE SUB CMonthCal.SetCALID (BYVAL calid AS UINT)
   SendMessageW(m_hCtl, MCM_SETCALID, calid, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the color for a given portion of a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetColor (BYVAL iColor AS LONG, BYVAL clr AS COLORREF) AS COLORREF
   RETURN SendMessageW(m_hCtl, MCM_SETCOLOR, iColor, clr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the view for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetCurrentView (BYVAL dwNewView AS DWORD) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, MCM_SETCURRENTVIEW, 0, dwNewView)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the currently selected date for a month calendar control. If the specified date is
' not in view, the control updates the display to bring it into view.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetCurSel (BYREF st AS SYSTEMTIME) AS LONG
   IF VARPTR(st) = 0 THEN RETURN FALSE
   RETURN SendMessageW(m_hCtl, MCM_SETCURSEL, 0, CAST(LPARAM, @st))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the day states for all months that are currently visible within a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetDayState (BYVAL iMonths AS LONG, BYVAL rgds AS MONTHDAYSTATE PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, MCM_SETDAYSTATE, iMonths, CAST(LPARAM, rgds))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the first day of the week for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetFirstDayOfWeek (BYVAL iDay AS LONG) AS DWORD
   RETURN SendMessageW(m_hCtl, MCM_SETFIRSTDAYOFWEEK, 0, iDay)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the maximum number of days that can be selected in a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetMaxSelCount (BYVAL iMax AS UINT) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, MCM_SETMAXSELCOUNT, iMax, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the scroll rate for a month calendar control. The scroll rate is the number of
' months that the control moves its display when the user clicks a scroll button.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetMonthDelta (BYVAL iDelta AS LONG) AS LONG
   RETURN SendMessageW(m_hCtl, MCM_SETMONTHDELTA, iDelta, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum and maximum allowable dates for a month calendar control.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetRange (BYVAL fWhichLimit AS DWORD, BYVAL rgst AS SYSTEMTIME PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, MCM_SETRANGE, fWhichLimit, CAST(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the selection for a month calendar control to a given date range.
' ========================================================================================
PRIVATE FUNCTION CMonthCal.SetSelRange (BYVAL rgst AS SYSTEMTIME PTR) AS BOOLEAN
   RETURN SendMessageW(m_hCtl, MCM_SETSELRANGE, 0, CAST(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the "today" selection for a month calendar control.
' ========================================================================================
PRIVATE SUB CMonthCal.SetToday (BYREF st AS SYSTEMTIME)
   IF VARPTR(st) = 0 THEN EXIT SUB
   SendMessage m_hCtl, MCM_SETTODAY, 0, CAST(LPARAM, @st)
END SUB
' ========================================================================================

' ========================================================================================
' Calculates how many calendars will fit in the given rectangle, and then returns the
' minimum size that a rectangle needs to be to fit that number of calendars.
' ========================================================================================
FUNCTION CMonthCal.SizeRectToMin (BYREF rc AS RECT) AS LONG
   IF VARPTR(rc) = 0 THEN RETURN 0
   RETURN SendMessageW(m_hCtl, MCM_SIZERECTTOMIN, 0, CAST(LPARAM, @rc))
END FUNCTION
' ========================================================================================

END NAMESPACE
