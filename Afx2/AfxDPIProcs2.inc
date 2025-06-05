' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxDPIProcs2.inc
' Purpose:  High DPI procedures
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

NAMESPACE Afx2

' // Identifies the dots per inch (dpi) setting for a thread, process, or window.
enum DPI_AWARENESS
  DPI_AWARENESS_INVALID = -1
  DPI_AWARENESS_UNAWARE = 0
  DPI_AWARENESS_SYSTEM_AWARE = 1
  DPI_AWARENESS_PER_MONITOR_AWARE = 2
end enum

' // Identifies the awareness context for a window.
#define DPI_AWARENESS_CONTEXT HANDLE
#define DPI_AWARENESS_CONTEXT_UNAWARE              (CAST(DPI_AWARENESS_CONTEXT,-1))
#define DPI_AWARENESS_CONTEXT_SYSTEM_AWARE         (CAST(DPI_AWARENESS_CONTEXT,-2))
#define DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE    (CAST(DPI_AWARENESS_CONTEXT,-3))
#define DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 (CAST(DPI_AWARENESS_CONTEXT,-4))
#define DPI_AWARENESS_CONTEXT_UNAWARE_GDISCALED    (CAST(DPI_AWARENESS_CONTEXT,-5))

' // Identifies the DPI hosting behavior for a window.
' // This behavior allows windows created in the thread to host
' // child windows with a different DPI_AWARENESS_CONTEXT
enum DPI_HOSTING_BEHAVIOR
  DPI_HOSTING_BEHAVIOR_INVALID = -1
  DPI_HOSTING_BEHAVIOR_DEFAULT = 0
  DPI_HOSTING_BEHAVIOR_MIXED = 1
end enum

' // Describes per-monitor DPI scaling behavior overrides for child windows within dialogs.
' // The values in this enumeration are bitfields and can be combined.
enum DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS
  DCDC_DEFAULT = &h0000
  DCDC_DISABLE_FONT_UPDATE = &h0001
  DCDC_DISABLE_RELAYOUT = &h0002
end enum

' // In Per Monitor v2 contexts, dialogs will automatically respond to DPI changes by
' // resizing themselves and re-computing the positions of their child windows
' // (here referred to as re-layouting). This enum works in conjunction with
' // SetDialogDpiChangeBehavior in order to override the default DPI scaling behavior for dialogs.
' // This does not affect DPI scaling behavior for the child windows of dialogs
' // (beyond re-layouting), which is controlled by DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS.
enum DIALOG_DPI_CHANGE_BEHAVIORS
  DDC_DEFAULT = &h0000
  DDC_DISABLE_ALL = &h0001
  DDC_DISABLE_RESIZE = &h0002
  DDC_DISABLE_CONTROL_RELAYOUT = &h0004
end enum

'extern "Windows-MS"
'DECLARE FUNCTION AdjustWindowRectExForDpi (BYVAL lpRect AS RECT PTR, BYVAL dwStyle AS DWORD, BYVAL bMenu As WINBOOL, _
'   BYVAL dwExStyle AS DWORD, BYVAL dpi AS UINT) AS WINBOOL
'DECLARE FUNCTION AreDpiAwarenessContextsEqual (BYVAL dpiContextA AS DPI_AWARENESS_CONTEXT, BYVAL dpiContextB AS DPI_AWARENESS_CONTEXT) AS WINBOOL
'DECLARE FUNCTION EnableNonClientDpiScaling (BYVAL hwnd AS HWND) AS WINBOOL
'DECLARE FUNCTION GetAwarenessFromDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS DPI_AWARENESS
'DECLARE FUNCTION GetDialogControlDpiChangeBehavior (BYVAL hwnd AS HWND) AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS
'DECLARE FUNCTION GetDialogDpiChangeBehavior (BYVAL hDlg AS HWND) AS DIALOG_DPI_CHANGE_BEHAVIORS
'DECLARE FUNCTION GetDpiAwarenessContextForProcess (BYVAL hProcess AS HANDLE) AS DPI_AWARENESS_CONTEXT
'DECLARE FUNCTION GetDpiForSystem() AS UINT
'DECLARE FUNCTION GetDpiForWindow (BYVAL hwnd AS HWND) AS UINT
'DECLARE FUNCTION GetDpiFromDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS UINT
'DECLARE FUNCTION GetSystemDpiForProcess (BYVAL hProcess AS HANDLE) AS UINT
'DECLARE FUNCTION GetSystemMetricsForDpi (BYVAL nIndex AS int_, BYVAL dpi AS UINT) AS int_
'DECLARE FUNCTION GetThreadDpiAwarenessContext () AS DPI_AWARENESS_CONTEXT
'DECLARE FUNCTION GetThreadDpiHostingBehavior () AS DPI_HOSTING_BEHAVIOR
'DECLARE FUNCTION GetWindowDpiAwarenessContext (BYVAL hwnd AS HWND) AS DPI_AWARENESS_CONTEXT
'DECLARE FUNCTION InheritWindowMonitor (BYVAL hwnd AS HWND, BYVAL hwndInherit AS HWND) AS WINBOOL
'DECLARE FUNCTION IsValidDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS WINBOOL
'DECLARE FUNCTION LogicalToPhysicalPointForPerMonitorDPI (BYVAL hwnd AS HWND, BYVAL lpPoint AS LPPOINT) AS WINBOOL
'DECLARE FUNCTION PhysicalToLogicalPointForPerMonitorDPI (BYVAL hwnd AS HWND, BYVAL lpPoint AS LPPOINT) AS WINBOOL
'DECLARE FUNCTION SetDialogControlDpiChangeBehavior (BYVAL hwnd AS HWND, BYVAL mask AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS, _
'   BYVAL values AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS) AS WINBOOL
'DECLARE FUNCTION SetProcessDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS WINBOOL
'DECLARE FUNCTION SetThreadCursorCreationScaling (BYVAL cursorDPI AS UINT) AS UINT
'DECLARE FUNCTION SetThreadDpiAwarenessContext (BYVAL dpiContext AS DPI_AWARENESS_CONTEXT) AS DPI_AWARENESS_CONTEXT
'DECLARE FUNCTION SetThreadDpiHostingBehavior (BYVAL value AS DPI_HOSTING_BEHAVIOR) AS DPI_HOSTING_BEHAVIOR
'DECLARE FUNCTION SystemParametersInfoForDpi (BYVAL uiAction AS UINT, BYVAL uiParam AS UINT, BYVAL pvParam AS PVOID, BYVAL fWinIni AS UINT, BYVAL dpi AS UINT) AS WINBOOL
'end extern

' ========================================================================================
' Calculates the required size of the window rectangle, based on the desired size of the
' client rectangle and the provided DPI. 
' ========================================================================================
PRIVATE FUNCTION AfxAdjustWindowRectExForDpi (BYREF rc AS RECT, BYVAL dwStyle AS DWORD, _
   BYVAL bMenu AS BOOLEAN, BYVAL dwExStyle AS DWORD, BYVAL dpi AS UINT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pAdjustWindowRectExForDpi AS FUNCTION (BYREF rc AS RECT, BYVAL dwStyle AS DWORD, _
      BYVAL bMenu AS BOOLEAN, BYVAL dwExStyle AS DWORD, BYVAL dpi AS UINT) AS BOOLEAN
   pAdjustWindowRectExForDpi = DyLibSymbol(pLib, "AdjustWindowRectExForDpi")
   IF pAdjustWindowRectExForDpi THEN FUNCTION = pAdjustWindowRectExForDpi(rc, dwStyle, bMenu, dwExStyle, dpi)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether two DPI_AWARENESS_CONTEXT values are identical.
' ========================================================================================
PRIVATE FUNCTION AfxAreDpiAwarenessContextsEqual (BYVAL dpiContextA AS DPI_AWARENESS_CONTEXT, _
   BYVAL dpiContextB AS DPI_AWARENESS_CONTEXT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pAreDpiAwarenessContextsEqual AS FUNCTION (BYVAL dpiContextA AS DPI_AWARENESS_CONTEXT, _
       BYVAL dpiContextB AS DPI_AWARENESS_CONTEXT) AS BOOLEAN
   pAreDpiAwarenessContextsEqual = DyLibSymbol(pLib, "AreDpiAwarenessContextsEqual")
   IF pAreDpiAwarenessContextsEqual THEN FUNCTION = pAreDpiAwarenessContextsEqual(dpiContextA, dpiContextB)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' In high-DPI displays, enables automatic display scaling of the non-client area portions
' of the specified top-level window. Must be called during the initialization of that window.
' ========================================================================================
PRIVATE FUNCTION AfxEnableNonClientDpiScaling (BYVAL hwnd AS HWND) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pEnableNonClientDpiScaling AS FUNCTION (BYVAL hwnd AS ..HWND) AS BOOLEAN
   pEnableNonClientDpiScaling = DyLibSymbol(pLib, "EnableNonClientDpiScaling")
   IF pEnableNonClientDpiScaling THEN FUNCTION = pEnableNonClientDpiScaling(hwnd)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the DPI_AWARENESS value from a DPI_AWARENESS_CONTEXT.
' ========================================================================================
PRIVATE FUNCTION AfxGetAwarenessFromDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS DPI_AWARENESS
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pGetAwarenessFromDpiAwarenessContext AS FUNCTION (BYVAL value AS DPI_AWARENESS_CONTEXT) AS DPI_AWARENESS
   pGetAwarenessFromDpiAwarenessContext = DyLibSymbol(pLib, "GetAwarenessFromDpiAwarenessContext")
   IF pGetAwarenessFromDpiAwarenessContext THEN FUNCTION = pGetAwarenessFromDpiAwarenessContext(value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves per-monitor DPI scaling behavior and overrides of a child window in a dialog.
' ========================================================================================
PRIVATE FUNCTION AfxGetDialogControlDpiChangeBehavior (BYVAL hwnd AS HWND) AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pGetDialogControlDpiChangeBehavior AS FUNCTION (BYVAL hwnd AS ..HWND) AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS
   pGetDialogControlDpiChangeBehavior = DyLibSymbol(pLib, "GetDialogControlDpiChangeBehavior")
   IF pGetDialogControlDpiChangeBehavior THEN FUNCTION = pGetDialogControlDpiChangeBehavior(hwnd)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the flags that might have been set on a given dialog by an earlier call to SetDialogDpiChangeBehavior.
' If that function was never called on the dialog, the return value will be zero.
' ========================================================================================
PRIVATE FUNCTION AfxGetDialogDpiChangeBehavior (BYVAL hDlg AS HWND) AS DIALOG_DPI_CHANGE_BEHAVIORS
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pGetDialogDpiChangeBehavior AS FUNCTION (BYVAL hDlg AS HWND) AS DIALOG_DPI_CHANGE_BEHAVIORS
   pGetDialogDpiChangeBehavior = DyLibSymbol(pLib, "GetDialogDpiChangeBehavior")
   IF pGetDialogDpiChangeBehavior THEN FUNCTION = pGetDialogDpiChangeBehavior(hDlg)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a DPI_AWARENESS_CONTEXT handle for the specified process.
' ========================================================================================
PRIVATE FUNCTION AfxGetDpiAwarenessContextForProcess (BYVAL hProcess AS HANDLE) AS DPI_AWARENESS_CONTEXT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pGetDpiAwarenessContextForProcess AS FUNCTION (BYVAL hProcess AS HANDLE) AS DPI_AWARENESS_CONTEXT
   pGetDpiAwarenessContextForProcess = DyLibSymbol(pLib, "GetDpiAwarenessContextForProcess")
   IF pGetDpiAwarenessContextForProcess THEN FUNCTION = pGetDpiAwarenessContextForProcess(hProcess)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the system DPI.
' ========================================================================================
PRIVATE FUNCTION AfxGetDpiForSystem () AS UINT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pGetDpiForSystem AS FUNCTION () AS UINT
   pGetDpiForSystem = DyLibSymbol(pLib, "GetDpiForSystem")
   IF pGetDpiForSystem THEN FUNCTION = pGetDpiForSystem()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the dots per inch (dpi) value for the specified window.
' ========================================================================================
PRIVATE FUNCTION AfxGetDpiForWindow (BYVAL hwnd AS HWND) AS UINT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pGetDpiForWindow AS FUNCTION (BYVAL hwnd AS ..HWND) AS UINT
   pGetDpiForWIndow = DyLibSymbol(pLib, "GetDpiForWindow")
   IF pGetDpiForWindow THEN FUNCTION = pGetDpiForWindow(hwnd)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the DPI from a given DPI_AWARENESS_CONTEXT handle. This enables you to determine
' the DPI of a thread without needed to examine a window created within that thread.
' ========================================================================================
PRIVATE FUNCTION AfxGetDpiFromDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS UINT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pGetDpiFromDpiAwarenessContext AS FUNCTION (BYVAL value AS DPI_AWARENESS_CONTEXT) AS UINT
   pGetDpiFromDpiAwarenessContext = DyLibSymbol(pLib, "GetDpiFromDpiAwarenessContext")
   IF pGetDpiFromDpiAwarenessContext THEN FUNCTION = pGetDpiFromDpiAwarenessContext(value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the system DPI associated with a given process.
' ========================================================================================
PRIVATE FUNCTION AfxGetSystemDpiForProcess (BYVAL hProcess AS HANDLE) AS UINT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pGetSystemDpiForProcess AS FUNCTION (BYVAL hProcess AS HANDLE) AS UINT
   pGetSystemDpiForProcess = DyLibSymbol(pLib, "GetSystemDpiForProcess")
   IF pGetSystemDpiForProcess THEN FUNCTION = pGetSystemDpiForProcess(hProcess)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the specified system metric or system configuration setting taking into account a provided DPI.
' ========================================================================================
PRIVATE FUNCTION AfxGetSystemMetricsForDpi (BYVAL nIndex AS LONG, BYVAL dpi AS UINT) AS LONG
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pGetSystemMetricsForDpi AS FUNCTION (BYVAL nIndex AS LONG, BYVAL dpi AS UINT) AS LONG
   pGetSystemMetricsForDpi = DyLibSymbol(pLib, "GetSystemMetricsForDpi")
   IF pGetSystemMetricsForDpi THEN FUNCTION = pGetSystemMetricsForDpi(nIndex, dpi)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the DPI_AWARENESS_CONTEXT for the current thread.
' ========================================================================================
PRIVATE FUNCTION AfxGetThreadDpiAwarenessContext () AS DPI_AWARENESS_CONTEXT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pGetThreadDpiAwarenessContext AS FUNCTION () AS DPI_AWARENESS_CONTEXT
   pGetThreadDpiAwarenessContext = DyLibSymbol(pLib, "GetThreadDpiAwarenessContext")
   IF pGetThreadDpiAwarenessContext THEN FUNCTION = pGetThreadDpiAwarenessContext()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the DPI_HOSTING_BEHAVIOR from the current thread.
' ========================================================================================
PRIVATE FUNCTION AfxGetThreadDpiHostingBehavior () AS DPI_HOSTING_BEHAVIOR
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pGetThreadDpiHostingBehavior AS FUNCTION () AS DPI_HOSTING_BEHAVIOR
   pGetThreadDpiHostingBehavior = DyLibSymbol(pLib, "GetThreadDpiHostingBehavior")
   IF pGetThreadDpiHostingBehavior THEN FUNCTION = pGetThreadDpiHostingBehavior()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the DPI_AWARENESS_CONTEXT associated with a window.
' ========================================================================================
PRIVATE FUNCTION AfxGetWindowDpiAwarenessContext (BYVAL hwnd AS HWND) AS DPI_AWARENESS_CONTEXT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pGetWindowDpiAwarenessContext AS FUNCTION (BYVAL hwnd AS ..HWND) AS DPI_AWARENESS_CONTEXT
   pGetWindowDpiAwarenessContext = DyLibSymbol(pLib, "GetWindowDpiAwarenessContext")
   IF pGetWindowDpiAwarenessContext THEN FUNCTION = pGetWindowDpiAwarenessContext(hwnd)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes a specified window to inherit the monitor of another window.
' ========================================================================================
PRIVATE FUNCTION AfxInheritWindowMonitor (BYVAL hwnd AS HWND, BYVAL hwndInherit AS HWND) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pInheritWindowMonitor AS FUNCTION (BYVAL hwnd AS ..HWND, BYVAL hwndInherit AS ..HWND) AS BOOLEAN
   pInheritWindowMonitor = DyLibSymbol(pLib, "InheritWindowMonitor")
   IF pInheritWindowMonitor THEN FUNCTION = pInheritWindowMonitor(hwnd, hwndInherit)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a specified DPI_AWARENESS_CONTEXT is valid and supported by the current system.
' ========================================================================================
PRIVATE FUNCTION AfxIsValidDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pIsValidDpiAwarenessContext AS FUNCTION (BYVAL value AS DPI_AWARENESS_CONTEXT) AS BOOLEAN
   pIsValidDpiAwarenessContext = DyLibSymbol(pLib, "IsValidDpiAwarenessContext")
   IF pIsValidDpiAwarenessContext THEN FUNCTION = pIsValidDpiAwarenessContext(value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a point in a window from logical coordinates into physical coordinates,
' regardless of the dots per inch (dpi) awareness of the caller.
' ========================================================================================
PRIVATE FUNCTION AfxLogicalToPhysicalPointForPerMonitorDPI (BYVAL hwnd AS HWND, BYREF pt AS POINT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pLogicalToPhysicalPointForPerMonitorDPI AS FUNCTION (BYVAL hwnd AS ..HWND, BYREF pt AS POINT) AS BOOLEAN
   pLogicalToPhysicalPointForPerMonitorDPI = DyLibSymbol(pLib, "LogicalToPhysicalPointForPerMonitorDPI")
   IF pLogicalToPhysicalPointForPerMonitorDPI THEN FUNCTION = pLogicalToPhysicalPointForPerMonitorDPI(hwnd, pt)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a point in a window from physical coordinates into logical coordinates,
' regardless of the dots per inch (dpi) awareness of the caller.
' ========================================================================================
PRIVATE FUNCTION AfxPhysicalToLogicalPointForPerMonitorDPI (BYVAL hwnd AS HWND, BYREF pt AS POINT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pPhysicalToLogicalPointForPerMonitorDPI AS FUNCTION (BYVAL hwnd AS ..HWND, BYREF pt AS POINT) AS BOOLEAN
   pPhysicalToLogicalPointForPerMonitorDPI = DyLibSymbol(pLib, "PhysicalToLogicalPointForPerMonitorDPI")
   IF pPhysicalToLogicalPointForPerMonitorDPI THEN FUNCTION = pPhysicalToLogicalPointForPerMonitorDPI(hwnd, pt)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Overrides the default per-monitor DPI scaling behavior of a child window in a dialog.
' ========================================================================================
PRIVATE FUNCTION AfxSetDialogControlDpiChangeBehavior (BYVAL hwnd AS HWND, BYVAL mask AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS, _
   BYVAL values AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pSetDialogControlDpiChangeBehavior AS FUNCTION (BYVAL hwnd AS ..HWND, BYVAL mask AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS, _
      BYVAL values AS DIALOG_CONTROL_DPI_CHANGE_BEHAVIORS) AS BOOLEAN
   pSetDialogControlDpiChangeBehavior = DyLibSymbol(pLib, "SetDialogControlDpiChangeBehavior")
   IF pSetDialogControlDpiChangeBehavior THEN FUNCTION = pSetDialogControlDpiChangeBehavior(hwnd, mask, values)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current process to a specified dots per inch (dpi) awareness context.
' ========================================================================================
PRIVATE FUNCTION AfxSetProcessDpiAwarenessContext (BYVAL value AS DPI_AWARENESS_CONTEXT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN FALSE
   DIM pSetProcessDpiAwarenessContext AS FUNCTION (BYVAL value AS DPI_AWARENESS_CONTEXT) AS BOOLEAN
   pSetProcessDpiAwarenessContext = DyLibSymbol(pLib, "SetProcessDpiAwarenessContext")
   IF pSetProcessDpiAwarenessContext THEN FUNCTION = pSetProcessDpiAwarenessContext(value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current process to a specified dots per inch (dpi) awareness context.
' ========================================================================================
PRIVATE FUNCTION AfxSetThreadCursorCreationScaling (BYVAL cursorDPI AS UINT) AS UINT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pSetThreadCursorCreationScaling AS FUNCTION (BYVAL cursorDPI AS UINT) AS UINT
   pSetThreadCursorCreationScaling = DyLibSymbol(pLib, "SetProcessDpiAwarenessContext")
   IF pSetThreadCursorCreationScaling THEN FUNCTION = pSetThreadCursorCreationScaling(cursorDPI)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set the DPI awareness for the current thread to the provided value.
' ========================================================================================
PRIVATE FUNCTION AfxSetThreadDpiAwarenessContext(BYVAL dpiContext AS DPI_AWARENESS_CONTEXT) AS DPI_AWARENESS_CONTEXT
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pSetThreadDpiAwarenessContext AS FUNCTION (BYVAL dpiContext AS DPI_AWARENESS_CONTEXT) AS DPI_AWARENESS_CONTEXT
   pSetThreadDpiAwarenessContext = DyLibSymbol(pLib, "SetThreadDpiAwarenessContext")
   IF pSetThreadDpiAwarenessContext THEN FUNCTION = pSetThreadDpiAwarenessContext(dpiContext)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the thread's DPI_HOSTING_BEHAVIOR. This behavior allows windows created in the thread
' to host child windows with a different DPI_AWARENESS_CONTEXT.
' ========================================================================================
PRIVATE FUNCTION AfxSetThreadDpiHostingBehavior (BYVAL value AS DPI_HOSTING_BEHAVIOR) AS DPI_HOSTING_BEHAVIOR
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pSetThreadDpiHostingBehavior AS FUNCTION (BYVAL value AS DPI_HOSTING_BEHAVIOR) AS DPI_HOSTING_BEHAVIOR
   pSetThreadDpiHostingBehavior = DyLibSymbol(pLib, "SetThreadDpiHostingBehavior")
   IF pSetThreadDpiHostingBehavior THEN FUNCTION = pSetThreadDpiHostingBehavior(value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the value of one of the system-wide parameters, taking into account the provided DPI value.
' ========================================================================================
PRIVATE FUNCTION AfxSystemParametersInfoForDpi (BYVAL uiAction AS UINT, BYVAL uiParam AS UINT, BYVAL pvParam AS PVOID, BYVAL fWinIni AS UINT, BYVAL dpi AS UINT) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad("user32.dll")
   IF pLib = NULL THEN RETURN 0
   DIM pSystemParametersInfoForDpi AS FUNCTION (BYVAL uiAction AS UINT, BYVAL uiParam AS UINT, BYVAL pvParam AS PVOID, BYVAL fWinIni AS UINT, BYVAL dpi AS UINT) AS BOOLEAN
   pSystemParametersInfoForDpi = DyLibSymbol(pLib, "SystemParametersInfoForDpi")
   IF pSystemParametersInfoForDpi THEN FUNCTION = pSystemParametersInfoForDpi(uiAction, uiParam, pvParam, fWinIni, dpi)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' // Identifies dots per inch (dpi) awareness values.
enum PROCESS_DPI_AWARENESS
  PROCESS_DPI_UNAWARE = 0
  PROCESS_SYSTEM_DPI_AWARE = 1
  PROCESS_PER_MONITOR_DPI_AWARE = 2
end enum

' // Identifies the dots per inch (dpi) setting for a monitor.
enum MONITOR_DPI_TYPE
  MDT_EFFECTIVE_DPI = 0
  MDT_ANGULAR_DPI = 1
  MDT_RAW_DPI = 2
  MDT_DEFAULT
end enum

' ========================================================================================
' Queries the dots per inch (dpi) of a display.
' ========================================================================================
PRIVATE FUNCTION AfxGetDpiForMonitor (BYVAL hMonitor AS HMONITOR, BYVAL dpiTYpe AS MONITOR_DPI_TYPE, _
   BYREF dpiX AS UINT, BYREF dpiY AS UINT) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad("shcore.dll")
   IF pLib = NULL THEN RETURN E_POINTER
   DIM pGetDpiForMonitor AS FUNCTION (BYVAL hMonitor AS ..HMONITOR, BYVAL dpiTYpe AS MONITOR_DPI_TYPE, _
       BYREF dpiX AS UINT, BYREF dpiY AS UINT) AS HRESULT
   pGetDpiForMonitor = DyLibSymbol(pLib, "GetDpiForMonitor")
   IF pGetDpiForMonitor THEN FUNCTION = pGetDpiForMonitor(hMonitor, dpiTYpe, dpiX, dpiY)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the dots per inch (dpi) awareness of the specified process.
' ========================================================================================
PRIVATE FUNCTION AfxGetProcessDpiAwareness (BYVAL hprocess AS HANDLE, BYREF value AS PROCESS_DPI_AWARENESS) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad("shcore.dll")
   IF pLib = NULL THEN RETURN E_POINTER
   DIM pGetProcessDpiAwareness AS FUNCTION (BYVAL hprocess AS HANDLE, BYREF value AS PROCESS_DPI_AWARENESS) AS HRESULT
   pGetProcessDpiAwareness = DyLibSymbol(pLib, "GetProcessDpiAwareness")
   IF pGetProcessDpiAwareness THEN FUNCTION = pGetProcessDpiAwareness(hprocess, value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the process-default DPI awareness level. This is equivalent to calling
' SetProcessDpiAwarenessContext with the corresponding DPI_AWARENESS_CONTEXT value.
' ========================================================================================
PRIVATE FUNCTION AfxSetProcessDpiAwareness (BYVAL value As PROCESS_DPI_AWARENESS) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad("shcore.dll")
   IF pLib = NULL THEN RETURN E_POINTER
   DIM pSetProcessDpiAwareness AS FUNCTION (BYVAL value AS PROCESS_DPI_AWARENESS) AS HRESULT
   pSetProcessDpiAwareness = DyLibSymbol(pLib, "SetProcessDpiAwareness")
   IF pSetProcessDpiAwareness THEN FUNCTION = pSetProcessDpiAwareness(value)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                              *** DPI RELATED PROCEDURES ***
' ########################################################################################

' ========================================================================================
' Sets the process-default DPI awareness to system-DPI awareness. This is equivalent to calling
' SetProcessDpiAwarenessContext with a DPI_AWARENESS_CONTEXT value of DPI_AWARENESS_CONTEXT_SYSTEM_AWARE.
' Note: SetProcessDPIAware is subject to a possible race condition if a DLL caches dpi
' settings during initialization. For this reason, it is recommended that dpi-aware be set
' through the application (.exe) manifest rather than by calling SetProcessDPIAware.
' Return value: TRUE on success; FALSE on failure.
' ========================================================================================
PRIVATE FUNCTION AfxSetProcessDPIAware () AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("user32.dll")
   IF pLib = 0 THEN RETURN FALSE
   DIM pProc AS FUNCTION () AS LONG
   pProc = DyLibSymbol(pLib, "SetProcessDPIAware")
   IF pProc = 0 THEN RETURN FALSE
   FUNCTION = pProc()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the current process is dots per inch (dpi) aware such that it adjusts
' the sizes of UI elements to compensate for the dpi setting.
' Return value: TRUE or FALSE
' [IsProcessDPIAware is available for use in the operating systems specified in the
' Requirements section. It may be altered or unavailable in subsequent versions.
' Instead, use GetProcessDPIAwareness.]
' ========================================================================================
PRIVATE FUNCTION AfxIsProcessDPIAware () AS BOOLEAN
   DIM AS ANY PTR pLib = DyLibLoad("user32.dll")
   IF pLib = 0 THEN RETURN FALSE
   DIM pProc AS FUNCTION () AS LONG
   pProc = DyLibSymbol(pLib, "IsProcessDPIAware")
   IF pProc = 0 THEN RETURN FALSE
   FUNCTION = pProc()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the value of the UseDpiScaling setting (Vista/Windows 7+).
' Returns TRUE if the OS uses DPI scaling; FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxUseDpiScaling () AS BOOLEAN
   DIM hkRes AS HKEY, dwType AS DWORD, dwData AS DWORD, cbData AS DWORD
   IF RegOpenKeyExW(HKEY_CURRENT_USER, "Software\Microsoft\Windows\DWM", 0, KEY_QUERY_VALUE, @hkRes) = ERROR_SUCCESS THEN
      IF hkRes THEN
         cbData = SIZEOF(cbData)
         DIM hr AS LONG = RegQueryValueExW(hkRes, "UseDpiScaling", 0, @dwType, CPTR(BYTE PTR, @dwData), @cbData)
         RegCloseKey hkRes
         IF hr = ERROR_SUCCESS THEN FUNCTION = (dwData <> 0)
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the number of pixels per logical inch along the screen width of the desktop
' window. In a system with multiple display monitors, this value is the same for all monitors.
' ========================================================================================
PRIVATE FUNCTION AfxLogPixelsX () AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM dpiX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC HWND_DESKTOP, hDC
   RETURN dpiX
END FUNCTION
' ========================================================================================
#define AfxGetDpi AfxLogPixelsX
#define AfxGetDpiX AfxLogPixelsX

' ========================================================================================
' Retrieves the number of pixels per logical inch along the screen height of the desktop
' window. In a system with multiple display monitors, this value is the same for all monitors.
' ========================================================================================
PRIVATE FUNCTION AfxLogPixelsY () AS LONG
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM dpiY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, hDC
   RETURN dpiY
END FUNCTION
' ========================================================================================
#define AfxGetDpiY AfxLogPixelsY

' ========================================================================================
' Retrieves the desktop horizontal scaling ratio.
' ========================================================================================
PRIVATE FUNCTION AfxScaleRatioX () AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM rxRatio AS SINGLE = (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
   RETURN rxRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the desktop vertical scaling ratio.
' ========================================================================================
PRIVATE FUNCTION AfxScaleRatioY () AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM ryRatio AS SINGLE = (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
   RETURN ryRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cx * (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cy * (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxUnscaleX (BYVAL cx AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cx / (GetDeviceCaps(hDC, LOGPIXELSX) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unscales a vertical coordinate according the DPI (dots per pixel) being used by the desktop.
' ========================================================================================
PRIVATE FUNCTION AfxUnscaleY (BYVAL cy AS SINGLE) AS SINGLE
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   FUNCTION = cy / (GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ReleaseDC HWND_DESKTOP, hDC
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if screen resolution meets minimum requirements.
' Parameters:
' - cxMin = Minimum screen resolution width in pixels.
' - cxMin = Minimum screen resolution height in pixels.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxIsResolutionAtLeast (BYVAL cxMin AS LONG, BYVAL cyMin AS LONG) AS BOOLEAN
   DIM ScreenWidth AS LONG = GetSystemMetrics(SM_CXSCREEN)
   DIM ScreenHeight AS LONG = GetSystemMetrics(SM_CYSCREEN)
   IF (cxMin <= ScreenWidth) AND (cyMin <= ScreenHeight) THEN FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if screen resolution meets minimum requirements in relative pixels,
' e.g. for a screen resolution of 1920x1080 pixels and a DPI of 192 (scaling ratio = 2),
' the maximum relative pixels for a DPI aware application is 960x540.
' - cxMin = Minimum screen resolution width in relative pixels.
' - cxMin = Minimum screen resolution height in relative pixels.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxIsDPIResolutionAtLeast (BYVAL cxMin AS LONG, BYVAL cyMin AS LONG) AS BOOLEAN
   ' // Get de DPI values used by the desktop window
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM dpiX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   DIM dpiY AS LONG = GetDeviceCaps(hDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, hDC
   ' // Scale the values
   cxMin = cxMin * dpiX / 96
   cyMin = cyMin * dpiX / 96
   ' // Calculate the width and height of the primary display monitor, in pixels
   DIM ScreenWidth AS LONG = GetSystemMetrics(SM_CXSCREEN)
   DIM ScreenHeight AS LONG = GetSystemMetrics(SM_CYSCREEN)
   IF (cxMin <= ScreenWidth) AND (cyMin <= ScreenHeight) THEN FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the logical width of the monitor that the window is currently displayed on.
' If the application to which the window belongs is not DPI aware, a monitor with a
' width resolution of 1920 pixels in a computer using 192 DPI, will return 960 pixels.
' ========================================================================================
PRIVATE FUNCTION AfxGetMonitorLogicalWidth (BYVAL hwnd AS HWND = NULL) AS DWORD
   ' // Get the monitor that the window is currently displayed on
   IF hwnd = NULL THEN hwnd = GetDesktopWindow
   DIM hMonitor AS HMONITOR = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
   ' // Get the logical width of the monitor.
   DIM miex AS MONITORINFOEXW
   miex.cbSize = SIZEOF(miex)
   GetMonitorInfoW(hMonitor, CAST(LPMONITORINFO, @miex))
   RETURN (miex.rcMonitor.right  - miex.rcMonitor.left)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the logical height of the monitor that the window is currently displayed on.
' If the application to which the window belongs is not DPI aware, a monitor with an
' height resolution of 1080 pixels in a computer using 192 DPI, will return 540 pixels.
' ========================================================================================
PRIVATE FUNCTION AfxGetMonitorLogicalHeight (BYVAL hwnd AS HWND = NULL) AS DWORD
   ' // Get the monitor that the window is currently displayed on
   IF hwnd = NULL THEN hwnd = GetDesktopWindow
   DIM hMonitor AS HMONITOR = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
   ' // Get the logical width of the monitor.
   DIM miex AS MONITORINFOEXW
   miex.cbSize = SIZEOF(miex)
   GetMonitorInfoW(hMonitor, CAST(LPMONITORINFO, @miex))
   RETURN (miex.rcMonitor.bottom - miex.rcMonitor.top)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the horizontal scaling of the monitor that the window is currently displayed on.
' If the application to which the window belongs is not DPI aware, a computer using
' 192 DPI, will return an scaling ratio of 2.
' ========================================================================================
PRIVATE FUNCTION AfxGetMonitorHorizontalScaling (BYVAL hwnd AS HWND = NULL) AS DWORD
   ' // Get the monitor that the window is currently displayed on
   IF hwnd = NULL THEN hwnd = GetDesktopWindow
   DIM hMonitor AS HMONITOR = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
   ' // Get the logical width and height of the monitor.
   DIM miex AS MONITORINFOEXW
   miex.cbSize = sizeof(miex)
   GetMonitorInfoW(hMonitor, CAST(LPMONITORINFO, @miex))
   DIM cxLogical AS LONG = (miex.rcMonitor.right  - miex.rcMonitor.left)
   DIM cyLogical AS LONG = (miex.rcMonitor.bottom - miex.rcMonitor.top)
   ' // Get the physical width and height of the monitor.
   DIM dm AS DEVMODEW
   dm.dmSize = sizeof(dm)
   dm.dmDriverExtra = 0
   EnumDisplaySettingsW(miex.szDevice, ENUM_CURRENT_SETTINGS, @dm)
   DIM cxPhysical AS LONG = dm.dmPelsWidth
   DIM cyPhysical as LONG = dm.dmPelsHeight
   ' // Calculate the scaling factor.
   DIM horzScale AS DOUBLE = cxPhysical / cxLogical
   DIM vertScale AS DOUBLE = cyPhysical / cyLogical
   RETURN horzScale
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the vertical scaling of the monitor that the window is currently displayed on.
' If the application to which the window belongs is not DPI aware, a computer using
' 192 DPI, will return an scaling ratio of 2.
' ========================================================================================
PRIVATE FUNCTION AfxGetMonitorVerticalScaling (BYVAL hwnd AS HWND = NULL) AS DWORD
   ' // Get the monitor that the window is currently displayed on
   IF hwnd = NULL THEN hwnd = GetDesktopWindow
   DIM hMonitor AS HMONITOR = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
   ' // Get the logical width and height of the monitor.
   DIM miex AS MONITORINFOEXW
   miex.cbSize = sizeof(miex)
   GetMonitorInfoW(hMonitor, CAST(LPMONITORINFO, @miex))
   DIM cxLogical AS LONG = (miex.rcMonitor.right  - miex.rcMonitor.left)
   DIM cyLogical AS LONG = (miex.rcMonitor.bottom - miex.rcMonitor.top)
   ' // Get the physical width and height of the monitor.
   DIM dm AS DEVMODEW
   dm.dmSize = sizeof(dm)
   dm.dmDriverExtra = 0
   EnumDisplaySettingsW(miex.szDevice, ENUM_CURRENT_SETTINGS, @dm)
   DIM cxPhysical AS LONG = dm.dmPelsWidth
   DIM cyPhysical as LONG = dm.dmPelsHeight
   ' // Calculate the scaling factor.
   DIM horzScale AS DOUBLE = cxPhysical / cxLogical
   DIM vertScale AS DOUBLE = cyPhysical / cyLogical
   RETURN vertScale
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads a specified icon resource with a client-specified system metric.
' See: https://msdn.microsoft.com/en-us/library/windows/desktop/bb775701(v=vs.85).aspx
' ========================================================================================
PRIVATE FUNCTION AfxLoadIconMetric (BYVAL hinst AS HINSTANCE, BYVAL pszName AS PCWSTR, BYVAL lims AS LONG, BYVAL phico AS HICON PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("comctl32.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pLoadIconMetric AS FUNCTION (BYVAL hinst AS HINSTANCE, BYVAL pszName AS PCWSTR, BYVAL lims AS LONG, BYVAL phico AS HICON PTR) AS HRESULT
   pLoadIconMetric = DyLibSymbol(pLib, "LoadIconMetric")
   IF pLoadIconMetric THEN FUNCTION = pLoadIconMetric(hinst, pszName, lims, phico)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

END NAMESPACE

