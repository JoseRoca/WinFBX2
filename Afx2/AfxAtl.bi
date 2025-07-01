' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxAtl.bi
' Purpose:  ATL procedures and interfaces
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
#include once "windows.bi"
#include once "Afx2/AfxCom2.inc"

CONST ATL_DLLNAME = "ATL.DLL"   ' --> change as needed

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

NAMESPaCE Afx2

' ########################################################################################
' Base types for declaring ABSTRACT interface methods in other types that inherit from
' these ones. Afx_IUnknown extends the built-in OBJECT type which provides run-time type
' information for all types derived from it using Extends. Extending the built-in Object
' type allows to add an extra hidden vtable pointer field at the top of the Type. The
' vtable is used to dispatch Virtual and Abstract methods and to access information for
' run-time type identification used by Operator Is.
' ########################################################################################

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_Iunknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) as HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

' ########################################################################################

' ========================================================================================
' Creates a connection between an object's connection point and a client's sink.
' ========================================================================================
PRIVATE FUNCTION AtlAdvise (BYVAL pUnkCP AS IUnknown PTR, BYVAL pUnk AS IUnknown PTR, BYREF iid AS IID, BYREF pdw AS DWORD) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAdvise AS FUNCTION (BYVAL pUnkCP AS IUnknown PTR, BYVAL pUnk AS IUnknown PTR, BYREF iid AS ..IID, BYREF pdw AS DWORD) AS HRESULT
   pAtlAdvise = DyLibSymbol(pLib, "AtlAdvise")
   IF pAtlAdvise THEN FUNCTION = pAtlAdvise(pUnkCP, pUnk, iid, pdw)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches a previously created control to the specified window.
' ========================================================================================
PRIVATE FUNCTION AtlAxAttachControl (BYVAL pControl AS IUnknown PTR, BYVAL hwnd AS ..HWND, BYVAL ppUnkContainer AS IUnknown PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxAttachControl AS FUNCTION (BYVAL pControl AS IUnknown PTR, BYVAL hwnd AS ..HWND, BYVAL ppUnkContainer AS IUnknown PTR PTR) AS HRESULT
   pAtlAxAttachControl = DyLibSymbol(pLib, "AtlAxAttachControl")
   IF pAtlAxAttachControl THEN FUNCTION = pAtlAxAttachControl(pControl, hwnd, ppUnkContainer)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ActiveX control, initializes it, and hosts it in the specified window.
' ========================================================================================
PRIVATE FUNCTION AtlAxCreateControl (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS ..HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxCreateControl AS FUNCTION (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS ..HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR) AS HRESULT
   pAtlAxCreateControl = DyLibSymbol(pLib, "AtlAxCreateControl")
   IF pAtlAxCreateControl THEN FUNCTION = pAtlAxCreateControl(pwszName, hwnd, pStream, ppUnkContainer)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ActiveX control, initializes it, and hosts it in the specified window. An
' interface pointer and event sink for the new control can also be created.
' ========================================================================================
PRIVATE FUNCTION AtlAxCreateControlEx (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR, _
   BYVAL ppUnkControl AS IUnknown PTR PTR, BYVAL iidSink AS REFIID = @IID_NULL, BYVAL punkSink AS IUnknown PTR = NULL) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxCreateControlEx AS FUNCTION (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS ..HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR, _
      BYVAL ppUnkControl AS IUnknown PTR PTR, BYVAL iidSink AS REFIID = @IID_NULL, BYVAL punkSink AS IUnknown PTR = NULL) AS HRESULT
   pAtlAxCreateControlEx = DyLibSymbol(pLib, "AtlAxCreateControlEx")
   IF pAtlAxCreateControlEx THEN FUNCTION = pAtlAxCreateControlEx(pwszName, hwnd, pStream, ppUnkContainer, ppUnkControl, iidSink, punkSink)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a licensed ActiveX control, initializes it, and hosts it in the specified window.
' ========================================================================================
PRIVATE FUNCTION AtlAxCreateControlLic (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS ..HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR, BYVAL bstrLic AS BSTR = NULL) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxCreateControlLic AS FUNCTION (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS ..HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR, BYVAL bstrLic AS BSTR) AS HRESULT
   pAtlAxCreateControlLic = DyLibSymbol(pLib, "AtlAxCreateControlLic")
   IF pAtlAxCreateControlLic THEN FUNCTION = pAtlAxCreateControlLic(pwszName, hwnd, pStream, ppUnkContainer, bstrLic)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a licensed ActiveX control, initializes it, and hosts it in the specified window.
' ========================================================================================
PRIVATE FUNCTION AtlAxCreateControlLicEx (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR, _
   BYVAL ppUnkControl AS IUnknown PTR PTR, BYVAL iidSink AS REFIID = @IID_NULL, BYVAL punkSink AS IUnknown PTR = NULL, BYVAL bstrLic AS BSTR = NULL) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxCreateControlLicEx AS FUNCTION (BYVAL pwszName AS WSTRING PTR, BYVAL hwnd AS ..HWND, BYVAL pStream AS IStream PTR, BYVAL ppUnkContainer AS IUnknown PTR PTR, _
      BYVAL ppUnkControl AS IUnknown PTR PTR, BYVAL iidSink AS REFIID = @IID_NULL, BYVAL punkSink AS IUnknown PTR = NULL, BYVAL bstrLic AS BSTR = NULL) AS HRESULT
   pAtlAxCreateControlLicEx = DyLibSymbol(pLib, "AtlAxCreateControlLicEx")
   IF pAtlAxCreateControlLicEx THEN FUNCTION = pAtlAxCreateControlLicEx(pwszName, hwnd, pStream, ppUnkContainer, ppUnkControl, iidSink, punkSink, bstrLic)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a modeless dialog box from a dialog template provided by the user.
' ========================================================================================
PRIVATE FUNCTION AtlAxCreateDialog (BYVAL hInst AS HINSTANCE, BYVAL lpTemplateName AS WSTRING PTR, BYVAL hwndParent AS HWND, _
   BYVAL lpDialogProc AS DLGPROC, BYVAL dwInitParam AS LPARAM) AS HWND
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxCreateDialog AS FUNCTION (BYVAL hInst AS HINSTANCE, BYVAL lpTemplateName AS WSTRING PTR, BYVAL hwndParent AS HWND, BYVAL lpDialogProc AS DLGPROC, BYVAL dwInitParam AS LPARAM) AS HWND
   pAtlAxCreateDialog = DyLibSymbol(pLib, "AtlAxCreateDialogW")
   IF pAtlAxCreateDialog THEN FUNCTION = pAtlAxCreateDialog(hInst, lpTemplateName, hwndParent, lpDialogProc, dwInitParam)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a modal dialog box from a dialog template provided by the user.
' ========================================================================================
PRIVATE FUNCTION AtlAxDialogBox (BYVAL hInst AS HINSTANCE, BYVAL lpTemplateName AS WSTRING PTR, BYVAL hwndParent AS HWND, _
   BYVAL lpDialogProc AS DLGPROC, BYVAL dwInitParam AS LPARAM) AS INT_PTR
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxDialogBox AS FUNCTION (BYVAL hInst AS HINSTANCE, BYVAL lpTemplateName AS WSTRING PTR, BYVAL hwndParent AS HWND, BYVAL lpDialogProc AS DLGPROC, BYVAL dwInitParam AS LPARAM) AS INT_PTR
   pAtlAxDialogBox = DyLibSymbol(pLib, "AtlAxDialogBoxW")
   IF pAtlAxDialogBox THEN FUNCTION = pAtlAxDialogBox(hInst, lpTemplateName, hwndParent, lpDialogProc, dwInitParam)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Obtains a direct interface pointer to the control contained inside a specified window
' given its handle.
' ========================================================================================
PRIVATE FUNCTION AtlAxGetControl OVERLOAD (BYVAL hwnd AS ..HWND, BYVAL pUnk AS IUnknown PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxGetControl AS FUNCTION (BYVAL hwnd AS ..HWND, BYVAL pUnk AS IUnknown PTR PTR) AS HRESULT
   pAtlAxGetControl = DyLibSymbol(pLib, "AtlAxGetControl")
   IF pAtlAxGetControl THEN FUNCTION = pAtlAxGetControl(hwnd, pUnk)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AtlAxGetControl OVERLOAD (BYVAL hwnd AS ..HWND) AS IUnknown PTR
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxGetControl AS FUNCTION (BYVAL hwnd AS ..HWND, BYVAL pUnk AS IUnknown PTR PTR) AS HRESULT
   pAtlAxGetControl = DyLibSymbol(pLib, "AtlAxGetControl")
   DIM pUnk AS IUnknown PTR
   IF pAtlAxGetControl THEN pAtlAxGetControl(hwnd, @pUnk)
   DyLibFree(pLib)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Obtains a direct interface pointer to the container for a specified window (if any),
' given its handle.
' ========================================================================================
PRIVATE FUNCTION AtlAxGetHost (BYVAL hwnd AS ..HWND, BYVAL pUnk AS IUnknown PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxGetHost AS FUNCTION (BYVAL hwnd AS ..HWND, BYVAL pUnk AS IUnknown PTR PTR) AS HRESULT
   pAtlAxGetHost = DyLibSymbol(pLib, "AtlAxGetHost")
   IF pAtlAxGetHost THEN FUNCTION = pAtlAxGetHost(hwnd, pUnk)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' This function initializes ATL's control hosting code by registering the "AtlAxWin" and
' "AtlAxWinLic" window classes plus a couple of custom window messages.
' ========================================================================================
PRIVATE FUNCTION AtlAxWinInit () AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlAxWinInit AS FUNCTION () AS BOOLEAN
   pAtlAxWinInit = DyLibSymbol(pLib, "AtlAxWinInit")
   IF pAtlAxWinInit THEN FUNCTION = pAtlAxWinInit()
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an interface pointer to another interface pointer of the same type.
' ========================================================================================
PRIVATE FUNCTION AtlComPtrAssign (BYVAL pp AS IUnknown PTR PTR, BYVAL p AS IUnknown PTR) AS IUnknown PTR
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlComPtrAssign AS FUNCTION (BYVAL pp AS IUnknown PTR PTR, BYVAL p AS IUnknown PTR) AS IUnknown PTR
   pAtlComPtrAssign = DyLibSymbol(pLib, "AtlComPtrAssign")
   IF pAtlComPtrAssign THEN FUNCTION = pAtlComPtrAssign(pp, p)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a device context for the device specified in the DVTARGETDEVICE structure.
' ========================================================================================
PRIVATE FUNCTION AtlCreateTargetDC (BYVAL hdc_ AS ..HDC, BYVAL dv AS DVTARGETDEVICE PTR) AS HDC
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlCreateTargetDC AS FUNCTION (BYVAL hdc_ AS ..HDC, BYVAL dv AS DVTARGETDEVICE PTR) AS HDC
   pAtlCreateTargetDC = DyLibSymbol(pLib, "AtlCreateTargetDC")
   IF pAtlCreateTargetDC THEN FUNCTION = pAtlCreateTargetDC(hdc_, dv)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the marshal data in the stream, then releases the stream pointer.
' ========================================================================================
PRIVATE FUNCTION AtlFreeMarshalStream (BYVAL pStream AS IStream PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlFreeMarshalStream AS FUNCTION (BYVAL pStream AS IStream PTR) AS HRESULT
   pAtlFreeMarshalStream = DyLibSymbol(pLib, "AtlFreeMarshalStream")
   IF pAtlFreeMarshalStream THEN FUNCTION = pAtlFreeMarshalStream(pStream)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the marshal data in the stream, then releases the stream pointer.
' ========================================================================================
PRIVATE FUNCTION AtlDevModeW2A (BYVAL lpDevModeA AS DEVMODEA PTR, BYVAL lpDevModeW AS DEVMODEW PTR) AS DEVMODEA PTR
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlDevModeW2A AS FUNCTION (BYVAL lpDevModeA AS DEVMODEA PTR, BYVAL lpDevModeW AS DEVMODEW PTR) AS DEVMODEA PTR
   pAtlDevModeW2A = DyLibSymbol(pLib, "AtlDevModeW2A")
   IF pAtlDevModeW2A THEN FUNCTION = pAtlDevModeW2A(lpDevModeA, lpDevModeW)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the marshal data in the stream, then releases the stream pointer.
' ========================================================================================
PRIVATE FUNCTION AtlGetVersion (BYVAL pReserved AS ANY PTR = NULL) AS DWORD
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlGetVersion AS FUNCTION (BYVAL pReserved AS ANY PTR = NULL) AS DWORD
   pAtlGetVersion = DyLibSymbol(pLib, "AtlGetVersion")
   IF pAtlGetVersion THEN FUNCTION = pAtlGetVersion(pReserved)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an object's size in HIMETRIC units (each unit is 0.01 millimeter) to a size in
' pixels on the screen device.
' ========================================================================================
PRIVATE SUB AtlHiMetricToPixel (BYVAL lpSizeInHiMetric AS SIZEL PTR, BYVAL lpSizeInPix AS SIZEL PTR)
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN EXIT SUB
   DIM pAtlHiMetricToPixel AS SUB (BYVAL lpSizeInHiMetric AS SIZEL PTR, BYVAL lpSizeInPix AS SIZEL PTR)
   pAtlHiMetricToPixel = DyLibSymbol(pLib, "AtlHiMetricToPixel")
   IF pAtlHiMetricToPixel THEN pAtlHiMetricToPixel(lpSizeInHiMetric, lpSizeInPix)
   DyLibFree(pLib)
END SUB
' ========================================================================================

' ========================================================================================
' Releases the marshal data in the stream, then releases the stream pointer.
' ========================================================================================
PRIVATE FUNCTION AtlLoadTypeLib (BYVAL hinst AS HINSTANCE, BYVAL lpszIndex AS WSTRING PTR, BYVAL pbstrPath AS BSTR PTR, BYVAL ppTypeLib AS ITypeLIb PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlLoadTypeLib AS FUNCTION (BYVAL hinst AS HINSTANCE, BYVAL lpszIndex AS WSTRING PTR, BYVAL pbstrPath AS BSTR PTR, BYVAL ppTypeLib AS ITypeLIb PTR PTR) AS HRESULT
   pAtlLoadTypeLib = DyLibSymbol(pLib, "AtlLoadTypeLib")
   IF pAtlLoadTypeLib THEN FUNCTION = pAtlLoadTypeLib(hinst, lpszIndex, pbstrPath, ppTypeLib)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the marshal data in the stream, then releases the stream pointer.
' ========================================================================================
PRIVATE FUNCTION AtlMarshalPtrInProc (BYVAL pUnk AS IUnknown PTR, BYVAL iid AS ..IID PTR, BYVAL pstm AS IStream PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlMarshalPtrInProc AS FUNCTION (BYVAL pUnk AS IUnknown PTR, BYVAL iid AS ..IID PTR, BYVAL pstm AS IStream PTR PTR) AS HRESULT
   pAtlMarshalPtrInProc = DyLibSymbol(pLib, "AtlMarshalPtrInProc")
   IF pAtlMarshalPtrInProc THEN FUNCTION = pAtlMarshalPtrInProc(pUnk, iid, pstm)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an object's size in HIMETRIC units (each unit is 0.01 millimeter) to a size in
' pixels on the screen device.
' ========================================================================================
PRIVATE SUB AtlPixelToHiMetric (BYVAL lpPix AS SIZEL PTR, BYVAL lpHiMetric AS SIZEL PTR)
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN EXIT SUB
   DIM pAtlPixelToHiMetric AS SUB (BYVAL lpPix AS SIZEL PTR, BYVAL lpHiMetric AS SIZEL PTR)
   pAtlPixelToHiMetric = DyLibSymbol(pLib, "AtlPixelToHiMetric")
   IF pAtlPixelToHiMetric THEN pAtlPixelToHiMetric(lpPix, lpHiMetric)
   DyLibFree(pLib)
END SUB
' ========================================================================================

' ========================================================================================
' Converts an object's size in HIMETRIC units (each unit is 0.01 millimeter) to a size in
' pixels on the screen device.
' ========================================================================================
PRIVATE FUNCTION AtlRegisterTypeLib (BYVAL hInstTypeLib AS HINSTANCE, BYVAL lpszIndex AS WSTRING PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlRegisterTypeLib AS FUNCTION (BYVAL hInstTypeLib AS HINSTANCE, BYVAL lpszIndex AS WSTRING PTR) AS HRESULT
   pAtlRegisterTypeLib = DyLibSymbol(pLib, "AtlRegisterTypeLib")
   IF pAtlRegisterTypeLib THEN FUNCTION = pAtlRegisterTypeLib(hInstTypeLib, lpszIndex)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an object's size in HIMETRIC units (each unit is 0.01 millimeter) to a size in
' pixels on the screen device.
' ========================================================================================
PRIVATE FUNCTION AtlSetErrorInfo (BYVAL clsid AS CLSID PTR, BYVAL lpszDesc AS WSTRING PTR, BYVAL dwHelpID AS DWORD, _
   BYVAL lpszHelpFile AS WSTRING PTR, BYVAL iid AS ..IID PTR, BYVAL hRes AS HRESULT, BYVAL hInst AS HINSTANCE) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlSetErrorInfo AS FUNCTION (BYVAL clsid AS ..CLSID PTR, BYVAL lpszDesc AS WSTRING PTR, BYVAL dwHelpID AS DWORD, _
      BYVAL lpszHelpFile AS WSTRING PTR, BYVAL iid AS ..IID PTR, BYVAL hRes AS HRESULT, BYVAL hInst AS HINSTANCE) AS HRESULT
   pAtlSetErrorInfo = DyLibSymbol(pLib, "AtlSetErrorInfo")
   IF pAtlSetErrorInfo THEN FUNCTION = pAtlSetErrorInfo(clsid, lpszDesc, dwHelpID, lpszHelpFile, iid, hRes, hInst)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Terminates the connection established through AtlAdvise.
' ========================================================================================
PRIVATE FUNCTION AtlUnadvise (BYVAL pUnkCP AS IUnknown PTR, BYVAL iid AS ..IID PTR, BYVAL dw AS DWORD) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlUnadvise AS FUNCTION (BYVAL pUnkCP AS IUnknown PTR, BYVAL iid AS ..IID PTR, BYVAL dw AS DWORD) AS HRESULT
   pAtlUnadvise = DyLibSymbol(pLib, "AtlUnadvise")
   IF pAtlUnadvise THEN FUNCTION = pAtlUnadvise(pUnkCP, iid, dw)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the marshal data in the stream, then releases the stream pointer.
' ========================================================================================
PRIVATE FUNCTION AtlUnmarshalPtr (BYVAL pstm AS IStream PTR, BYVAL iid AS ..IID PTR, BYVAL ppUnk AS IUnknown PTR PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlUnmarshalPtr AS FUNCTION (BYVAL pstm AS IStream PTR, BYVAL iid AS ..IID PTR, BYVAL ppUnk AS IUnknown PTR PTR) AS HRESULT
   pAtlUnmarshalPtr = DyLibSymbol(pLib, "AtlUnmarshalPtr")
   IF pAtlUnmarshalPtr THEN FUNCTION = pAtlUnmarshalPtr(pstm, iid, ppUnk)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' This function is called to unregister a type library.
' ========================================================================================
PRIVATE FUNCTION AtlUnRegisterTypeLib (BYVAL hInstTypeLib AS HINSTANCE, BYVAL lpszIndex AS WSTRING PTR) AS HRESULT
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlUnRegisterTypeLib AS FUNCTION (BYVAL hInstTypeLib AS HINSTANCE, BYVAL lpszIndex AS WSTRING PTR) AS HRESULT
   pAtlUnRegisterTypeLib = DyLibSymbol(pLib, "AtlUnRegisterTypeLib")
   IF pAtlUnRegisterTypeLib THEN FUNCTION = pAtlUnRegisterTypeLib(hInstTypeLib, lpszIndex)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' This function is called to unregister a type library.
' ========================================================================================
PRIVATE FUNCTION AtlWaitWithMessageLoop (BYVAL hEvent AS HANDLE) AS BOOLEAN
   DIM pLib AS ANY PTR = DyLibLoad(ATL_DLLNAME)
   IF pLib = NULL THEN RETURN FALSE
   DIM pAtlWaitWithMessageLoop AS FUNCTION (BYVAL hEvent AS HANDLE) AS BOOLEAN
   pAtlWaitWithMessageLoop = DyLibSymbol(pLib, "AtlWaitWithMessageLoop")
   IF pAtlWaitWithMessageLoop THEN FUNCTION = pAtlWaitWithMessageLoop(hEvent)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the interface of the ActiveX control given the handle of its ATL container
' ========================================================================================
FUNCTION AtlAxGetDispatch (BYVAL hwndControl AS HWND) AS IDispatch PTR
   ' // Get the IUnknown of the OCX hosted in the control
   DIM pUnk AS IUnknown PTR
   DIM hr AS HRESULT = AtlAxGetControl(hwndControl, @pUnk)
   IF hr <> 0 THEN EXIT FUNCTION
   ' // Query for the IDispatch interface
   DIM pDispatch AS IDispatch PTR
   hr = pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDispatch)
   IF hr = S_OK THEN FUNCTION = pDispatch
END FUNCTION
' ========================================================================================

' ########################################################################################
' Library name: ATLLib
' Documentation string: ATL 2.0 Type Library
' GUID: {44EC0535-400F-11D0-9DCD-00A0C90391D3}
' Version: 1.0, Locale ID = 0
' Path: C:\Windows\SysWOW64\atl.dll
' Attributes: 8 [&h00000008]  [HasDiskImage]
' Code generated by the TypeLib Browser for Free Basic 1.00
' © 2016-2020 by José Roca. All rights reserved. Freeware. Use at your own risk.
' Date: 22 may. 2025 Time: 02:32:37
' ########################################################################################

' // Dual interfaces - Forward references
TYPE Afx_IAxWinAmbientDispatch AS Afx_IAxWinAmbientDispatch_
' // Dispatchable interfaces - Forward references
TYPE Afx_IDocHostUIHandlerDispatch AS Afx_IDocHostUIHandlerDispatch_

' ########################################################################################
' Interface name: IDocHostUIHandlerDispatch
' IID: {425B5AF0-65F1-11D1-9611-0000F81E0D0D}
' Documentation string: IDocHostUIHandlerDispatch Interface
' Attributes =  4096 [&h00001000] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 15
' ########################################################################################

#ifndef __Afx_IDocHostUIHandlerDispatch_INTERFACE_DEFINED__
#define __Afx_IDocHostUIHandlerDispatch_INTERFACE_DEFINED__

TYPE Afx_IDocHostUIHandlerDispatch_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION ShowContextMenu (BYVAL dwID AS ULONG, BYVAL x AS ULONG, BYVAL y AS ULONG, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR, BYVAL dwRetVal AS HRESULT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHostInfo (BYVAL pdwFlags AS ULONG PTR, BYVAL pdwDoubleClick AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowUI (BYVAL dwID AS ULONG, BYVAL pActiveObject AS IUnknown PTR, BYVAL pCommandTarget AS IUnknown PTR, BYVAL pFrame AS IUnknown PTR, BYVAL pDoc AS IUnknown PTR, BYVAL dwRetVal AS HRESULT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION HideUI () AS HRESULT
   DECLARE ABSTRACT FUNCTION UpdateUI () AS HRESULT
   DECLARE ABSTRACT FUNCTION EnableModeless (BYVAL fEnable AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnDocWindowActivate (BYVAL fActivate AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnFrameWindowActivate (BYVAL fActivate AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION ResizeBorder (BYVAL left AS LONG, BYVAL top AS LONG, BYVAL right AS LONG, BYVAL bottom AS LONG, BYVAL pUIWindow AS IUnknown PTR, BYVAL fFrameWindow AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION TranslateAccelerator (BYVAL hWnd AS ULONG, BYVAL nMessage AS ULONG, BYVAL wParam AS ULONG, BYVAL lParam AS ULONG, BYVAL bstrGuidCmdGroup AS BSTR, BYVAL nCmdID AS ULONG, BYVAL dwRetVal AS HRESULT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetOptionKeyPath (BYVAL pbstrKey AS BSTR PTR, BYVAL dw AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDropTarget (BYVAL pDropTarget AS IUnknown PTR, BYVAL ppDropTarget AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetExternal (BYVAL ppDispatch AS IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION TranslateUrl (BYVAL dwTranslate AS ULONG, BYVAL bstrURLIn AS BSTR, BYVAL pbstrURLOut AS BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FilterDataObject (BYVAL pDO AS IUnknown PTR, BYVAL ppDORet AS IUnknown PTR PTR) AS HRESULT
END TYPE

#endif
' ########################################################################################


' ########################################################################################
' Interface name: IAxWinAmbientDispatch
' IID: {B6EA2051-048A-11D1-82B9-00C04FB9942E}
' Documentation string: IAxWinAmbientDispatch Interface
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 28
' ########################################################################################

#ifndef __Afx_IAxWinAmbientDispatch_INTERFACE_DEFINED__
#define __Afx_IAxWinAmbientDispatch_INTERFACE_DEFINED__

TYPE Afx_IAxWinAmbientDispatch_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION put_AllowWindowlessActivation (BYVAL pbCanWindowlessActivate AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AllowWindowlessActivation (BYVAL pbCanWindowlessActivate AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_BackColor (BYVAL pclrBackground AS OLE_COLOR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BackColor (BYVAL pclrBackground AS OLE_COLOR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ForeColor (BYVAL pclrForeground AS OLE_COLOR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ForeColor (BYVAL pclrForeground AS OLE_COLOR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_LocaleID (BYVAL plcidLocaleID AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_LocaleID (BYVAL plcidLocaleID AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_UserMode (BYVAL pbUserMode AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_UserMode (BYVAL pbUserMode AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DisplayAsDefault (BYVAL pbDisplayAsDefault AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DisplayAsDefault (BYVAL pbDisplayAsDefault AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Font (BYVAL pFont AS IFontDisp PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Font (BYVAL pFont AS IFontDisp PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MessageReflect (BYVAL pbMsgReflect AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MessageReflect (BYVAL pbMsgReflect AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShowGrabHandles (BYVAL pbShowGrabHandles AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShowHatching (BYVAL pbShowHatching AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DocHostFlags (BYVAL pdwDocHostFlags AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DocHostFlags (BYVAL pdwDocHostFlags AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DocHostDoubleClickFlags (BYVAL pdwDocHostDoubleClickFlags AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DocHostDoubleClickFlags (BYVAL pdwDocHostDoubleClickFlags AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AllowContextMenu (BYVAL pbAllowContextMenu AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AllowContextMenu (BYVAL pbAllowContextMenu AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AllowShowUI (BYVAL pbAllowShowUI AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AllowShowUI (BYVAL pbAllowShowUI AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_OptionKeyPath (BYVAL pbstrOptionKeyPath AS BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_OptionKeyPath (BYVAL pbstrOptionKeyPath AS BSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

END NAMESPACE
