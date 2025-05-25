' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxImageList.bi
' Purpose:  ImageList declarations
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
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"
#include once "win/commctrl.bi"
'/* header files for imported files */
'#include "oaidl.h"
'#include "ocidl.h"

NAMESPACE Afx2
CONST CLSID_ImageList = "{7C476BA2-02B1-48f4-8048-B24619DDC058}"
CONST IID_IImageList ="{46EB5926-582E-4017-9FDF-E8998DAA0950}"
CONST IID_IImageList2 = "{192b9d83-50fc-457b-90a0-2b82a8b5dae1}"

' // Forward references
TYPE IImageList AS IImageList_
TYPE IImageList2 AS IImageList2_

' ========================================================================================
' Creates a single instance of an imagelist and returns an interface pointer to it.
' ========================================================================================
#ifndef __ImageList_CoCreateInstance__
#define __ImageList_CoCreateInstance__
PRIVATE FUNCTION ImageList_CoCreateInstance (BYVAL rclsid AS REFCLSID, BYVAL punkOuter As IUnknown PTR, BYVAL riid AS REFIID, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("Comctl32.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL rclsid AS REFCLSID, BYVAL punkOuter As IUnknown PTR, BYVAL riid AS REFIID, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   pProc = DyLibSymbol(pLib, "ImageList_CoCreateInstance")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(rclsid, punkOuter, riid, @ppv)
   DyLibFree(pLib)
END FUNCTION
#endif
' ========================================================================================

#define ILIF_ALPHA               &h00000001
#define ILIF_LOWQUALITY          &h00000002
#define ILDRF_IMAGELOWQUALITY    &h00000001
#define ILDRF_OVERLAYLOWQUALITY  &h00000010

' ========================================================================================
' IUnknown interface
' ========================================================================================
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
' ========================================================================================

' ########################################################################################
' Interface name = IImageList
' Inherited interface = IUnknown
' ########################################################################################
#ifndef __IImageList_INTERFACE_DEFINED__
#define __IImageList_INTERFACE_DEFINED__
TYPE IImageList_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Add (BYVAL hbmImage AS HBITMAP, BYVAL hbmMask AS HBITMAP, BYVAL pi AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReplaceIcon (BYVAL i AS LONG, BYVAL hicon AS HICON, BYVAL pi AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetOverlayImage (BYVAL iImage AS LONG, BYVAL iOverlay AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION Replace (BYVAL i AS LONG, BYVAL hbmImage AS HBITMAP, BYVAL hbmMask AS HBITMAP) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddMasked (BYVAL hbmImage AS HBITMAP, BYVAL crMask AS COLORREF, BYVAL pi AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Draw (BYVAL pimldp AS IMAGELISTDRAWPARAMS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL i AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL i AS LONG, BYVAL flags AS UINT, BYVAL picon AS HICON PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetImageInfo (BYVAL i AS LONG, BYVAL pImageInfo AS IMAGEINFO PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Copy (BYVAL iDist AS LONG, BYVAL punkSrc AS IUnknown PTR, BYVAL iSrc AS LONG, BYVAL uFlags AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION Merge (BYVAL i1 AS LONG, BYVAL punk2 AS IUnknown PTR, BYVAL i2 AS LONG, BYVAL dx AS LONG, BYVAL dy AS LONG, BYVAL riid AS REFIID, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Clone (BYVAL riid AS REFIID, BYVAL ppw AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetImageRect (BYVAL i AS LONG, BYVAL prc AS RECT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIconSize (BYVAL cx AS LONG PTR, BYVAL cy AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetIconSize (BYVAL cx AS LONG, BYVAL cy AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetImageCount (BYVAL pi AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetImageCount (BYVAL uNewCount AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetBkColor (BYVAL clrBk AS COLORREF, BYVAL pclr AS COLORREF PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetBkColor (BYVAL pclr AS COLORREF PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION BeginDrag (BYVAL iTrack AS LONG, BYVAL dxHotspot AS LONG, BYVAL dyHotspot AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION EndDrag () AS HRESULT
   DECLARE ABSTRACT FUNCTION DragEnter (BYVAL hwndLock AS HWND, BYVAL x AS LONG, BYVAL y AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION DragLeave (BYVAL hwndLock AS HWND) AS HRESULT
   DECLARE ABSTRACT FUNCTION DragMove (BYVAL x AS LONG, BYVAL y AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetDragCursorImage (BYVAL pUnk AS IUnknown PTR, BYVAL iDrag AS LONG, BYVAL dxHotspot AS LONG, BYVAL dyHotspot AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION DragShowNolock (BYVAL fShow AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDragImage (BYVAL ppt AS POINT PTR, BYVAL pptHotspot AS POINT PTR, BYVAL riid AS REFIID, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetItemFlags (BYVAL i AS LONG, BYVAL dwFlags AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetOverlayImage (BYVAL iOverlay AS LONG, BYVAL piIndex AS LONG PTR) AS HRESULT
END TYPE
TYPE LPIIMAGELIST AS IImageList PTR
#endif

#define ILR_DEFAULT                  &h0000
#define ILR_HORIZONTAL_LEFT          &h0000
#define ILR_HORIZONTAL_CENTER        &h0001
#define ILR_HORIZONTAL_RIGHT         &h0002
#define ILR_VERTICAL_TOP             &h0000
#define ILR_VERTICAL_CENTER          &h0010
#define ILR_VERTICAL_BOTTOM          &h0020
#define ILR_SCALE_CLIP               &h0000
#define ILR_SCALE_ASPECTRATIO        &h0100

#define ILGOS_ALWAYS       &h00000000
#define ILGOS_FROMSTANDBY  &h00000001
#define ILFIP_ALWAYS       &h00000000
#define ILFIP_FROMSTANDBY  &h00000001
#define ILDI_PURGE         &h00000001
#define ILDI_STANDBY       &h00000002
#define ILDI_RESETACCESS   &h00000004
#define ILDI_QUERYACCESS   &h00000008

TYPE IMAGELISTSTATS
   cbSize AS DWORD
   cAlloc AS LONG
   cUsed AS LONG
   cStandby AS LONG
END TYPE

' ########################################################################################
' Interface name = IImageList2
' Inherited interface = IImageList
' ########################################################################################
#ifndef __IImageList2_INTERFACE_DEFINED__
#define __IImageList2_INTERFACE_DEFINED__
TYPE IImageList2_ EXTENDS IImageList
   DECLARE ABSTRACT FUNCTION Resize (BYVAL cxNewIconSize AS LONG, BYVAL cyNewIconSize AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetOriginalSize (BYVAL iImage AS LONG, BYVAL dwFlags AS DWORD, BYVAL pcx AS LONG PTR, BYVAL pcy AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetOriginalSize (BYVAL iImage AS LONG, BYVAL cx AS LONG, BYVAL cy AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetCallback (BYVAL pUnk AS IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetCallback (BYVAL riid AS REFIID, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ForceImagePresent (BYVAL iImage AS LONG, BYVAL dwFlags AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION DiscardImages (BYVAL iFirstImage AS LONG, BYVAL iLastImage AS LONG, BYVAL dwFlags AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION PreloadImages (BYVAL pimldp AS IMAGELISTDRAWPARAMS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetStatistics (BYVAL pils AS IMAGELISTSTATS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Initialize (BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL flags AS UINT, BYVAL cInitial AS LONG, BYVAL cGrow AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION Replace2 (BYVAL i AS LONG, BYVAL hbmImage AS HBITMAP, BYVAL hbmMask AS HBITMAP, BYVAL punk AS IUnknown PTR, BYVAL dwFlags AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReplaceFromImageList (BYVAL i AS LONG, BYVAL pil AS IImageList PTR, BYVAL iSrc AS LONG, BYVAL punk AS IUnknown PTR, BYVAL dwFlags AS DWORD) AS HRESULT
END TYPE
TYPE LPIIMAGELIST2 AS IImageList2 PTR
#endif

END NAMESPACE
