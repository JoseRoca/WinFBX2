' ########################################################################################
' Platform: Microsoft Windows
' Filename: CImageList.inc
' Purpose:  Image list wrapper class.
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
#include once "Afx2/AfxImageList.bi"
#include once "Afx2/CControl.inc"

NAMESPACE Afx2

' ========================================================================================
' Macro for debug
' To allow debugging, define _CIMGL_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CIMGL_DEBUG_
   #define _CIMGL_DEBUG_ 0
#ENDIF
#ifndef _CIMGL_DP_
   #define _CIMGL_DP_ 1
   #MACRO CIMGL_DP(st)
      #IF (_CIMGL_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' CImageList class
' ========================================================================================
TYPE CImageList EXTENDS CControl

Public:
   m_hImageList AS ..HIMAGELIST   ' // Image list handle
   m_hInst AS ..HMODULE           ' // Instance handle
   DIM m_Overlays(1 TO 15) AS LONG

Public:
   DECLARE CONSTRUCTOR (BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL dwFlags AS DWORD = 0, BYVAL cInitial AS LONG = 0, BYVAL cGrow AS LONG = 0)
   DECLARE DESTRUCTOR

   DECLARE PROPERTY hImageList () AS ..HIMAGELIST
   DECLARE FUNCTION LoadImage (BYREF wszImage AS WSTRING, BYVAL dwType AS DWORD, BYVAL cxDesired AS LONG = 0, _
      BYVAL cyDesired AS LONG = 0, BYVAL bShared AS BOOLEAN = FALSE) AS ..HIMAGELIST
   DECLARE FUNCTION LoadBitmap (BYREF wszImage AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYREF wszMask AS WSTRING = "") AS LONG
   DECLARE FUNCTION AddBitmap (BYVAL hImage AS HBITMAP, BYVAL hMask AS HBITMAP = NULL) AS LONG
   DECLARE FUNCTION AddIcon (BYVAL hImage AS HICON) AS LONG
   DECLARE FUNCTION LoadIcon (BYREF wszImage AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0) AS LONG
   DECLARE FUNCTION LoadResIcon (BYVAL wID AS WORD, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION LoadResBitmap (BYVAL wID AS WORD, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL hMask AS HBITMAP, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION LoadNamedResIcon (BYREF wszName AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION LoadNamedResBitmap (BYREF wszName AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL hMask AS HBITMAP, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   DECLARE FUNCTION AddMasked (BYVAL hImage AS HBITMAP, BYVAL dwMaskColor AS COLORREF = 0) AS LONG
   DECLARE FUNCTION LoadMasked (BYREF wszImage AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL dwMaskColor AS COLORREF) AS LONG
   DECLARE FUNCTION LoadResBitmapMasked (BYVAL wID AS WORD, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL dwMaskColor AS COLORREF) AS LONG
   DECLARE FUNCTION LoadNamedResBitmapMasked (BYREF wszName AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL dwMaskColor AS COLORREF) AS LONG
   DECLARE SUB OverlayImage (BYVAL iImage AS LONG, BYVAL iOverlay AS LONG)
   DECLARE FUNCTION OverlayIndex (BYVAL iImage AS LONG) AS LONG
   DECLARE FUNCTION Duplicate () AS ..HIMAGELIST
   DECLARE FUNCTION Copy (BYVAL iDst AS LONG, BYVAL iSrc AS LONG, BYVAL dwFlag AS DWORD) AS BOOLEAN
   DECLARE FUNCTION Remove (BYVAL idx AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetIconSize (BYREF cx AS LONG, BYREF cy AS LONG) AS BOOLEAN
   DECLARE PROPERTY ImageCount () AS LONG
   DECLARE PROPERTY BackgroundColor () AS COLORREF
   DECLARE PROPERTY BackgroundColor (BYVAL clrBk AS COLORREF)
   
END TYPE
' ========================================================================================

' ========================================================================================
' CImageList constructor
' Creates the image list
' - cx       = The width, in pixels, of each image.
' - cy       = The height, in pixels, of each image.
' - dwFlags  = A set of bit flags that specify the type of image list to create.
' - cInitial = The number of images that the image list initially contains.
' - cGrow    = The number of images by which the image list can grow when the system
'              needs to make room for new images. This parameter represents the number
'              of new images that the resized image list can contain.
' WARNING: Don't destroy the image list. The class does it when it is destroyed.
' ========================================================================================
PRIVATE CONSTRUCTOR CImageList (BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL dwFlags AS DWORD = 0, BYVAL cInitial AS LONG = 0, BYVAL cGrow AS LONG = 0)
   m_hInst = GetModuleHandle(NULL)
   ' // Create and initialize the image list
   IF dwFlags = 0 THEN dwFlags = ILC_COLOR OR ILC_MASK
   IF cInitial = 0 THEN cInitial = 16
   m_hImageList = ImageList_Create(cx, cy, dwFlags, cInitial, cGrow)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CImageList destructor
' ========================================================================================
PRIVATE DESTRUCTOR CImageList
   ' // Destroy the image list
   IF m_hImageList THEN ImageList_Destroy(m_hImageList)
END DESTRUCTOR
' ========================================================================================

' =====================================================================================
' Returns the handle of the image list
' =====================================================================================
PROPERTY CImageList.hImageList () AS ..HIMAGELIST
   PROPERTY = m_hImageList
END PROPERTY
' =====================================================================================

' =====================================================================================
' Loads an image from file or resource.
' - wszImage  = [in] Path to file that contains the image or images (the number of
'               images is inferred from the width of the bitmap), or name of the bitmap,
'               icon or cursor in the resource file (.RES). If the image resource uses
'               an integral identifier, wszImage should begin with a number symbol (#)
'               followed by the identifier in an ASCII format, e.g., "#998". Otherwise,
'               use the text identifier name for the image.
' - dwType    = [in] Specifies the type of image to be loaded(IMAGE_BITMAP, IMAGE_ICON, or IMAGE_CURSOR).
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' - bShared   = [in, opt] Shares the image handle if the image is loaded multiple times.
' If the function succeeds, the return value is the handle of the newly loaded image.
' If the function fails, the return value is NULL.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadImage (BYREF wszImage AS WSTRING, BYVAL dwType AS DWORD, BYVAL cxDesired AS LONG = 0, _
BYVAL cyDesired AS LONG = 0, BYVAL bShared AS BOOLEAN = FALSE) AS ..HIMAGELIST
   DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD, dwFlag AS DWORD
   IF bShared THEN dwFlag = LR_SHARED
   IF LEFT(wszImage, 1) = "#" THEN
      wID = VAL(MID(wszImage, 2))
      dwID = MAKELONG(wID, 0)
      hImage = ..LoadImageW(m_hInst, MAKEINTRESOURCEW(dwID), dwType, cxDesired, cyDesired, LR_DEFAULTCOLOR OR dwFlag)
   ELSEIF INSTR(wszImage, ".") THEN
      hImage = ..LoadImageW(NULL, @wszImage, dwType, cxDesired, cyDesired, LR_LOADFROMFILE OR LR_DEFAULTCOLOR)
   ELSE
      hImage = ..LoadImageW(m_hInst, @wszImage, dwType, cxDesired, cyDesired, LR_DEFAULTCOLOR)
   END IF
   RETURN hImage
END FUNCTION
' =====================================================================================

' =====================================================================================
' Loads a bitmap from file or resource and adds it to the image list.
' - wszImage  = [in] Path to file that contains the image or images (the number of
'               images is inferred from the width of the bitmap), or name of the bitmap
'               in the resource file (.RES). If the image resource uses an integral
'               identifier, strImage should begin with a number symbol (#) followed by
'               the identifier in an ASCII format, e.g., "#998". Otherwise, use the
'               text identifier name for the image.
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' - wszMask   = [in, opt] Path to file that contains the mask. If no mask is used with
'               the image list, this parameter is ignored.
' Returns the index of the first new image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadBitmap (BYREF wszImage AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYREF wszMask AS WSTRING = "") AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM hImage AS HANDLE, hMask AS HBITMAP
   hImage = this.LoadImage(wszImage, IMAGE_BITMAP, cxDesired, cyDesired)
   IF LEN(wszMask) = 0 THEN hMask = CAsT(HBITMAP, this.LoadImage(wszMask, IMAGE_BITMAP, cxDesired, cyDesired))
   IF hImage THEN
      FUNCTION = ImageList_Add(m_hImageList, CAST(HBITMAP, hImage), hMask)
      DeleteObject hImage
   END IF
   IF hMask THEN DeleteObject hMask
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a bitmap to the image list.
' - hImage = [in] Handle to the bitmap that contains the image or images. The number of
'            images is inferred from the width of the bitmap.
' - hMask  = [in, opt] Handle to the bitmap that contains the mask. If no mask is used
'            with the image list, this parameter is ignored.
' Returns the index of the first new image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.AddBitmap (BYVAL hImage AS HBITMAP, BYVAL hMask AS HBITMAP = NULL) AS LONG
   FUNCTION = -1   ' // Assume error for now
   IF hImage = NULL THEN EXIT FUNCTION
   FUNCTION = ImageList_Add(m_hImageList, hImage, hMask)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds an icon to the image list.
' - hImage = [in] Handle to the icon or cursor that contains the bitmap and mask for
'            the new image.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.AddIcon (BYVAL hImage AS HICON) AS LONG
   FUNCTION = -1   ' // Assume error for now
   IF hImage = NULL THEN EXIT FUNCTION
   IF m_hImageList THEN ImageList_ReplaceIcon(m_hImageList, -1, hImage)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Loads an icon from file or resource and adds it to the image list.
' - wszImage  = [in] Path to file that contains the icon, or name of the icon in the
'               resource file (.RES). If the image resource uses an integral identifier,
'               strImage should begin with a number symbol (#) followed by the identifier
'               in an ASCII format, e.g., "#998". Otherwise, use the text identifier
'               name for the image.
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadIcon (BYREF wszImage AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM hImage AS HICON = CAST(HICON, this.LoadImage(wszImage, IMAGE_ICON, cxDesired, cyDesired))
   IF hImage THEN
      IF m_hImageList THEN FUNCTION = ImageList_ReplaceIcon(m_hImageList, -1, hImage)
      DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Loads an icon from a resource file and adds it to the image list.
' - wID       = [in] Identifier of the icon in the resource file.
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' - bShared   = [in, opt] Shares the image handle if the image is loaded multiple times.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadResIcon (BYVAL wID AS WORD, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM dwID AS DWORD = MAKELONG(wID, 0)
   DIM dwFlag AS DWORD
   IF bShared THEN dwFlag = LR_SHARED
   DIM hImage AS HICON = LoadImageW(m_hInst, MAKEINTRESOURCEW(dwID), IMAGE_ICON, cxDesired, cyDesired, LR_DEFAULTCOLOR OR dwFlag)
   IF hImage THEN
      IF m_hImageList THEN FUNCTION = ImageList_ReplaceIcon(m_hImageList, -1, hImage)
      IF bShared = 0 THEN DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Loads a bitmap from a resource file and adds it to the image list.
' - wID       = [in] Identifier of the bitmap in the resource file.
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' - hMask     = [in, opt] Handle to the bitmap that contains the mask. If no mask is used
'               with the image list, this parameter is ignored.
' - bShared   = [in, opt] Shares the image handle if the image is loaded multiple times.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadResBitmap (BYVAL wID AS WORD, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL hMask AS HBITMAP, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM dwID AS DWORD = MAKELONG(wID, 0)
   DIM dwFlag AS DWORD
   IF bShared THEN dwFlag = LR_SHARED
   DIM hImage AS HBITMAP = LoadImageW(m_hInst, MAKEINTRESOURCEW(dwID), IMAGE_BITMAP, cxDesired, cyDesired, LR_DEFAULTCOLOR OR dwFlag)
   IF hImage THEN
      IF m_hImageList THEN FUNCTION = ImageList_Add(m_hImageList, hImage, hMask)
      IF bShared = 0 THEN DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================


' =====================================================================================
' Loads an icon from a resource file.
' - wszName   = [in] Name of the icon in the resource file.
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' - bShared   = [in, opt] Shares the image handle if the image is loaded multiple times.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadNamedResIcon (BYREF wszName AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM dwFlag AS DWORD
   IF bShared THEN dwFlag = LR_SHARED
   DIM hImage AS HICON = LoadImageW(m_hInst, wszName, IMAGE_ICON, cxDesired, cyDesired, LR_DEFAULTCOLOR OR dwFlag)
   IF hImage THEN
      IF m_hImageList THEN FUNCTION = ImageList_ReplaceIcon(m_hImageList, -1, hImage)
      IF bShared = 0 THEN DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Loads a bitmap from a resource file.
' - wszName   = [in] Name of the bitmap in the resource file.
' - cxDesired = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired = [in, opt] The desired height, in pixels, of the icon.
' - hMask     = [in, opt] Handle to the bitmap that contains the mask. If no mask is used
'               with the image list, this parameter is ignored.
' - bShared   = [in, opt] Shares the image handle if the image is loaded multiple times.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadNamedResBitmap (BYREF wszName AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL hMask AS HBITMAP, BYVAL bShared AS BOOLEAN = FALSE) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM dwFlag AS DWORD
   IF bShared THEN dwFlag = LR_SHARED
   DIM hImage AS HBITMAP = LoadImageW(m_hInst, wszName, IMAGE_BITMAP, cxDesired, cyDesired, LR_DEFAULTCOLOR OR dwFlag)
   IF hImage THEN
      IF m_hImageList THEN FUNCTION = ImageList_Add(m_hImageList, hImage, hMask)
      IF bShared = 0 THEN DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds an image or images to an image list, generating a mask from the specified bitmap.
' - hImage      = [in] Handle to the bitmap that contains the image or images. The number of
'                 images is inferred from the width of the bitmap.
' - dwMaskColor = [in, opt] Color used to generate the mask.
' Returns the index of the first new image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.AddMasked (BYVAL hImage AS HBITMAP, BYVAL dwMaskColor AS COLORREF = 0) AS LONG
   FUNCTION = -1   ' // Assume error for now
   IF dwMaskColor = 0 THEN dwMaskColor = RGB(255, 0 ,255)
   IF hImage THEN FUNCTION = ImageList_AddMasked(m_hImageList, hImage, dwMaskColor)
END FUNCTION
' =====================================================================================


' =====================================================================================
' Loads an image or images to an image list, generating a mask from the specified bitmap.
' - wszImage    = [in] Path to the bitmap that contains the image or images (the number
'                 of images is inferred from the width of the bitmap). If the image
'                 resource uses an integral identifier, strImage should begin with a
'                 number symbol (#) followed by the identifier in an ASCII format, e.g.,
'                 "#998". Otherwise, use the text identifier name for the image.
' - cxDesired   = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired   = [in, opt] The desired height, in pixels, of the icon.
' - dwMaskColor = [in] Color used to generate the mask.
' Returns the index of the first new image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadMasked (BYREF wszImage AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL dwMaskColor AS COLORREF = 0) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM hImage AS HBITMAP = CAST(HBITMAP, this.LoadImage(wszImage, IMAGE_BITMAP, cxDesired, cyDesired))
   IF hImage THEN
      IF dwMaskColor = 0 THEN dwMaskColor = RGB(255, 0 ,255)
      FUNCTION = ImageList_AddMasked(m_hImageList, hImage, dwMaskColor)
      DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Loads an image or images to an image list, generating a mask from the specified mask color.
' - wID         = [in] Identifier of the bitmap in the resource file.
' - cxDesired   = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired   = [in, opt] The desired height, in pixels, of the icon.
' - dwMaskColor = [in, opt] Color used to generate the mask.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadResBitmapMasked (BYVAL wID AS WORD, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL dwMaskColor AS COLORREF) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM dwID AS DWORD = MAKELONG(wID, 0)
   DIM hImage AS HBITMAP = LoadImageW(m_hInst, MAKEINTRESOURCEW(dwID), IMAGE_BITMAP, cxDesired, cyDesired, LR_DEFAULTCOLOR)
   IF hImage THEN
      IF dwMaskColor = 0 THEN dwMaskColor = RGB(255, 0 ,255)
      IF m_hImageList THEN FUNCTION = ImageList_AddMasked(m_hImageList, hImage, dwMaskColor)
      DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================


' =====================================================================================
' Loads a bitmap from a resource file.
' - wszName     = [in] Name of the bitmap in the resource file.
' - cxDesired   = [in, opt] The desired width, in pixels, of the icon.
' - cyDesired   = [in, opt] The desired height, in pixels, of the icon.
' - dwMaskColor = [in, opt] Color used to generate the mask.
' Returns the index of the image if successful, or -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.LoadNamedResBitmapMasked (BYREF wszName AS WSTRING, BYVAL cxDesired AS LONG = 0, BYVAL cyDesired AS LONG = 0, BYVAL dwMaskColor AS COLORREF) AS LONG
   FUNCTION = -1   ' // Assume error for now
   DIM hImage AS HBITMAP = LoadImageW(m_hInst, @wszName, IMAGE_BITMAP, cxDesired, cyDesired, LR_DEFAULTCOLOR)
   IF hImage THEN
      IF dwMaskColor = 0 THEN dwMaskColor = RGB(255, 0 ,255)
      IF m_hImageList THEN FUNCTION = ImageList_AddMasked(m_hImageList, hImage, dwMaskColor)
      DeleteObject hImage
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a specified image to the list of images to be used as overlay masks. An image
' list can have up to four overlay masks in version 4.70 and earlier and up to 15 in
' version 4.71. This method assigns an overlay mask index to the specified image.
' - iImage   = The zero-based index of an image in the image list. This index
'              identifies the image to use as an overlay mask.
' - iOverlay = The one-based index of the overlay mask.
' =====================================================================================
PRIVATE SUB CImageList.OverlayImage (BYVAL iImage AS LONG, BYVAL iOverlay AS LONG)
   IF (iOverlay >= 1) AND (iOverlay <= 15) THEN
      IF ImageList_SetOverlayImage(m_hImageList, iImage, iOverlay) <> E_INVALIDARG THEN
         ' // Note: the image index is increased by one to be able to distinguish when a slot has been initialized.
         m_Overlays(iOverlay) = iImage + 1
     END IF
   END IF
END SUB
' =====================================================================================

' =====================================================================================
' Returns the overlay index for the specified image index.
' iImage = The zero-based index of an image in image list that is used as an overlay mask.
' =====================================================================================
PRIVATE FUNCTION CImageList.OverlayIndex (BYVAL iImage AS LONG) AS LONG
   DIM AS LONG iItem, iOverlay
   iImage += 1
   iItem = 1
   WHILE iOverlay <= 15
      IF m_Overlays(iItem) = iImage THEN
         iOverlay = iItem
         EXIT WHILE
      END IF
      iItem += 1
   WEND
   FUNCTION = iOverlay
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a duplicate of the image list.
' =====================================================================================
PRIVATE FUNCTION CImageList.Duplicate () AS ..HIMAGELIST
   RETURN ImageList_Duplicate(m_hImageList)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Copies images within the image list.
' - iDst   = The zero-based index of the image to be used as the destination of the copy operation.
' - iSrc   = The zero-based index of the image to be used as the source of the copy operation.
' - dwFlag = The bit flag value that specifies the type of copy operation to be made.
'   This parameter can be one of the following values:
'   ILCF_MOVE = The source image is copied to the destination image's index. This
'               operation results in multiple instances of a given image.
'   ILCF_SWAP = The source and destination images exchange positions within the image list.
' Returns nonzero if successful, or zero otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.Copy (BYVAL iDst AS LONG, BYVAL iSrc AS LONG, BYVAL dwFlag AS DWORD) AS BOOLEAN
   RETURN ImageList_Copy(m_hImageList, iDst, m_hImageList, iSrc, dwFlag)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Removes an image from an image list.
' idx = The index of the image to remove. If this parameter is -1, then all images are removed.
' Returns nonzero if successful, or zero otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.Remove (BYVAL idx AS LONG) AS BOOLEAN
   RETURN ImageList_Remove(m_hImageList, idx)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of images in an image list. All images in an image list have
' the same dimensions.
' cx = The width, in pixels, of each image.
' cy = The height, in pixels, of each image.
' Returns nonzero if successful, or zero otherwise.
' =====================================================================================
PRIVATE FUNCTION CImageList.GetIconSize (BYREF cx AS LONG, BYREF cy AS LONG) AS BOOLEAN
   IF VARPTR(cx) = 0 OR VARPTR(cy) = 0 THEN RETURN FALSE
   RETURN ImageList_GetIconSize(m_hImageList, @cx, @cy)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the number of images in an image list.
' =====================================================================================
PRIVATE PROPERTY CImageList.ImageCount () AS LONG
   RETURN ImageList_GetImageCount(m_hImageList)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Retrieves the current background color for an image list.
' =====================================================================================
PRIVATE PROPERTY CImageList.BackgroundColor () AS COLORREF
   RETURN ImageList_GetBkColor(m_hImageList)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the background color for an image list. This function only works if you add an
' icon or use ImageList_AddMasked with a black and white bitmap. Without a mask, the
' entire image is drawn; hence the background color is not visible.
' =====================================================================================
PRIVATE PROPERTY CImageList.BackgroundColor (BYVAL clrBk AS COLORREF)
   ImageList_SetBkColor(m_hImageList, clrBk)
END PROPERTY
' =====================================================================================

END NAMESPACE
