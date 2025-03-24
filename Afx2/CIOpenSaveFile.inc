' ########################################################################################
' Microsoft Windows
' File: COpenSaveFile.inc
' Contents: Open/Save dialogs
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx2/AfxWin2.inc"
USING Afx2

' ========================================================================================
' Macro for debug
' To allow debugging, define _COPENFILE_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _COPENFILE_DEBUG_
   #define _COPENFILE_DEBUG_ 0
#ENDIF
#ifndef _COPENFILE_DP_
   #define _COPENFILE_DP_ 1
   #MACRO COPENFILE_DP(st)
      #IF (_COPENFILE_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE Afx2

' ========================================================================================
' IDisplayOpenFile class
' ========================================================================================
TYPE IDisplayOpenFile

Private:
   m_IFileOpenDialog AS IFileOpenDialog PTR    ' // Pointer to the IFileOpenDialog interface
   m_hwndOwner AS HWND                         ' // Handle of the owner of the dialog
   m_xPos AS LONG                              ' // Horizontal position, in pixels, relative to the parent window.
   m_yPos AS LONG                              ' // Vertical position, in pixels, relative to the parent window.
   m_dwsTitle AS DWSTRING                      ' // The title to be displayed in the title bar of the dialog box
   m_dwsFolder AS DWSTRING                     ' // The name of the initial file directory to be displayed
   m_rgFilterSpec (ANY) AS COMDLG_FILTERSPEC   ' // Array of COMDLG_FILTERSPEC structures
   m_startFile AS DWSTRING                     ' // File name to be used as the initial file selection
   m_defExt AS DWSTRING                        ' // Default extension to be appended to the selected file name if the user does not enter it
   m_Flags AS DWORD                            ' // Flags
   m_bCenter AS BOOLEAN                        ' // Centered dialog

Public:
   DECLARE CONSTRUCTOR (BYVAL xPos AS LONG, BYVAL yPos AS LONG, BYVAL bCenter AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowOpen (BYVAL hwndOwner AS HWND) AS HRESULT
   DECLARE SUB AddFileType (BYREF dwsName AS DWSTRING, BYREF dwsSpec AS DWSTRING)
   DECLARE FUNCTION SetFileTypes () AS HRESULT
   DECLARE FUNCTION SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS CONST COMDLG_FILTERSPEC PTR) AS HRESULT
   DECLARE FUNCTION SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   DECLARE FUNCTION GetFileTypeIndex () AS UINT
   DECLARE FUNCTION SetOptions (BYVAL fos AS FILEOPENDIALOGOPTIONS) AS HRESULT
   DECLARE FUNCTION GetOptions () AS FILEOPENDIALOGOPTIONS
   DECLARE FUNCTION SetDefaultFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   DECLARE FUNCTION SetFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   DECLARE FUNCTION GetFolder () AS IShellItem PTR
   DECLARE FUNCTION GetCurrentSelection () AS IShellItem PTR
   DECLARE FUNCTION SetFileName (BYVAL pwszName AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION GetFileName () AS DWSTRING
   DECLARE FUNCTION SetTitle (BYVAL pwszTitle AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetOkButtonLabel (BYVAL pwszText AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetFileNameLabel (BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION GetResult () AS IShellItem PTR
   DECLARE FUNCTION AddPlace (BYVAL psi AS IShellItem PTR, BYVAL fdap AS FDAP) AS HRESULT
   DECLARE FUNCTION SetDefaultExtension (BYVAL pwszDefaultExtension AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION Close (BYVAL hr AS HRESULT) AS HRESULT
   DECLARE FUNCTION SetClientGuid (BYVAL guid AS const GUID const PTR) AS HRESULT
   DECLARE FUNCTION ClearClientData () AS HRESULT
   DECLARE FUNCTION GetResults () AS IShellItemArray PTR
   DECLARE FUNCTION GetSelectedItems () AS IShellItemArray PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR IDisplayOpenFile (BYVAL xPos AS LONG, BYVAL yPos AS LONG, BYVAL bCenter AS BOOLEAN = FALSE)
   m_xPos = xPos
   m_yPos = yPos
   m_bCenter = bCenter
   CoInitialize NULL
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR IDisplayOpenFile
   m_IFileOpenDialog->lpvtbl->Release(m_IFileOpenDialog)
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Display he dialog
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.ShowOpen (BYVAL hwndOwner AS HWND) AS HRESULT
   ' // Before displaying the dialog get its coordinates
   DIM pOleWindow AS IOleWindow PTR
   DIM hOleWindow AS HWND
   ' // Get a pointer to the IOleWindow interface
   m_IFileOpenDialog->lpvtbl->QueryInterface(m_IFileOpenDialog, @IID_IOleWindow, @pOleWindow)
   ' // Get the window handle of the dialog
   IF @pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @hOleWindow)
   IF hOleWindow THEN
      DIM rc AS RECT
      ' // Get the dimensions of the window
      GetWindowRect(hOleWindow, @rc)
      ' // Convert the coordinates to be relative to the parent
      MapWindowPoints(HWND_DESKTOP, GetParent(hwndOwner), CAST(POINT PTR, @rc), 2)
      DIM xPos AS LONG = rc.Left
      DIM yPos AS LONG = rc.Top
      IF (xPos = 0 ANDALSO yPos = 0) THEN
         ' // Center the open dialog on the desktop window
'         AfxCenterWindow (hOleWindow, hwndOwner)

         ' // Center the open dialog on the parent window
         DIM AS RECT rc, rcDlg, rcOwner
         GetWindowRect(hwndOwner, @rcOwner)
         GetWindowRect(hOleWindow, @rcDlg)
         CopyRect(@rc, @rcOwner)
         ' // Offset the owner and dialog box rectangles so that right and bottom 
         ' // values represent the width and height, and then offset the owner again 
         ' // to discard space taken up by the dialog box. 
         OffsetRect(@rcDlg, -rcDlg.left, -rcDlg.top)
         OffsetRect(@rc, -rc.left, -rc.top)
         OffsetRect(@rc, -rcDlg.right, -rcDlg.bottom)
         ' // The new position is the sum of half the remaining space and the owner's 
         ' // original position. 
         SetWindowPos(hOleWindow, HWND_TOP, rcOwner.left + (rc.right / 2), _
                      rcOwner.top + (rc.bottom / 2), 0, 0, SWP_NOSIZE)

      ELSE
'         DIM rc AS RECT
'         ' // Get the dimensions of the bounding rectangle of the parent window
'         GetWindowRect hwndOwner, @rc
'         ' // Move the open dialog to the specified location, relative to the parent window
'         SetWindowPos hOleWindow, 0, rc.Left + xPos, rc.Top + yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER

         DIM rc AS RECT
         DIM AS ANY PTR pLib = DyLibLoad("Dwmapi.DLL")
         IF pLib  THEN
            DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS ANY PTR, BYVAL cbAttribute AS DWORD) AS HRESULT
            pProc = DyLibSymbol(pLib, "DwmGetWindowAttribute")
            ' DWMWA_EXTENDED_FRAME_BOUNDS = 9
            IF pProc THEN pProc(hwndOwner, 9, @rc, SIZEOF(rc))
            DyLibFree(pLib)
         END IF
         ' // Move the open dialog to the specified location, relative to the parent window
         SetWindowPos hOleWindow, 0, rc.Left + xPos, rc.Top + yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER

      END IF   
   END IF
   ' // Display the dialog
   RETURN m_IFileOpenDialog->lpvtbl->Show(m_IFileOpenDialog, hwndOwner)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' ========================================================================================
'PRIVATE FUNCTION IDisplayOpenFile.SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS CONST COMDLG_FILTERSPEC PTR) AS HRESULT
'   RETURN m_IFileOpenDialog->lpvtbl->Show(m_IFileOpenDialog, cFileTypes, @rgFilterSpec(1))
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the file type that appears as selected in the dialog.
' This method must be called before the dialog is showing.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetFileTypeIndex(m_IFileOpenDialog, iFileType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetFileTypeIndex () AS UINT
   DIM iFileType AS UINT
   m_IFileOpenDialog->lpvtbl->GetFileTypeIndex(m_IFileOpenDialog, @iFileType)
   RETURN iFileType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an event handler that listens for events coming from the dialog.
' ========================================================================================
'PRIVATE FUNCTION IDisplayOpenFile.Advise (BYVAL pfde as IFileDialogEvents PTR) AS DWORD
'   DIM dwCookie AS DWORD
'   SetLastError(m_IFileOpenDialog->lpvtbl->Advise(m_IFileOpenDialog, @dwCookie))
'   RETURN dwCookie
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an event handler that was attached through the Advise method.
' ========================================================================================
'PRIVATE FUNCTION IDisplayOpenFile.Unadvise (BYVAL pfde as IFileDialogEvents PTR, BYVAL dwCookie AS DWORD) AS HRESULT
'   m_IFileOpenDialog->lpvtbl->Unadvise(m_IFileOpenDialog, dwCookie)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets flags to control the behavior of the dialog.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetOptions (BYVAL fos AS FILEOPENDIALOGOPTIONS) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetOptions(m_IFileOpenDialog, fos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current flags that are set to control dialog behavior.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetOptions () AS FILEOPENDIALOGOPTIONS
   DIM fos AS FILEOPENDIALOGOPTIONS
   SetLastError(m_IFileOpenDialog->lpvtbl->GetOptions(m_IFileOpenDialog, @fos))
   RETURN fos
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetDefaultFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetDefaultFolder(m_IFileOpenDialog, psi)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a folder that is always selected when the dialog is opened, regardless of previous user action.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetFolder(m_IFileOpenDialog, psi)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets either the folder currently selected in the dialog, or, if the dialog is not currently
' displayed, the folder that is to be selected when the dialog is opened.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetFolder () AS IShellItem PTR
   DIM psi AS IShellItem PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetFolder(m_IFileOpenDialog, @psi))
   RETURN psi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetCurrentSelection () AS IShellItem PTR
   DIM psi AS IShellItem PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetCurrentSelection(m_IFileOpenDialog, @psi))
   RETURN psi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetFileName (BYVAL pwszName AS WSTRING PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetFileName(m_IFileOpenDialog, pwszName)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text currently entered in the dialog's File name edit box.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetFileName () AS DWSTRING
   DIM pwszName AS WSTRING PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetFileName(m_IFileOpenDialog, @pwszName))
   RETURN pwszName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the title of the dialog.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetTitle (BYVAL pwszTitle AS WSTRING PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetTitle(m_IFileOpenDialog, pwszTitle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the Open or Save button.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetOkButtonLabel (BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetOkButtonLabel(m_IFileOpenDialog, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of the label next to the file name edit box.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetFileNameLabel (BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetFileNameLabel(m_IFileOpenDialog, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the choice that the user made in the dialog.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetResult () AS IShellItem PTR
   DIM psi AS IShellItem PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetResult(m_IFileOpenDialog, @psi))
   RETURN psi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a folder to the list of places available for the user to open or save items.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.AddPlace (BYVAL psi AS IShellItem PTR, BYVAL fdap AS FDAP) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->AddPlace(m_IFileOpenDialog, psi, fdap)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default extension to be added to file names.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetDefaultExtension (BYVAL pwszDefaultExtension AS WSTRING PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetDefaultExtension(m_IFileOpenDialog, pwszDefaultExtension)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes the dialog.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.Close (BYVAL hr AS HRESULT) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->Close(m_IFileOpenDialog, hr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes the dialog.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetClientGuid (BYVAL guid AS const GUID const PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetClientGuid(m_IFileOpenDialog, guid)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Instructs the dialog to clear all persisted state information.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.ClearClientData () AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->ClearClientData(m_IFileOpenDialog)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deprecated. SetFilter is no longer available for use as of Windows 7.
' ========================================================================================
'PRIVATE FUNCTION IDisplayOpenFile.SetFilter (BYVAL pFilter AS IShellItemFilter PTR) AS HRESULT
'   RETURN m_IFileOpenDialog->lpvtbl->SetFilter(m_IFileOpenDialog, pFilter)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetResults () AS IShellItemArray PTR
   DIM penum AS IShellItemArray PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetResults(m_IFileOpenDialog, @penum))
   RETURN penum
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.GetSelectedItems () AS IShellItemArray PTR
   DIM psai AS IShellItemArray PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetSelectedItems(m_IFileOpenDialog, @psai))
   RETURN psai
END FUNCTION
' ========================================================================================

' =====================================================================================
' Adds a file type and pattern to the table.
' =====================================================================================
PRIVATE SUB IDisplayOpenFile.AddFileType (BYREF dwsName AS DWSTRING, BYREF dwsSpec AS DWSTRING)
   REDIM PRESERVE m_rgFilterSpec(UBOUND(m_rgFilterSpec) + 1) AS COMDLG_FILTERSPEC
   m_rgFilterSpec(UBOUND(m_rgFilterSpec)).pszName = *dwsName
   m_rgFilterSpec(UBOUND(m_rgFilterSpec)).pszSpec = *dwsSpec
END SUB
' =====================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' This overloaded functions uses the m_rgFilterSpec array to set the file types.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetFileTypes () AS HRESULT
   DIM cFileTypes AS LONG = UBOUND(m_rgFilterSpec) - LBOUND(m_rgFilterSpec) + 1
   IF cFileTYpes = 0 THEN RETURN E_UNEXPECTED
   RETURN m_IFileOpenDialog->lpvtbl->SetFileTypes(m_IFileOpenDialog, cFileTypes, @m_rgFilterSpec(0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' ========================================================================================
PRIVATE FUNCTION IDisplayOpenFile.SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS CONST COMDLG_FILTERSPEC PTR) AS HRESULT
   RETURN m_IFileOpenDialog->lpvtbl->SetFileTypes(m_IFileOpenDialog, cFileTypes, @m_rgFilterSpec(0))
END FUNCTION
' ========================================================================================

' ########################################################################################

' ========================================================================================
' DisplayOpenFile class
' ========================================================================================
TYPE DisplayOpenFile

Private:
   m_hParent AS HWND                       ' // Handle of the parent window or dialog
   m_xPos AS LONG                          ' // Horizontal position, in pixels, relative to the parent window
   m_yPos AS LONG                          ' // Vertical position, in pixels, relative to the parent window
   m_Flags AS DWORD                        ' // The style attributes of the dialog
   m_FlagsEx AS DWORD                      ' // Extra style attributes of the dialog
   m_Index AS UINT                         ' // File type that appears as selected in the dialog
   m_wszTitle AS WSTRING * MAX_PATH        ' // The title to be displayed in the title bar of the dialog box
   m_wszInitiaDir AS WSTRING * MAX_PATH    ' // The name of the initial file directory to be displayed
   m_wszFileName AS WSTRING * MAX_PATH     ' // The file name to be used as the initial file selection
   m_wszDefaultExt AS WSTRING * MAX_PATH   ' // Default extension to be appended to the selected file name if the user does not enter it
   m_dwsFilter AS DWSTRING                 ' // Filter (pairs of strings separated with "|")
   m_FilterIndex AS DWORD                  ' // The file type that appears as selected in the dialog

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowOpen (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
   DECLARE PROPERTY Title () AS DWSTRING
   DECLARE PROPERTY Title (BYREF wszTitle AS WSTRING)
   DECLARE PROPERTY Flags () AS DWORD
   DECLARE PROPERTY Flags (BYVAL dwFlags AS DWORD)
   DECLARE PROPERTY FlagsEx () AS DWORD
   DECLARE PROPERTY FlagsEx (BYVAL dwFlagsEx AS DWORD)
   DECLARE PROPERTY DefaultExt () AS DWSTRING
   DECLARE PROPERTY DefaultExt (BYREF wszDefaultExt AS WSTRING)
   DECLARE PROPERTY InitialDir () AS DWSTRING
   DECLARE PROPERTY InitialDir (BYREF wszInitialDir AS WSTRING)
   DECLARE PROPERTY FilterIndex () AS DWORD
   DECLARE PROPERTY FilterIndex (BYVAL dwFilterIndex AS DWORD)
   DECLARE PROPERTY FileName () AS DWSTRING
   DECLARE PROPERTY FileName (BYREF wszFileName AS WSTRING)
   DECLARE PROPERTY Filter () AS DWSTRING
   DECLARE PROPERTY Filter (BYREF dwsFilter AS DWSTRING)

Private:
   DECLARE STATIC FUNCTION OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR DisplayOpenFile
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR DisplayOpenFile
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Open file dialog hook callback procedure
' ========================================================================================
FUNCTION DisplayOpenFile.OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR
   DIM pOFN AS OPENFILENAMEW PTR
   DIM AS LONG xPos, yPos
   ' // Process the WM_INITDIALOG message
   IF uMsg = WM_INITDIALOG THEN
      ' // Get the custom data
      pOFN = CAST(OPENFILENAMEW PTR, lParam)
      IF pOFN THEN
         xPos = CSHORT(LOWORD(pOFN->lCustData))
         yPos = CSHORT(HIWORD(pOFN->lCustData))
      END IF
      ' // Note: hwnd is the handle of the child dialog box.
      ' // Use GetParent to get the handle of the Open dialog box.
      DIM hOpenFileDlg AS .HWND = GetParent(hwnd)
      ' // Get the parent of the open file dialog.
      ' // It can be another window or dialog, or the desktop.
      DIM hParent AS .HWND = GetParent(hOpenFileDlg)
      IF hParent = NULL THEN hParent = GetDesktopWindow
      ' // If the user has passed -1, -1 as the position
      IF (xPos = -1 ANDALSO yPos = -1) THEN
         ' // Center the open dialog on the parent window
         DIM AS RECT rc, rcDlg, rcOwner
         GetWindowRect(hParent, @rcOwner)
         GetWindowRect(hOpenFileDlg, @rcDlg)
         CopyRect(@rc, @rcOwner)
         ' // Offset the owner and dialog box rectangles so that right and bottom 
         ' // values represent the width and height, and then offset the owner again 
         ' // to discard space taken up by the dialog box. 
         OffsetRect(@rcDlg, -rcDlg.left, -rcDlg.top)
         OffsetRect(@rc, -rc.left, -rc.top)
         OffsetRect(@rc, -rcDlg.right, -rcDlg.bottom)
         ' // The new position is the sum of half the remaining space and the owner's 
         ' // original position. 
         SetWindowPos(hOpenFileDlg, HWND_TOP, rcOwner.left + (rc.right / 2), _
                      rcOwner.top + (rc.bottom / 2), 0, 0, SWP_NOSIZE)
      ELSE
'         ' // Get the dimensions of the bounding rectangle of the parent window
'         DIM rc AS RECT
'         GetWindowRect hParent, @rc
'         ' // Move the open dialog to the specified location, relative to the parent window
'         SetWindowPos hOpenFileDlg, 0, rc.Left + x, rc.Top + y, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
         ' // If the windows has been shown, GetWindowRect takes into account the drop shadows.
         ' // We must use the DwmGetWindowAttribute function to get the dimensions of the windows without the drop shadows;
         ' // otherwise, the open file dialog will be displaced to the left.
         DIM rc AS RECT
         DIM AS ANY PTR pLib = DyLibLoad("Dwmapi.DLL")
         IF pLib  THEN
            DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS ANY PTR, BYVAL cbAttribute AS DWORD) AS HRESULT
            pProc = DyLibSymbol(pLib, "DwmGetWindowAttribute")
            ' DWMWA_EXTENDED_FRAME_BOUNDS = 9
            IF pProc THEN pProc(hParent, 9, @rc, SIZEOF(rc))
            DyLibFree(pLib)
         END IF
         ' // Move the open dialog to the specified location, relative to the parent window
         SetWindowPos hOpenFileDlg, 0, rc.Left + xPos, rc.Top + yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
      END IF   
   END IF
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Display the dialog.
' Usage example:
'   DIM pOFN AS DisplayOpenFile
'   pOFN.FileName = "*.*"
'   pOFN.InitialDir = CURDIR
'   pOFN.Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
'   pOFN.DefaultExt = "BAS"
'   pOFN.FilterIndex = 1
'   'pOFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
'   pOFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_ALLOWMULTISELECT
'   DIM dwsFile AS DWSTRING = pOFN.ShowOpen(hDlg, 20, 20, TRUE)
'   --or-- DIM dwsFile AS DWSTRING = pOFN.ShowOpen(hDlg, -1, -1, TRUE) to center the dialog
'   PRINT dwsFile
' ========================================================================================
PRIVATE FUNCTION DisplayOpenFile.ShowOpen (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING

   ' // Filter is a sequence of WSTRINGs with a final (extra) double null terminator
   ' // The "|" characters are replaced with nulls
   DIM wszMarkers AS WSTRING * 4 = "||"
   IF RIGHT(m_dwsFilter, 1) <> "|" THEN wszMarkers += "|"
   m_dwsFilter = m_dwsFilter & wszMarkers
   DIM dwFilterStrSize AS DWORD = LEN(m_dwsFilter)
   ' // Replace markers("|") with nulls
   DIM pchar AS WCHAR PTR = *m_dwsFilter
   FOR i AS LONG = 0 TO LEN(m_dwsFilter) - 1
      IF pchar[i] = ASC("|") THEN pchar[i] = 0
   NEXT

   ' // If the initial directory has not been specified, assume the current directory
   IF LEN(m_wszInitiaDir) = 0 THEN m_wszInitiaDir = CURDIR
   ' // The size of the buffer must be at least MAX_PATH characters
   DIM dwBufLen AS DWORD = MAX_PATH   ' // Make room for at least one path
   IF (m_Flags AND OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT THEN dwBufLen = 32768  ' // 64 Kb buffer
   ' // Allocate the file name and a marker ("|") to be replaced with a null
   DIM dwsFileName AS DWSTRING = m_wszFileName & "|"
   ' // Store the position of the marker
   DIM cbPos AS LONG = LEN(dwsFileName) - 1
   ' // Allocate room for the buffer
   IF LEN(dwsFileName) < dwBufLen THEN dwsFileName += SPACE(dwBufLen - LEN(dwsFileName))
   DIM dwFileStrSize AS DWORD = LEN(dwsFileName)
   ' // The filename must be null terminated (replace the marker with a null)
   pchar = *dwsFileName
   pchar[cbPos] = 0

   ' // Fill the members of the OPENFILENAMEW structure
   DIM ofn AS OPENFILENAMEW
   ofn.lStructSize = SIZEOF(ofn)
   ofn.hwndOwner = hParent
   ofn.lpstrFilter = *m_dwsFilter
   IF m_FilterIndex < 1 THEN m_FilterIndex = 1
   ofn.nFilterIndex = m_FilterIndex
   ofn.lpstrFile = *dwsFileName
   ofn.nMaxFile = dwFileStrSize
   IF m_wszInitiaDir = "" THEN m_wszInitiaDir = CURDIR
   ofn.lpstrInitialDir = @m_wszInitiaDir
   IF LEN(m_wszTitle) THEN ofn.lpstrTitle = @m_wszTitle
   IF LEN(m_wszDefaultExt) THEN ofn.lpstrDefExt = @m_wszDefaultExt
   ofn.lCustData = MAKELONG(xPos, yPos)
   IF bHook THEN
      ' // Set a hook
      m_Flags = m_Flags OR OFN_EXPLORER OR OFN_ENABLEHOOK
      ofn.Flags = m_Flags
      ofn.lpfnHook = @DisplayOpenFile.OFNHookProc
   END IF
   ' // Call the open file dialog
   IF GetOpenFilenameW(@ofn) THEN
      IF (m_Flags AND OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT THEN
         pchar = *dwsFileName
         FOR i AS LONG = 0 TO dwFileStrSize - 1
            ' // If double null, exit
            IF pchar[i] = 0 AND pchar[i + 1] = 0 THEN EXIT FOR
            ' // Replace null with ","
            IF pchar[i] = 0 THEN pchar[i] = ASC(",")
         NEXT
      END IF
      ' // Trim trailing spaces
      dwsFileName = RTRIM(dwsFileName, CHR(32))
      IF RIGHT(dwsFileName, 1) = "," THEN dwsFileName = LEFT(dwsFileName, LEN(dwsFileName) - 1)
   ELSE
      ' // Buffer too small
      IF CommDlgExtendedError = FNERR_BUFFERTOOSMALL THEN
         dwBufLen = ASC(dwsFileName)   ' // needed buffer lenght
      END IF
      dwsFileName = ""
   END IF

   ' // Return the selected file or files
   RETURN dwsFileName

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.Title () AS DWSTRING
   RETURN m_wszTitle
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.Title (BYREF wszTitle AS WSTRING)
   m_wszTitle = wszTitle
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.Flags () AS DWORD
   RETURN m_Flags
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.FlagsEx () AS DWORD
   RETURN m_FlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.Flags (BYVAL dwFlags AS DWORD)
   m_Flags = dwFlags
END PROPERTY
' ========================================================================================
' ========================================================================================
' Currently, the only option is OFN_EX_NOPLACESBAR
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.FlagsEx (BYVAL dwFlagsEx AS DWORD)
   m_FlagsEx = dwFlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.DefaultExt () AS DWSTRING
   RETURN m_wszDefaultExt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.DefaultExt (BYREF wszDefaultExt AS WSTRING)
   m_wszDefaultExt = wszDefaultExt
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the folder used as a the initial directory
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.InitialDir () AS DWSTRING
   RETURN m_wszInitiaDir
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.InitialDir (BYREF wszInitialDir AS WSTRING)
   m_wszInitiaDir = wszInitialDir
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.FilterIndex () AS DWORD
   RETURN m_FilterIndex
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.FilterIndex (BYVAL dwFilterIndex AS DWORD)
   m_FilterIndex = dwFilterIndex
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.FileName () AS DWSTRING
   RETURN m_wszFileName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.FileName (BYREF wszFileName AS WSTRING)
   m_wszFileName = wszFileName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.Filter () AS DWSTRING
   RETURN m_dwsFilter
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY DisplayOpenFile.Filter (BYREF dwsFilter AS DWSTRING)
   m_dwsFilter = dwsFilter
END PROPERTY
' ========================================================================================


' ########################################################################################

' ========================================================================================
' DisplaySaveFile class
' ========================================================================================
TYPE DisplaySaveFile

Private:
   m_hParent AS HWND                       ' // Handle of the parent window or dialog
   m_xPos AS LONG                          ' // Horizontal position, in pixels, relative to the parent window
   m_yPos AS LONG                          ' // Vertical position, in pixels, relative to the parent window
   m_Flags AS DWORD                        ' // The style attributes of the dialog
   m_FlagsEx AS DWORD                      ' // Extra style attributes of the dialog
   m_Index AS UINT                         ' // File type that appears as selected in the dialog
   m_wszTitle AS WSTRING * MAX_PATH        ' // The title to be displayed in the title bar of the dialog box
   m_wszInitiaDir AS WSTRING * MAX_PATH    ' // The name of the initial file directory to be displayed
   m_wszFileName AS WSTRING * MAX_PATH     ' // The file name to be used as the initial file selection
   m_wszDefaultExt AS WSTRING * MAX_PATH   ' // Default extension to be appended to the selected file name if the user does not enter it
   m_dwsFilter AS DWSTRING                 ' // Filter (pairs of strings separated with "|")
   m_FilterIndex AS DWORD                  ' // The file type that appears as selected in the dialog

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowSave (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
   DECLARE PROPERTY Title () AS DWSTRING
   DECLARE PROPERTY Title (BYREF wszTitle AS WSTRING)
   DECLARE PROPERTY Flags () AS DWORD
   DECLARE PROPERTY Flags (BYVAL dwFlags AS DWORD)
   DECLARE PROPERTY FlagsEx () AS DWORD
   DECLARE PROPERTY FlagsEx (BYVAL dwFlagsEx AS DWORD)
   DECLARE PROPERTY DefaultExt () AS DWSTRING
   DECLARE PROPERTY DefaultExt (BYREF wszDefaultExt AS WSTRING)
   DECLARE PROPERTY InitialDir () AS DWSTRING
   DECLARE PROPERTY InitialDir (BYREF wszInitialDir AS WSTRING)
   DECLARE PROPERTY FilterIndex () AS DWORD
   DECLARE PROPERTY FilterIndex (BYVAL dwFilterIndex AS DWORD)
   DECLARE PROPERTY FileName () AS DWSTRING
   DECLARE PROPERTY FileName (BYREF wszFileName AS WSTRING)
   DECLARE PROPERTY Filter () AS DWSTRING
   DECLARE PROPERTY Filter (BYREF dwsFilter AS DWSTRING)

Private:
   DECLARE STATIC FUNCTION OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR DisplaySaveFile
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR DisplaySaveFile
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Open file dialog hook callback procedure
' ========================================================================================
FUNCTION DisplaySaveFile.OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR
   DIM pOFN AS OPENFILENAMEW PTR
   DIM AS LONG xPos, yPos
   ' // Process the WM_INITDIALOG message
   IF uMsg = WM_INITDIALOG THEN
      ' // Get the custom data
      pOFN = CAST(OPENFILENAMEW PTR, lParam)
      IF pOFN THEN
         xPos = CSHORT(LOWORD(pOFN->lCustData))
         yPos = CSHORT(HIWORD(pOFN->lCustData))
      END IF
      ' // Note: hwnd is the handle of the child dialog box.
      ' // Use GetParent to get the handle of the Open dialog box.
      DIM hOpenFileDlg AS .HWND = GetParent(hwnd)
      ' // Get the parent of the open file dialog.
      ' // It can be another window or dialog, or the desktop.
      DIM hParent AS .HWND = GetParent(hOpenFileDlg)
      IF hParent = NULL THEN hParent = GetDesktopWindow
      ' // If the user has passed -1, -1 as the position
      IF (xPos = -1 ANDALSO yPos = -1) THEN
         ' // Center the open dialog on the parent window
         DIM AS RECT rc, rcDlg, rcOwner
         GetWindowRect(hParent, @rcOwner)
         GetWindowRect(hOpenFileDlg, @rcDlg)
         CopyRect(@rc, @rcOwner)
         ' // Offset the owner and dialog box rectangles so that right and bottom 
         ' // values represent the width and height, and then offset the owner again 
         ' // to discard space taken up by the dialog box. 
         OffsetRect(@rcDlg, -rcDlg.left, -rcDlg.top)
         OffsetRect(@rc, -rc.left, -rc.top)
         OffsetRect(@rc, -rcDlg.right, -rcDlg.bottom)
         ' // The new position is the sum of half the remaining space and the owner's 
         ' // original position. 
         SetWindowPos(hOpenFileDlg, HWND_TOP, rcOwner.left + (rc.right / 2), _
                      rcOwner.top + (rc.bottom / 2), 0, 0, SWP_NOSIZE)
      ELSE
'         ' // Get the dimensions of the bounding rectangle of the parent window
'         DIM rc AS RECT
'         GetWindowRect hParent, @rc
'         ' // Move the open dialog to the specified location, relative to the parent window
'         SetWindowPos hOpenFileDlg, 0, rc.Left + x, rc.Top + y, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
         ' // If the windows has been shown, GetWindowRect takes into account the drop shadows.
         ' // We must use the DwmGetWindowAttribute function to get the dimensions of the windows without the drop shadows;
         ' // otherwise, the open file dialog will be displaced to the left.
         DIM rc AS RECT
         DIM AS ANY PTR pLib = DyLibLoad("Dwmapi.DLL")
         IF pLib  THEN
            DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS ANY PTR, BYVAL cbAttribute AS DWORD) AS HRESULT
            pProc = DyLibSymbol(pLib, "DwmGetWindowAttribute")
            ' DWMWA_EXTENDED_FRAME_BOUNDS = 9
            IF pProc THEN pProc(hParent, 9, @rc, SIZEOF(rc))
            DyLibFree(pLib)
         END IF
         ' // Move the open dialog to the specified location, relative to the parent window
         SetWindowPos hOpenFileDlg, 0, rc.Left + xPos, rc.Top + yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
      END IF   
   END IF
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Display the dialog.
' Usage example:
'   DIM pSFN AS DisplaySaveFile
'   pSFN.FileName = "*.*"
'   pSFN.InitialDir = CURDIR
'   pSFN.Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
'   pSFN.DefaultExt = "BAS"
'   pSFN.FilterIndex = 1
'   pSFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
'   DIM dwsFile AS DWSTRING = pSFN.ShowSave(hDlg, 20, 20, TRUE)
'   --or-- DIM dwsFile AS DWSTRING = pSFN.ShowSave(hDlg, -1, -1, TRUE) to center the dialog
'   PRINT dwsFile
' ========================================================================================
PRIVATE FUNCTION DisplaySaveFile.ShowSave (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING

   ' // Filter is a sequence of WSTRINGs with a final (extra) double null terminator
   ' // The "|" characters are replaced with nulls
   DIM wszMarkers AS WSTRING * 4 = "||"
   IF RIGHT(m_dwsFilter, 1) <> "|" THEN wszMarkers += "|"
   m_dwsFilter = m_dwsFilter & wszMarkers
   DIM dwFilterStrSize AS DWORD = LEN(m_dwsFilter)
   ' // Replace markers("|") with nulls
   DIM pchar AS WCHAR PTR = *m_dwsFilter
   FOR i AS LONG = 0 TO LEN(m_dwsFilter) - 1
      IF pchar[i] = ASC("|") THEN pchar[i] = 0
   NEXT

   ' // If the initial directory has not been specified, assume the current directory
   IF LEN(m_wszInitiaDir) = 0 THEN m_wszInitiaDir = CURDIR
   ' // Allocate the file name and a marker ("|") to be replaced with a null
   DIM dwsFileName AS DWSTRING = m_wszFileName & "|"
   ' // Store the position of the marker
   DIM cbPos AS LONG = LEN(dwsFileName) - 1
   ' // Allocate room for the buffer
   IF LEN(dwsFileName) < MAX_PATH THEN dwsFileName += SPACE(MAX_PATH - LEN(dwsFileName))
   DIM dwFileStrSize AS DWORD = LEN(dwsFileName)
   ' // The filename must be null terminated (replace the marker with a null)
   pchar = *dwsFileName
   pchar[cbPos] = 0

   ' // Fill the members of the OPENFILENAMEW structure
   DIM ofn AS OPENFILENAMEW
   ofn.lStructSize = SIZEOF(ofn)
   ofn.hwndOwner = hParent
   ofn.lpstrFilter = *m_dwsFilter
   IF m_FilterIndex < 1 THEN m_FilterIndex = 1
   ofn.nFilterIndex = m_FilterIndex
   ofn.lpstrFile = *dwsFileName
   ofn.nMaxFile = dwFileStrSize
   IF m_wszInitiaDir = "" THEN m_wszInitiaDir = CURDIR
   ofn.lpstrInitialDir = @m_wszInitiaDir
   IF LEN(m_wszTitle) THEN ofn.lpstrTitle = @m_wszTitle
   IF LEN(m_wszDefaultExt) THEN ofn.lpstrDefExt = @m_wszDefaultExt
   ofn.lCustData = MAKELONG(xPos, yPos)
   IF bHook THEN
      ' // Set a hook
      m_Flags = m_Flags OR OFN_EXPLORER OR OFN_ENABLEHOOK
      ofn.Flags = m_Flags
      ofn.lpfnHook = @DisplaySaveFile.OFNHookProc
   END IF

   ' // Call the save filename dialog
   IF GetSaveFilenameW(@ofn) = 0 THEN dwsFileName = ""

   ' // Return the selected file or files
   RETURN dwsFileName

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.Title () AS DWSTRING
   RETURN m_wszTitle
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.Title (BYREF wszTitle AS WSTRING)
   m_wszTitle = wszTitle
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.Flags () AS DWORD
   RETURN m_Flags
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.FlagsEx () AS DWORD
   RETURN m_FlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.Flags (BYVAL dwFlags AS DWORD)
   m_Flags = dwFlags
END PROPERTY
' ========================================================================================
' ========================================================================================
' Currently, the only option is OFN_EX_NOPLACESBAR
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.FlagsEx (BYVAL dwFlagsEx AS DWORD)
   m_FlagsEx = dwFlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.DefaultExt () AS DWSTRING
   RETURN m_wszDefaultExt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.DefaultExt (BYREF wszDefaultExt AS WSTRING)
   m_wszDefaultExt = wszDefaultExt
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the folder used as a the initial directory
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.InitialDir () AS DWSTRING
   RETURN m_wszInitiaDir
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.InitialDir (BYREF wszInitialDir AS WSTRING)
   m_wszInitiaDir = wszInitialDir
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.FilterIndex () AS DWORD
   RETURN m_FilterIndex
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.FilterIndex (BYVAL dwFilterIndex AS DWORD)
   m_FilterIndex = dwFilterIndex
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.FileName () AS DWSTRING
   RETURN m_wszFileName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.FileName (BYREF wszFileName AS WSTRING)
   m_wszFileName = wszFileName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.Filter () AS DWSTRING
   RETURN m_dwsFilter
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY DisplaySaveFile.Filter (BYREF dwsFilter AS DWSTRING)
   m_dwsFilter = dwsFilter
END PROPERTY
' ========================================================================================

END NAMESPACE
