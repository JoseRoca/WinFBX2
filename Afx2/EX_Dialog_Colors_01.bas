#define _CDIALOG_DEBUG_ 1
'#include once "windows.bi"
'#include once "win/commctrl.bi"
'#include once "win/uxtheme.bi"
#include once "Afx2/CDialog.inc"
USING Afx2

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

' // Control identifiers
ENUM
   IDC_EDIT1 = 1001
   IDC_EDIT2
   IDC_LABEL
   IDC_GROUPBOX
   IDC_COMBOBOX
   IDC_SIZEGRIP
END ENUM

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   AfxSetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles
   AfxEnableVisualStyles

   ' // Creeate an instance of the Dialog class. The default font is Segoe UI, 9 points.
   ' // You can specify the font, size, style and charset, e.g. DIM pDlg AS CDialog = CDialog("Tahoma", 9).
   DIM pDlg AS CDialog = CDialog("Segoe UI", 9)
   DIM dwStyle AS LONG = WS_OVERLAPPEDWINDOW
   DIM hDlg AS HWND = pDlg.DialogNewPixels(NULL, "Layout manager test", 0, 0, 0, 0, dwStyle)
   pDlg.DialogSetClient(450, 180)   ' // pixels
   pDlg.DialogCenter

   pDlg.DialogSetColor(0, RGB_GOLD)

   ' // Set icons
   DIM AS HICON hIconBig, hIconSmall
   ExtractIconExW("Shell32.dll", 165, @hIconBig, @hIconSmall, 1)
   pDlg.DialogSetIconEx(hIconBig, hIconSmall)

   ' // Add an Edit control
   DIM hEdit AS HWND = pDlg.ControlAdd("Edit", IDC_EDIT1, "", 15, 15, 305, 23)
   ' // Add a multiline Edit control
   pDlg.ControlAdd("EditMultiline", IDC_EDIT2, "", 15, 45, 305, 80)
   ' // Add more controls
   DIm hGroupBox AS HWND = pDlg.ControlAdd("GroupBox", IDC_GROUPBOX, "GroupBox", 335, 8, 100, 155)
   ' // Disable themes for the GroupBox control; otherwise, the caption colors can't be changed
   SetWindowTheme(hGroupBox, "", "")

   pDlg.ControlAdd("Button", IDCANCEL, "&Close", 245, 140, 75, 23)
   pDlg.ControlAdd("Label", IDC_LABEL, "Label", 50, 140, 75, 23)
   ' // Add a combobox control
   DIM hCombobox AS HWND = pDlg.ControlAdd("ComboBox", IDC_COMBOBOX, "", 345, 30, 80, 100)
   ' // Fill the control with some data
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 1 TO 9
      wszText = "Item " & RIGHT("00" & STR(i), 2)
      SendMessageW(hComboBox, CB_ADDSTRING, 0, CAST(LPARAM, @wszText))
   NEXT

   pDlg.ControlSetColor(IDC_EDIT1, RGB_WHITE, RGB_CORAL)
   pDlg.ControlSetColor(IDC_EDIT2, RGB_WHITE, RGB_LIGHTBLUE)
   pDlg.ControlSetColor(IDC_LABEL, RGB_YELLOW, RGB_GREEN)
   pDlg.ControlSetColor(IDC_GROUPBOX, RGB_WHITE, RGB_RED)

   ' // Anchor the controls
   pDlg.ControlAnchor(IDC_EDIT1, AFX_ANCHOR_WIDTH)
   pDlg.ControlAnchor(IDC_EDIT2, AFX_ANCHOR_HEIGHT_WIDTH)
   pDlg.ControlAnchor(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)
   pDlg.ControlAnchor(IDC_LABEL, AFX_ANCHOR_BOTTOM)
   pDlg.ControlAnchor(IDC_GROUPBOX, AFX_ANCHOR_HEIGHT_RIGHT)
   pDlg.ControlAnchor(IDC_COMBOBOX, AFX_ANCHOR_RIGHT)

   ' // Display and activate the dialog as modal
'   pDlg.DialogShowModal(@DlgProc)

   ' // Display and activate the dialog as modeless
   pDlg.DialogShowModeless(@DlgProc)

   ' // Message handler loop
   DO
      pDlg.DialogDoEvents
   LOOP WHILE IsWindow(hDlg)

   FUNCTION = pDlg.DialogEndResult

   ' // You can use a message pump instead (don't forget to use PostQuitMessage in WM_DESTROY)
'   DIM uMsg AS tagMsg
'   WHILE GetMessage(@uMsg, NULL, 0, 0)
'      IF IsDialogMessage(hDlg, @uMsg) = 0 THEN
'         TranslateMessage @uMsg
'         DispatchMessage @uMsg
'      END IF
'   WEND
'   FUNCTION = uMsg.wParam

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   ' // Pointer to the dialog class
   STATIC pDlg AS CDialog PTR
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
          CDIALOG_DP("WM_INITDIALOG - LPARAM = " & WSTR(lParam))
         ' // Get a pointer to the CDialog class
          pDlg = cast(CDialog PTR, lParam)
          RETURN TRUE

      CASE WM_SYSCOMMAND
         ' // Trap the SC_CLOSE message sent by Alt+F4 and the X-button
         ' // Abort the action
         ' IF (wParam AND &hFFF0) = SC_CLOSE THEN RETURN TRUE
         ' // Or destroy the dialog sending a WM_CLOSE message
         ' IF (wParam AND &hFFF0) = SC_CLOSE THEN
            ' SendMessageW hDlg, WM_CLOSE, 0, 0
            ' RETURN TRUE
         ' END IF
         ' // Or destroy the dialog calling the DialogEnd method
         ' IF (wParam AND &hFFF0) = SC_CLOSE THEN
            ' pDlg->DialogEnd(1)
            ' RETURN TRUE
         ' END IF

      CASE WM_GETMINMAXINFO
         ' Set the pointer to the address of the MINMAXINFO structure
         DIM ptmmi AS MINMAXINFO PTR = CAST(MINMAXINFO PTR, lParam)
         ' Set the minimum and maximum sizes that can be produced by dragging the borders of the window
         ptmmi->ptMinTrackSize.x = 300 * pDlg->rxRatio   ' // provisional values: change me
         ptmmi->ptMinTrackSize.y = 180 * pDlg->ryRatio   ' // provisional values: change me

'      CASE WM_SIZE
'          CDIALOG_DP("WM_SIZE " & WSTR(wParam) & " - width: " & WSTR(LOWORD(lParam)) & " - height: " & (HIWORD(lParam)))
'          No need to adjust anchored controls; they will be adjusted automatically

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_CLOSE
         ' // Multiline edit controls send a WM_CLOSE message when the Esc key is pressed.
         ' // This is a workaround; the proper way is to subclass the control and eat the ESC key.
         IF GetFocus = GetDlgItem(hDlg, IDC_EDIT2) THEN RETURN TRUE   ' // abort
         ' // End the application
         IF pDlg THEN pDlg->DialogEnd(1)
         ' // You can use DestroyWindow instead
         ' DestroyWindow hDlg

      CASE WM_DESTROY
         ' // Post a WM_QUIT message to end the message pump if you are using one
         ' PostQuitMessage(0)

   END SELECT

   RETURN FALSE

END FUNCTION
