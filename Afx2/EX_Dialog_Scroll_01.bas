#define UNICODE
#define _WIN32_WINNT &h0602
#define _CDIALOG_DEBUG_ 1
#include once "Afx2/CDialog.inc"
USING Afx2

#define IDC_LISTBOX 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

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
   DIM hDlg AS HWND = pDlg.DialogNewPixels(0, "Scrollable dialog", 0, 0, 0, 0, dwStyle)
   ' // Set a client size big enough to display all the controls
   pDlg.DialogSetClient(320, 335)
   ' // Make the dialog scrollable by shrinking the client size
   pDlg.DialogSetViewPort(250, 260)

   ' // Set icons
   DIM AS HICON hIconBig, hIconSmall
   ExtractIconExW("Shell32.dll", 165, @hIconBig, @hIconSmall, 1)
   pDlg.DialogSetIconEx(hIconBig, hIconSmall)

   ' // Add a listbox
   DIM hListBox AS HWND = pDlg.ControlAdd("ListBox", IDC_LISTBOX)
   SetWindowPos hListBox, NULL, 8 * pDlg.rxRatio, 8 * pDlg.ryRatio, 300 * pDlg.rxRatio, 280 * pDlg.ryRatio, SWP_NOZORDER

   ' // Fill the list box
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 1 TO 50
      wszText = "Item " & RIGHT("00" & WSTR(i), 2)
      SendMessageW(hListBox, LB_ADDSTRING, 0, CAST(LPARAM, @wszText))
   NEXT
   ' // Select the first item
   SendMessageW(hListBox, LB_SETCURSEL, 0, 0)

   ' // Add a cancel button
   pDlg.ControlAdd("Button", IDCANCEL, "&Cancel", 233, 298, 75, 23)

   ' // Center the window (must be done after shrinking the client size)
   pDlg.DialogCenter

   ' // Display and activate the dialog as modal
'   pDlg.DialogShowModal(@DlgProc)

   ' // Display and activate the dialog as modeless
   pDlg.DialogShowModeless(@DlgProc)

   ' // Message handler loop
   DO
      pDlg.DialogDoEvents
   LOOP WHILE IsWindow(hDlg)

   FUNCTION = pDlg.DialogEndResult

   ' // You can use a message pump instead
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
         OutputDebugStringW("DlgProc - WM_INITDIALOG " & WSTR(hDlg) & " - " & WSTR(wParam) & " - " & WSTR(lParam))
         ' // Get a pointer to the CDialog class
          pDlg = cast(CDialog PTR, lParam)

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_LISTBOX
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         IF pDlg THEN pDlg->DialogEnd(1)

      CASE WM_DESTROY
         OutputDebugStringW("DlgProc - WM_DESTROY")
         '// If using a message pumo, you mustpost a WM_QUIT message to exit the loop
         'PostQuitMessage(0)

   END SELECT

   RETURN FALSE

END FUNCTION
