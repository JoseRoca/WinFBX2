'#TEMPLATE DDT Dialog with a size box
#include once "Afx2/DDT.inc"
USING DDT

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_SIZEBOX = 1001

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Create a new dialog using dialog units
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with a Size Box", , , 200, 90, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a size box to the dialog
   ControlAddSizeBox, hDlg, IDC_SIZEBOX

   ' // Display and activate the dialog as modal
   DialogShowModal(hDlg, @DlgProc)

   RETURN DialogEndResult(hDlg)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
         RETURN TRUE

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
               END IF
         END SELECT

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            DIM hSizeBox AS HWND = ControlHandle(hDlg, IDC_SIZEBOX)
            ' // Reposition the sizebox
            IF wParam = SIZE_MAXIMIZED THEN
               ' Hide the sizebox if the window is maximized
               ShowWindow hSizeBox, SW_HIDE
            ELSE
               ' Reposition and show the sizebox
               DIM cx AS LONG = GetSystemMetrics(SM_CXVSCROLL)
               DIM cy AS LONG = GetSystemMetrics(SM_CYHSCROLL)
               SetWindowPos hSizeBox, NULL, LOWORD(lParam) - cx, HIWORD(lParam) - cy, cx, cy, SWP_NOZORDER OR SWP_SHOWWINDOW
            END IF
         END IF

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
