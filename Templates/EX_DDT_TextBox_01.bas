'#TEMPLATE DDT Dialog with textboxes
#include once "Afx2/DDT.inc"
USING DDT

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_EDIT1 = 1001
CONST IDC_EDIT2 = 1002

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dalog with Textboxes", , , 191, 100, WS_OVERLAPPEDWINDOW OR DS_CENTER)
   ' // Set the dialog's backgroung color
   DialogSetColor(hDlg, -1, RGB_GOLD)

   ' // Add a Textbox control
   ControlAddTextbox, hDlg, IDC_EDIT1, "", 8, 8, 174, 12
   ' // Add a Multiline Textbox control
   ControlAddMultilineTextbox, hDlg, IDC_EDIT2, "", 8, 24, 174, 43
   ' // Add a button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 132, 80, 50, 12

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_EDIT1, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDC_EDIT2, AFX_ANCHOR_HEIGHT_WIDTH)
   ControlAnchor(hDlg, IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

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

      CASE WM_CLOSE
         ' // Multiline edit controls send a WM_CLOSE message when the Esc key is pressed.
         ' // This is a workaround; the proper way is to subclass the control and eat the ESC key.
         IF GetFocus = ControlHandle(hDlg, IDC_EDIT2) THEN RETURN TRUE   ' // abort
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
