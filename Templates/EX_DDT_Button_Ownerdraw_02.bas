'#TEMPLATE DDT Dialog with an ownerdraw button with an image
#define _WIN32_WINNT &h0602
#include once "Afx2/DDT.inc"
USING DDT

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_OK = 1001

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dalog with an ownerdraw button",50, 50, 175, 65, _
      WS_OVERLAPPEDWINDOW OR DS_CENTER, , "Segoe UI", 9)

   ' // Add a button to the dialog
   ControlAdd("OwnerdrawButton", hDlg, IDC_OK, "&Ok", 105, 40, 50, 16, WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW OR WS_BORDER)

   ' // Anchor the button to the bottom and the right side of the dialog
   ControlAnchor(hDlg, IDC_OK, AFX_ANCHOR_BOTTOM_RIGHT)

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
            CASE IDC_OK
               MsgBox("OK button pressed", MB_ICONEXCLAMATION OR MB_TASKMODAL, "Message")
         END SELECT

      CASE WM_DRAWITEM
         DIM pDis AS DRAWITEMSTRUCT PTR = CAST(DRAWITEMSTRUCT PTR, lParam)
         IF pDis->CtlId <> IDC_OK THEN RETURN FALSE
         ' Icon identifiers in User32.dll.
         ' IDI_APPLICATION = 100, IDI_WARNING = 101, IDI_QUESTION = 102, IDI_ERROR = 103
         ' IDI_INFORMATION = 104, IDI_WINLOGO = 105, IDI_SHIELD = 106
         ' Used here to make the example simpler.
         ' Normally, you will load an icon or bitmap from a file or resource.
         DIM hIcon AS HICON = LoadImageW(GetModuleHandle("User32"), MAKEINTRESOURCEW(103), _
            IMAGE_ICON, AfxScaleX(30), AfxScaleY(30), LR_DEFAULTCOLOR)
         IF hIcon THEN
            DrawIcon pDis->hDC, 45, 0, hIcon
            DestroyIcon hIcon
         END IF
         RETURN TRUE

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
