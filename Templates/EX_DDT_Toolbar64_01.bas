'#TEMPLATE DDT Dialog with a Toolbar
'#RESOURCE "Toolbar_64.rc"
#define _WIN32_WINNT &h0602
#include once "Afx2/AfxGdiplus2.inc"
USING Afx2
#include once "Afx2/DDT.inc"
USING DDT

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_TOOLBAR = 1001
enum
   IDM_LEFTARROW = 2000
   IDM_RIGHTARROW
   IDM_HOME
   IDM_SAVEFILE
end enum

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with a Toolbar", , , 200, 85, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a toolbar
   ControlAddToolbar, hDlg, IDC_TOOLBAR

   ' // Create an image list for the toolbar
   DIM hImageList AS HIMAGELIST = ImageListNewIcon(64, 64, 32, 4)
   IF hImageList THEN
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_LEFT_64")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_RIGHT_64")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_HOME_64")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_SAVE_64")
      ToolBarSetImageList(hDlg, IDC_TOOLBAR, hImageList)
   END IF
   ' // Create a disabled image list for the toolbar
   DIM hImageListDisabled AS HIMAGELIST = ImageListNewIcon(64, 64, 32, 4)
   IF hImageListDisabled THEN
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_ARROW_LEFT_64", 60, TRUE)
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_ARROW_RIGHT_64", 60, TRUE)
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_HOME_64", 60, TRUE)
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_SAVE_64", 60, TRUE)
      ToolBarSetImageList(hDlg, IDC_TOOLBAR, hImageListDisabled, 1)
   END IF

   ' // Add buttons to the toolbar
   ToolbarAddButton hDlg, IDC_TOOLBAR, 1, IDM_LEFTARROW, 0, "Back"
   ToolbarAddButton hDlg, IDC_TOOLBAR, 2, IDM_RIGHTARROW, 0, "Forward"
   ToolbarAddButton hDlg, IDC_TOOLBAR, 3, IDM_HOME, 0, "Home"
   ToolbarAddButton hDlg, IDC_TOOLBAR, 4, IDM_SAVEFILE, 0, "Save"
   ' // Add a separator
   ToolbarAddSeparator hDlg, IDC_TOOLBAR, 2, 0, 4
   ' // Disable the IDM_SAVEFILE button
   ToolbarDisableButton hDlg, IDC_TOOLBAR, IDM_SAVEFILE

   ' // Add a cancel button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 120, 60, 50, 12

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_TOOLBAR, AFX_ANCHOR_WIDTH)
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
            CASE IDM_LEFTARROW
               MsgBox hDlg, "You have clicked the left arrow button", MB_OK, "Toolbar"
            CASE IDM_RIGHTARROW
               MsgBox hDlg, "You have clicked the right arrow button", MB_OK, "Toolbar"
            CASE IDM_HOME
               MsgBox hDlg, "You have clicked the home button", MB_OK, "Toolbar"
            CASE IDM_SAVEFILE
               MsgBox hDlg, "You have clicked the save button", MB_OK, "Toolbar"
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
