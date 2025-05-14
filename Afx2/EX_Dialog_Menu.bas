' ========================================================================================
' Test menus PB-style with dialog
' ========================================================================================

#include once "Afx2/CDialog.inc"
#include once "Afx2/AfxMenuProcs2.inc"
USING Afx2

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST ID_OPEN = 401
CONST ID_EXIT = 402
CONST ID_OPTION1 = 403
CONST ID_OPTION2 = 404
CONST ID_HELP = 405
CONST ID_ABOUT = 406


FUNCTION BuildMenu (BYVAL hDlg AS HWND) AS HMENU
   DIM lResult AS LONG

   ' ** First create a top-level menu:
   DIM hMenu AS HMENU = MenuNewBar

   ' ** Add a top-level menu item with a popup menu:
   DIM hPopup1 AS HMENU = MenuNewPopup
   MenuAddPopup hMenu, "&File", hPopup1, MF_ENABLED
   MenuAddString hPopup1, "&Open", ID_OPEN, MF_ENABLED
   MenuAddString hPopup1, "&Exit", ID_EXIT, MF_ENABLED
   MenuAddString hPopup1, "-", 0, 0

   ' ** Now we can add another item to the menu that will bring up a sub-menu. 
   ' First we obtain a new popup menu handle to distinguish it from the first 
   ' popup menu:
   DIM hPopup2 AS HMENU = MenuNewPopup

   ' ** Now add a new menu item to the first menu. 
   ' This item will bring up the sub-menu when selected:
   MenuAddPopup hPopup1, "&More Options", hPopup2, MF_ENABLED

   ' ** Now we will define the sub menu:
   MenuAddString hPopup2, "Option &1", ID_OPTION1, MF_ENABLED
   MenuAddString hPopup2, "Option &2", ID_OPTION2, MF_ENABLED

   ' ** Finally, we'll add a second top-level menu and popup.
   ' For this popup, we can reuse the first popup variable:
   hPopup1 = MenuNewPopup
   MenuAddPopup hMenu, "&Help", hPopup1, MF_ENABLED
   MenuAddString hPopup1, "&Help", ID_HELP, MF_ENABLED
   MenuAddString hPopup1, "-", 0, 0
   MenuAddString hPopup1, "&About", ID_ABOUT, MF_ENABLED
   
   ' Attach the menu to the dialog
   MenuAttach hMenu, hDlg

   ' // Bold the Exit menu item
   MenuBoldItem(hMenu, ID_EXIT)
   ' // Right justify the Help menu
'   MenuRightJustifyItem(hMenu, 1)

   RETURN hMenu
END FUNCTION

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   AfxSetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)

   ' // Creeate an instance of the Dialog class. The default font is Segoe UI, 9 points.
   ' // You can specify the font, size, style and charset, e.g. DIM pDlg AS CDialog = CDialog("Tahoma", 9).
   DIM pDlg AS CDialog = CDialog("Segoe UI", 9)
   DIM dwStyle AS LONG = WS_OVERLAPPEDWINDOW OR DS_MODALFRAME OR DS_CENTER
   DIM hDlg AS HWND = pDlg.DialogNew(0, "Dialog New", 50, 50, 175, 75, dwStyle)

   ' // Set icons
   DIM AS HICON hIconBig, hIconSmall
   ExtractIconExW("Shell32.dll", 165, @hIconBig, @hIconSmall, 1)
   pDlg.DialogSetIconEx(hIconBig, hIconSmall)

   ' // Remove the system close menu and disable the X button
   ' pDlg.SysMenuRemoveCloseOption
   ' // Retore the system close menu
   ' pDlg.SysMenuRestoreCloseOption
   
   ' // Add controls to the dialog
   pDlg.ControlAdd "GroupBox", 101, "Just a simple question", 5, 5, 160, 55
   pDlg.ControlAdd "Label", 102, "What's your name ?", 10, 20, 80, 10
   pDlg.ControlAdd "Edit", 103, "", 90, 19, 70, 12
   DIM hCtl AS HWND = pDlg.ControlAdd("Button", 104, "&Ok", 80, 40, 50, 12, BS_DEFPUSHBUTTON)

   ' // Build and attach a menu
   DIM hMenu AS HMENU = BuildMenu(hDlg)

   ' // Create a keyboard accelerator table
   pDlg.AddAccelerator FVIRTKEY OR FCONTROL, "O", ID_OPEN  ' // Ctrl+O - Open
   pDlg.AddAccelerator FVIRTKEY OR FCONTROL, "H", ID_HELP  ' // Ctrl+H - Help
   pDlg.AddAccelerator FVIRTKEY OR FCONTROL, "E", ID_EXIT  ' // Ctrl+E - Exit
   pDlg.AddAccelerator FVIRTKEY OR FCONTROL, "A", ID_ABOUT ' // Ctrl+A - About
   pDlg.CreateAcceleratorTable

   ' // Display and activate the dialog as modal
'   pDlg.DialogShowModal(@DlgProc)
'   FUNCTION = pDlg.DialogEndResult

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
         ' // You can add controls here this way:
         ' // Get a pointer to the CDialog class
          pDlg = cast(CDialog PTR, lParam)
'         ' // Add a menu
'         DIM hMenu AS HMENU = BuildMenu
'         pDlg.MenuAttach hMenu
'         ' // Add controls to the dialog
'         pDlg.ControlAdd "GroupBox", 101, "Just a simple question", 5, 5, 160, 55
'         pDlg.ControlAdd "Label", 102, "What's your name ?", 10, 20, 80, 10
'         pDlg.ControlAdd "Edit", 103, "", 90, 19, 70, 12
'         pDlg.ControlAdd "Button", 104, "Test", 10, 40, 50, 12, BS_DEFPUSHBUTTON
'         pDlg.ControlAdd "Button", 105, "&Ok", 80, 40, 50, 12


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


      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE 104 'IDOK
               DIM dws AS DWSTRING = AfxGetWindowText(GetDlgItem(hDlg, 103))
               IF LEN(dws) THEN MessageBoxW(0, "Your name is " & dws, "Answer", MB_ICONINFORMATION OR MB_TASKMODAL)
               IF LEN(dws) = 0 THEN MessageBoxW(0, "Hmm ...", "No answer", MB_ICONEXCLAMATION OR MB_TASKMODAL)
           CASE ID_OPEN TO ID_ABOUT
               MessageBoxW(0, "WM_COMMAND received from a menu item!", "Test menu", MB_TASKMODAL)
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         IF pDlg THEN pDlg->DialogEnd(1)
         ' // You can use DestroyWindow instead
         ' DestroyWindow hDlg

      CASE WM_DESTROY

   END SELECT

   RETURN FALSE

END FUNCTION
