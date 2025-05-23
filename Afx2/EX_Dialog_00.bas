#define _CDIALOG_DEBUG_ 1
#include once "Afx2/CDialog.inc"
USING Afx2

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
   DIM dwStyle AS LONG = WS_OVERLAPPEDWINDOW ' OR DS_MODALFRAME OR DS_CENTER
   DIM hDlg AS HWND = pDlg.DialogNew(0, "Dialog New", 0, 0, 0, 0, dwStyle) ', 50, 50, 175, 75, dwStyle)
   IF pDlg.UsesUnits THEN
      pDlg.DialogSetClient(175, 75)   ' // dialog units
   ELSE
      pDlg.DialogSetClient(306, 141)   ' // pixels
   END IF
   pDlg.DialogCenter

   ' // Set icons
   DIM AS HICON hIconBig, hIconSmall
   ExtractIconExW("Shell32.dll", 165, @hIconBig, @hIconSmall, 1)
   pDlg.DialogSetIconEx(hIconBig, hIconSmall)

   ' // Remove the system close menu and disable the X button
   ' pDlg.SysMenuRemoveCloseOption
   ' // Retore the system close menu
   ' pDlg.SysMenuRestoreCloseOption

'   ' // Add controls to the dialog
'   IF pDlg.UsesUnits THEN
'      pDlg.ControlAdd "GroupBox", 101, "Just a simple question", 5, 5, 160, 55
'      pDlg.ControlAdd "Label", 102, "What's your name ?", 10, 20, 80, 10
'      pDlg.ControlAdd "Edit", 103, "", 90, 19, 70, 12
'      DIM hCtl AS HWND = pDlg.ControlAdd("Button", 104, "&Ok", 80, 40, 50, 12, BS_DEFPUSHBUTTON)
'   ELSE   ' // pixels
'      pDlg.ControlAdd "GroupBox", 101, "Just a simple question", 9, 9, 280, 103
'      pDlg.ControlAdd "Label", 102, "What's your name ?", 18, 38, 140, 19
'      pDlg.ControlAdd "Edit", 103, "", 158, 36, 122, 22
'      DIM hCtl AS HWND = pDlg.ControlAdd("Button", 104, "&Ok", 140, 75, 88, 22, BS_DEFPUSHBUTTON)
'   END iF
   
      ' // Add controls to the dialog
   DIM AS LONG x1, x2, x3, x4, x5, x6, x7, x8
   DIM AS LONG cx1, cx2, cx3, cx4, cx5, cx6, cx7, cx8
   IF pDlg.UsesUnits THEN
      x1 =  5 : x2 =  5 : cx1 = 160 : cx2 = 55
      x3 = 10 : x4 = 20 : cx3 =  80 : cx4 = 10
      x5 = 90 : x6 = 19 : cx5 =  70 : cx6 = 12
      x7 = 80 : x8 = 40 : cx7 =  50 : cx8 = 12
   ELSE   ' // pixels
      x1 =   9 : x2 =  9 : cx1 = 280 : cx2 = 103
      x3 =  18 : x4 = 38 : cx3 = 140 : cx4 =  19
      x5 = 158 : x6 = 36 : cx5 = 122 : cx6 =  22
      x7 = 140 : x8 = 75 : cx7 =  88 : cx8 =  22
   END iF
   

   pDlg.ControlAdd "GroupBox", 101, "Just a simple question", x1, x2, cx1, cx2
   pDlg.ControlAdd "Label", 102, "What's your name ?", x3, x4, cx3, cx4
   pDlg.ControlAdd "Edit", 103, "", x5, x6, cx5, cx6
   DIM hCtl AS HWND = pDlg.ControlAdd("Button", 104, "&Ok", x7, x8, cx7, cx8, BS_DEFPUSHBUTTON)

   ' // Set the focus in the edit control
'   pDlg.ControlSetFocus(103)

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
