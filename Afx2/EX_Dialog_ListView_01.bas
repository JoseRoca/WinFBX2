#define _CDIALOG_DEBUG_ 1
#include once "Afx2/CDialog.inc"
#include once "Afx2/AfxDpiProcs2.inc"
#include once "Afx2/CListview.inc"
USING Afx2

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

#define IDC_LISTVIEW 1001
#define IDC_TEST 1002
#define IDC_OK 1003

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   AfxSetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Create an instance of the Dialog class. The default font is Segoe UI, 9 points.
   ' // You can specify the font, size, style and charset, e.g. DIM pDlg AS CDialog = CDialog("Tahoma", 9).
   DIM pDlg AS CDialog = CDialog("Segoe UI", 9)
   DIM dwStyle AS LONG = WS_OVERLAPPEDWINDOW OR DS_MODALFRAME OR DS_CENTER

   ' // Create dialog using dialog units
   DIM hDlg AS HWND = pDlg.DialogNew(0, "Dialog New",,, 440, 195, dwStyle)
'   ' // Create custom dialog using dialog units
'   DIM hDlg AS HWND = pDlg.DialogNew("MyClassName", 0, "Dialog New",,, 440, 195, dwStyle)

   ' // Create dialog using pixels
'   DIM hDlg AS HWND = pDlg.DialogNewPixels(0, "Dialog New",,, 770, 455, dwStyle)
   ' // Create custom dialog using pixels
'   DIM hDlg AS HWND = pDlg.DialogNewPixels("MyClassName", 0, "Dialog New",,, 770, 455, dwStyle)
   ' // Set icons
   DIM AS HICON hIconBig, hIconSmall
   ExtractIconExW("Shell32.dll", 165, @hIconBig, @hIconSmall, 1)
   pDlg.DialogSetIconEx(hIconBig, hIconSmall)

   ' // Remove the system close menu and disable the X button
   ' pDlg.SysMenuRemoveCloseOption
   ' // Retore the system close menu
   ' pDlg.SysMenuRestoreCloseOption  
   ' // Add controls to the dialog

   DIM hListView AS HWND
   IF pDlg.UsesUnits THEN
      hListView = pDlg.ControlAdd("ListView", IDC_LISTVIEW, "", 5, 5, 430, 165)
      pDlg.ControlAdd("Button", IDC_TEST, "&Test", 5, 176, 50, 12)
      pDlg.ControlAdd("Button", IDC_OK, "&Ok", 60, 176, 50, 12, BS_DEFPUSHBUTTON)
   ELSE
      hListView = pDlg.ControlAdd("ListView", IDC_LISTVIEW, "", 10, 10, 735, 350)
      pDlg.ControlAdd("Button", IDC_TEST, "&Test", 10, 375, 75, 25)
      pDlg.ControlAdd("Button", IDC_OK, "&Ok", 100, 375, 75, 25, BS_DEFPUSHBUTTON)
   END IF

   ' // Add some extended styles
   DIM pListView AS CListView = hListView
   DIM dwExStyle AS DWORD
   dwExStyle = pListview.GetExtendedListViewStyle
   dwExStyle = dwExStyle OR LVS_EX_FULLROWSELECT OR LVS_EX_GRIDLINES
   pListView.SetExtendedListViewStyle(dwExStyle)

   ' // Add the header's column names
   DIM lvc AS LVCOLUMNW
   lvc.mask = LVCF_FMT OR LVCF_WIDTH OR LVCF_TEXT OR LVCF_SUBITEM
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 0 TO 9
      wszText = "Column " & STR(i)
      pListView.AddColumn(i, wszText, AfxScaleX(105))
   NEXT

   ' // Populate the ListView with some data
   DIM lvi AS LVITEMW
   lvi.mask = LVIF_TEXT
   FOR i AS LONG = 0 to 29
      lvi.iItem = i
      wszText = "Column 1 Row" + WSTR(i)
      lvi.pszText = @wszText
      pListView.InsertItem(lvi)
      FOR x AS LONG = 0 TO 9
         wszText = "Column " & WSTR(x) & " Row" + WSTR(i)
         pListView.SetItemText(i, x, wszText)
      NEXT
   NEXT

   ' // Set the focus in the edit control
'   pDlg.ControlSetFocus(IDC_LISTVIEW)

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
'         OutputDebugStringW("DlgProc - WM_INITDIALOG " & WSTR(hDlg))
         OutputDebugStringW("DlgProc - WM_INITDIALOG " & WSTR(wParam) & " -" & GetDlgItem(hDlg, IDC_LISTVIEW))
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
         RETURN TRUE


      CASE WM_SYSCOMMAND
         CDIALOG_DP("WM_SYSCOMMAND")
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
               OutputDebugStringW("DlgProc - WM_COMMAND - WM_CLOSE")
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_OK
               MessageBoxW(0, "Ok", "Message", MB_ICONEXCLAMATION OR MB_TASKMODAL)
            CASE 104 ' Test

            CASE IDC_TEST
               SCOPE
                  DIM pListView AS CListView = CListView(hDlg, IDC_LISTVIEW)
'                  pListView.PixelsToDialogUnitsX(GetDlgItem(hDlg, IDC_LISTVIEW), 1)
'                  pListView.PixelsToDialogUnitsY(GetDlgItem(hDlg, IDC_LISTVIEW), 1)

'DIM AS SINGLE rx, ry
'DIM AS LONG pixx, pixy
'pDlg->DialogUnitsToPixels(100, 100, pixx, pixy)
'print "++++++++++", pixx, pixy
'pDlg->DialogUnitsToPixelsRatios(100, 100, rx, ry)
'print "++++++++++", rx, ry
'print 100*rx, 100*ry

'DIM AS LONG dlux, dluy
'pDlg->PixelsToDialogUnits(300, 375, dluX, dluY)
'print dluX, dluY
'pDlg->PixelsToDialogUnitsRatios(rx, ry)
'print "rx, ry ", rx, ry
'print 300 * rx, 375* ry

'DIM AS SINGLE rx, ry
'DIM AS LONG pixx, pixy
'DIM hListView AS HWND = GetDlgItem(hDlg, IDC_LISTVIEW)
'pListView.PixelsToDialogUnits(hListView, 100, 100, pixx, pixy)
'print "++++++++++", pixx, pixy
'pListView.PixelsToDialogUnitsRatios(hListView, rx, ry)
'print "++++++++++", rx, ry
'print 100*rx, 100*ry
'print pDlg->DlgToPixRX
'print pDlg->DlgToPixRY
'print pDlg->PixToDlgRX
'print pDlg->PixToDlgRY
'print "-------------------------"
'DIM hListView AS HWND = GetDlgItem(hDlg, IDC_LISTVIEW)
'print pListView.DlgToPixRX()
'print pListView.DlgToPixRY()
'print pListView.PixToDlgRX()
'print pListView.PixToDlgRY()


               END SCOPE

         END SELECT

      CASE WM_SIZE
         ' // Resize the ListView control and its header
'         IF wParam <> SIZE_MINIMIZED THEN
'            ' // Retrieve the handle of the ListView control
'            DIM hListView AS HWND = GetDlgItem(hDlg, IDC_LISTVIEW)
'            ' // Move the ListView control
'            MoveWindow hListView, 20, 20, pDlg->ControlGetClientWidth(IDC_LISTVIEW) - 5, pDlg->ControlGetClientHeight(IDC_LISTVIEW) - 5, TRUE
'         END IF

      CASE WM_CLOSE
         OutputDebugStringW("DlgProc - WM_CLOSE")
         ' // End the application
         IF pDlg THEN pDlg->DialogEnd(1)
         ' // You can use DestroyWindow instead
         ' DestroyWindow hDlg

      CASE WM_DESTROY
         OutputDebugStringW("DlgProc - WM_DESTROY")

   END SELECT

   RETURN FALSE

END FUNCTION
