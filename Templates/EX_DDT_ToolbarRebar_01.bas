'#TEMPLATE DDT Dialog - ToolBar-Rebar
'#RESOURCE "ToolbarRebar.rc"
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
CONST IDC_REBAR = 1002
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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog Toolbar-Rebar", , , 200, 85, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a toolbar
   ControlAddToolbar, hDlg, IDC_TOOLBAR

   ' // Calculate the size of the icon according the DPI
   DIM cx AS LONG = 24 * GetDpiForSystem \ 96

   ' // Create an image list for the toolbar
   DIM hImageList AS HIMAGELIST = ImageListNewIcon(cx, cx, 32, 4)
   IF hImageList THEN
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_LEFT_48")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_RIGHT_48")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_HOME_48")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_SAVE_48")
      ToolBarSetImageList(hDlg, IDC_TOOLBAR, hImageList)
   END IF
   ' // Create a disabled image list for the toolbar
   DIM hImageListDisabled AS HIMAGELIST = ImageListNewIcon(cx, cx, 32, 4)
   IF hImageListDisabled THEN
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_ARROW_LEFT_48", 60, TRUE)
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_ARROW_RIGHT_48", 60, TRUE)
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_HOME_48", 60, TRUE)
      AfxGdipAddIconFromRes(hImageListDisabled, hInstance, "IDI_SAVE_48", 60, TRUE)
      ToolBarSetImageList(hDlg, IDC_TOOLBAR, hImageListDisabled, 1)
   END IF

   ' // Add buttons to the toolbar
   ToolbarAddButton hDlg, IDC_TOOLBAR, 1, IDM_LEFTARROW
   ToolbarAddButton hDlg, IDC_TOOLBAR, 2, IDM_RIGHTARROW
   ToolbarAddButton hDlg, IDC_TOOLBAR, 3, IDM_HOME
   ToolbarAddButton hDlg, IDC_TOOLBAR, 4, IDM_SAVEFILE
   ' // Disable the IDM_SAVEFILE button
   ToolbarDisableButton hDlg, IDC_TOOLBAR, IDM_SAVEFILE

   ' // Create a rebar control
   ControlAddRebar, hDlg, IDC_REBAR
   DIM hRebar AS HWND = ControlHandle(hDlg, IDC_REBAR)

   ' // Add the band containing the toolbar control to the rebar
   ' // The size of the REBARBANDINFOW is different in Vista/Windows 7
   DIM rc AS RECT, wszText AS WSTRING * 260
   DIM rbbi AS REBARBANDINFOW
   IF AfxWindowsVersion >= 600 AND AfxComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF

   ' // Insert the toolbar in the rebar control
   DIM hToolBar AS HWND = ControlHandle(hDlg, IDC_TOOLBAR)
   rbbi.fMask      = RBBIM_STYLE OR RBBIM_CHILD OR RBBIM_CHILDSIZE OR _
                     RBBIM_SIZE OR RBBIM_ID OR RBBIM_IDEALSIZE OR RBBIM_TEXT
   rbbi.fStyle     = RBBS_CHILDEDGE
   rbbi.hwndChild  = hToolBar
   rbbi.cxMinChild = ScaleForDpiX(hDlg, 270)
   rbbi.cyMinChild = Toolbar_GetButtonHeight(hToolbar)
   rbbi.cx         = ScaleForDpiX(hDlg, 270)
   rbbi.cxIdeal    = ScaleForDpiX(hDlg, 270)
   wszText = "Toolbar"
   rbbi.lpText = @wszText
   '// Insert band into rebar
   Rebar_InsertBand(hRebar, -1, @rbbi)

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

      CASE WM_NOTIFY
         ' -------------------------------------------------------
         ' Notification messages are handled here.
         ' The TTN_GETDISPINFO message is sent by a ToolTip control
         ' to retrieve information needed to display a ToolTip window.
         ' ------------------------------------------------------
         DIM ptnmhdr AS NMHDR PTR              ' // Information about a notification message
         DIM ptttdi AS NMTTDISPINFOW PTR       ' // Tooltip notification message information
         DIM wszTooltipText AS WSTRING * 260   ' // Tooltip text

         DIM hRebar AS HWND = ControlHandle(hDlg, IDC_REBAR)
         ptnmhdr = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->code
            ' // The height of the rebar has changed
            CASE RBN_HEIGHTCHANGE
               ' // Get the coordinates of the client area
               DIM rc AS RECT
               GetClientRect hDlg, @rc
               ' // Send a WM_SIZE message to resize the controls
               SendMessageW hDlg, WM_SIZE, SIZE_RESTORED, MAKELONG(rc.Right - rc.Left, rc.Bottom - rc.Top)
            ' // Toolbar tooltips
            CASE TTN_GETDISPINFOW
               ptttdi = CAST(NMTTDISPINFOW PTR, lParam)
               ptttdi->hinst = NULL
               wszTooltipText = ""
               SELECT CASE ptttdi->hdr.hwndFrom
                  CASE Toolbar_GetTooltips(GetDlgItem(GetDlgItem(hDlg, IDC_REBAR), IDC_TOOLBAR))
                     SELECT CASE ptttdi->hdr.idFrom
                        CASE IDM_LEFTARROW  : wszTooltipText = "Back"
                        CASE IDM_RIGHTARROW : wszTooltipText = "Forward"
                        CASE IDM_HOME       : wszTooltipText = "Home"
                        CASE IDM_SAVEFILE   : wszTooltipText = "Save"
                     END SELECT
                     IF LEN(wszTooltipText) THEN ptttdi->lpszText = @wszTooltipText
               END SELECT
         END SELECT

      CASE WM_SIZE
   IF wParam <> SIZE_MINIMIZED THEN
' Force rebar to recalculate layout
SendMessageW(ControlHandle(hDlg, IDC_REBAR), WM_SIZE, 0, 0)

' Reassert visibility and z-order
SetWindowPos(ControlHandle(hDlg, IDC_REBAR), HWND_TOP, 0, 0, 0, 0, _
             SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_SHOWWINDOW)

SetWindowPos(ControlHandle(hDlg, IDC_TOOLBAR), HWND_TOP, 0, 0, 0, 0, _
             SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_SHOWWINDOW)

' Force repaint of both controls
InvalidateRect(ControlHandle(hDlg, IDC_REBAR), NULL, TRUE)
InvalidateRect(ControlHandle(hDlg, IDC_TOOLBAR), NULL, TRUE)
   END IF

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
