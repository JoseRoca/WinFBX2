# CButton Class

Wrapper class on top of the `CreateDialogIndirectParamW` API function.

**Include file**: CDialog.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructor) | Creates instances of the `CDialog` class. |
| [RegisterClass](#registerclass) | Registers a window class. |
| [DialogTemplate](#dialogtemplate) | Creates a dialog template in memory. |
| [ControlAdd](#controladd) | Adds a control to the dialog. |
| [hDialog](#hdialog) | Gets/sets the handle of the dialog. |
| [IsModal](#ismodal) | Checks if the dialog is modal. |
| [IsCustom](#iscustom) | Checks if it is a custom dialog. |
| [UsesPixels](#usespixels) | Checks if the dialog uses pixels. |
| [UsesUnits](#usesunits) | Checks if the dialog uses dialog units. |
| [DialogNew](#dialognew) | Creates a new dialog using dialog units. |
| [DialogNewPixels](#dialognewpixels) | Creates a new dialog using pixels. |
| [DialogShowModal](#dialogshowmodal) | Shows the dialog as modal. |
| [DialogShowModeless](#dialogshowmodeless) | Shows the dialog as modeless. |
| [DialogEnd](#dialogend) | End the dialog. |
| [DialogEndResult](#dialogendresult) | Returns the value passed to DialogEnd. |
| [DialogDoEvents](#dialogdoevents) | Message pump for modeless dialogs. |
| [GetDefId](#getdefid) | Retrieves the identifier of the default push button control for a dialog box. |
| [SetDefId](#setdefid) | Changes the identifier of the default push button for a dialog box. |
| [DialogReposition](#dialogreposition) | Repositions a top-level dialog box so that it fits within the desktop area. |
| [DialogShowState](#dialogshowstate) | Change the visible state of a dialog. |
| [DialogMaximize](#dialogmaximize) | Maximizes the dialog. |
| [DialogMinimize](#dialogminimize) | Minimizes the dialog. |
| [DialogHide](#dialoghide) | Hides the dialog. |
| [DialogNormalize](#dialognormalize) | Makes the dialog visible at its normal size and position. |
| [DialogDisable](#dialogdisable) | Disables the dialog. |
| [DialogEnable](#dialogenable) | Enables the dialog. |
| [DialogRedraw](#dialogredraw) | Redraws the dialog. |
| [ForceVisibleDisplay](#forcevisibledisplay) | Repositions the dialog if it is off-screen. |
| [DialogCenter](#dialogcenter) | Centers the dialog. |
| [DialogStabilize](#dialogstabilize) | Makes a Dialog stabilized (non-closeable). |
| [DialogNonStable](#dialognonstable) | Makes a Dialog non stable (closeable). |
| [IsDialogStabilized](#isdialogstabilized) | Checks if the dialog is stabilized. |
| [IsDialogNonStable](#isdialognonstable) | Checks if the dialog is non stable. |
| [DialogPost](#dialogpost) | Posts a message in the message queue. |
| [DialogSend](#dialogsends) | Sends a message to the dialog. |

---

# <a name="constructor"></a>Constructor

Creates instances of the `CDialog` class.

```
CONSTRUCTOR (BYREF fontName AS WSTRING = "Segoe UI", BYVAL ptSize AS LONG = 9, _
   BYVAL fontStyle AS BYTE = 0, BYVAL charset AS BYTE = DEFAULT_CHARSET) 
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontName* | Font name. |
| *ptSize* | Font point size. |
| *fontStyle* | Font style. |
| *charset* | Font character set. |

---

