# CDialog Class

**Introducing CDialog: A Streamlined Dialog Engine for FreeBasic**

`CDialog` is a lightweight yet powerful dialog engine tailored for FreeBasic developers seeking a clean, high-level approach to building Windows dialogs — without sacrificing control or clarity. Inspired by the straightforwardness of PowerBasic’s DDT model, yet engineered with modern design in mind, `CDialog` offers structured, maintainable UI development rooted in the familiar logic of the Windows API.

At its core, `CDialog` replicates and refines the DDT engine, delivering cleaner syntax, robust internal architecture, and simplified message handling. Whether you're building modal or modeless dialogs, adding controls, or managing layout behavior, `CDialog` keeps things intuitive and transparent — no bloat, no black boxes.

Among its standout features is the automatic anchoring system, which minimizes the tedium of UI resizing logic. By applying simple declarative rules, it handles control repositioning and scaling automatically — sparing developers the hassle of manual calculations. Crucially, this system works identically for dialogs defined in pixels or dialog units, providing consistency and flexibility across design styles.

As High DPI support has evolved — with dynamic DPI changes and multi-monitor setups becoming standard — managing traditional SDK windows has grown increasingly complex. While SDK windows offer unmatched flexibility, responding to real-time DPI shifts now demands considerable effort, including font recreation and intricate layout recalculations. `CDialog` embraces dialogs not as a limitation, but as a framework aligned with Microsoft's DPI-conscious strategy, offloading much of the adaptation logic to the dialog engine itself. Complemented by its built-in anchoring system, `CDialog` ensures that controls scale and reposition smoothly regardless of whether the layout uses pixels or dialog units — all without cluttering WM_SIZE handlers with tedious calculations.

And for those who prefer procedural syntax, `DDT.inc` bridges the gap by exposing DDT-style wrappers atop the `CDialog` class — allowing developers to adopt `CDialog` incrementally or stick with the idioms they already know.

Whether you're modernizing legacy utilities, crafting new tools, or just writing code for fun, `CDialog` delivers a flexible, DPI-aware foundation that respects both your time and your muscle memory.

**Include file**: CDialog.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructor) | Creates instances of the `CDialog` class. |
| [RegisterClass](#registerclass) | Registers a window class. |
| [DialogTemplate](#dialogtemplate) | Creates a dialog template in memory. |
| [hDialog](#hdialog) | Gets/sets the handle of the dialog. |
| [IsModal](#ismodal) | Checks if the dialog is modal. |
| [IsCustom](#iscustom) | Checks if it is a custom dialog. |
| [UsesPixels](#usespixels) | Checks if the dialog uses pixels. |
| [UsesUnits](#usesunits) | Checks if the dialog uses dialog units. |

---

# Dialog methods

| Name       | Description |
| ---------- | ----------- |
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
| [DialogForceVisibility](#dialogforcevisibility) | Repositions the dialog if it is off-screen. |
| [DialogCenter](#dialogcenter) | Centers the dialog. |
| [DialogStabilize](#dialogstabilize) | Makes a Dialog stabilized (non-closeable). |
| [DialogNonStable](#dialognonstable) | Makes a Dialog non stable (closeable). |
| [IsDialogStabilized](#isdialogstabilized) | Checks if the dialog is stabilized. |
| [IsDialogNonStable](#isdialognonstable) | Checks if the dialog is non stable. |
| [DialogPost](#dialogpost) | Posts a message in the message queue. |
| [DialogSend](#dialogsends) | Sends a message to the dialog. |
| [DialogGetSize](#dialoggetsize) | Gets the width and height of the dialog. |
| [DialogGetWidth](#dialoggetwidth) | Gets the width of the dialog. |
| [DialogGetHeight](#dialoggetheight) | Gets the height of the dialog. |
| [DialogSetSize](#dialogsetsize) | Sets the size of the dialog. |
| [DialogGetBounds](#dialoggetbounds) | Retrieves the bounds of a window without the drop shadows. |
| [DialogGetClient](#dialoggetclient) | Retrieves the coordinates of the dialog's client area. |
| [DialogGetClientWidth](#dialoggetclientwidth) | Gets the width of the the dialog's client area. |
| [DialogGetClientHeight](#dialoggetclientheight) | Gets the height of the the dialog's client area. |
| [DialogSetClient](#dialogsetclient) | Adjusts the bounding rectangle of the dialog based on the desired size of the client area. |
| [DialogGetText](#dialoggettext) | Gets the text of the dialog caption. |
| [DialogSetText](#dialogsettext) | Sets the text of the dialog caption. |
| [DialogGetLoc](#dialoggetloc) | Gets the location of the top left corner of the window. |
| [DialogSetLoc](#dialogsetloc) | Sets the location of the top left corner of the window. |
| [DialogGetUser](#dialoggetuser) | Retrieves a value from the user data area of the dialog. |
| [DialogSetUser](#dialogsetuser) | Sets a value in the user data area of a dialog. |
| [CBGetDlgMsgResult](#cbgetdlgmsgresult) | Gets the return value of a message processed in the dialog box procedure. |
| [CBSetDlgMsgResult](#cbsetdlgmsgresult) | Sets the return value of a message processed in the dialog box procedure. |

---

# Control methods

| Name       | Description |
| ---------- | ----------- |
| [ControlAdd](#controladd) | Adds a control to the dialog. |
| [ControlSetFocus](#controlsetfocus) | Sets the focus in the specified control of a dialog box. |
| [ControlDisable](#controldisable) | Disables the specified control. |
| [ControlEnable](#controlenable) | Enables the specified control. |
| [ControlGetLoc](#controlgetloc) | Gets the location of the top left corner of the window. |
| [ControlSetLoc](#controlsetloc) | Sets the location of the top left corner of the window. |
| [ControlGetText](#controlgettext) | Retrieves the text in a control or control caption. |
| [ControlSetText](#sontrolsettext) | Sets the text in a control or control caption. |
| [ControlHandle](#controlhandle) | Returns a window handle for the specified control ID. |
| [ControlHide](#controlhide) | Hides the specified control. |
| [ControlNormalize](#controlmormalize) | Makes visible the specified control. |
| [ControlKill](#controlkill) | Destroys the specified control. |
| [ControlShowState](#controlshowstate) | Changes the visible state of a control. |
| [ControlRedraw](#controlredraw) | Redraws the specified control. |
| [ControlPost](#controlpost) | Posts a message in the message queue. |
| [ControlSend](#controlsend) | Sends the specified message to the specified control. |
| [ControlGetSize](#controlgetsize) | Gets the width and height of the control. |
| [ControlGetWidth](#controlgetwidth) | Gets the width of the control. |
| [ControlGetHeight](#controlgetheight) | Gets the height of the control. |
| [ControlSetSize](#controlsetsize) | Sets the width and height of the specified window. |
| [ControlGetClient](#controlgetclient) | Retrieves the coordinates of the control's client area. |
| [ControlGetClientWidth](#controlgetclientwidth) | Gets the width of the control's client area. |
| [ControlGetClientHeight](#controlgetclientheight) | Gets the height of the control's client area. |
| [ControlSetClient](#controlsetclient) | Adjusts the bounding rectangle of the dialog based on the desired size of the client area. |
| [ControlCenterHoriz](#controlcenterhoriz) | Moves the control to the center of the dialog horizontally |
| [ControlCenterVert](#controlcentervert) | Moves the control to the center of the dialog vertically. |
| [ControlGetCheck](#controlgetcheck) | Gets the check state of a radio button or check box. |
| [ControlSetCheck](#controlsetcheck) | Sets the check state of a radio button or check box. |
| [ControlSetOption](#controlsetoption) | Sets the check state for an Option (radio) control, and unsets the check state for other Option buttons in a group. |

---

# Fonts

| Name       | Description |
| ---------- | ----------- |
| [DialogGetFont](#dialoggetfont) | Gets the handle of the font used by the dialog. |
| [DialogGetFontFaceName](#dialoggetfontfacename) | Gets the face name of the font used by the dialog. |
| [DialogGetFontPointSize](#dialoggetfontpointsize) | Gets the point size of the font used by the dialog. |
| [FontNew](#fontnew) | Creates a logical font. |
| [FontEnd](#fontend) | Destroys a font when it is no longer needed. |
| [ControlGetFont](#controlgetfont) | Gets the handle of the font used by the control. |
| [ControlSetFont](#controlsetfont) | Sets the font that a control is to use when drawing text. |
| [ControlGetFontFaceName](#controlgetfontfacename) | Gets the face name of the font used by the control. |
| [ControlGetFontPointSize](#controlgetfontpointsize) | Gets the point size of the font used by the control. |

---

# Images

| Name       | Description |
| ---------- | ----------- |
| [DialogSetIcon](#dialogseticon) | Changes both the dialog icon in the caption, and the icon shown in the ALT+TAB task list. |
| [DialogSetIconEx](#dialogseticonex) | Sets the big and small icons of the dialog. |
| [FindResourceType](#findresourcetype) | Finds the resource type given it's identifier or name. |
| [ControlSetImage](#controlsetimage) | Changes the icon or bitmap displayed in an image label control. |
| [ControlSetImageX](#controlsetimagex) | Changes the icon or bitmap displayed in an image label control. |
| [ControlSetImgButton](#controlsetimagebutton) | Changes the bitmap displayed in an image button control. |
| [ControlSetImgButtonX](#controlsetimagebuttonx) | Changes the bitmap displayed in an image button control. |

---


# Metric conversions

| Name       | Description |
| ---------- | ----------- |
| [DialogUnitsToPixels](#dialogunitstopixels) | Converts the specified dialog box units to screen units (pixels). |
| [DialogUnitsToPixelsRatios](#dialogunitstopixelsratios) | Retrieves the conversion ratios from dialog units to pixels. |
| [DluToPixRX](#dlutopixrx) | Retrieves the conversion ratios from dialog units to pixels. |
| [DluToPixRX](#dlutopixrx) | Retrieves the conversion ratios from dialog units to pixels. |
| [PixelsToDialogUnits](#pixelstodialogunits) | Converts the specified screen units (pixels) to dialog box units. |
| [PixelsToDialogUnitsRatios](#pixelstodialogunitsratios) | Retrieves the conversion ratios from pixels to dialog units. |
| [PixToDluRX](#pixtodlurx) | Retrieves the conversion ratio from pixels to dialog units. |
| [PixToDluRY](#pixtodlury) | Retrieves the conversion ratio from pixels to dialog units. |

---

# DPI scaling

| Name       | Description |
| ---------- | ----------- |
| [rxRatio](rxratio) | Returns the horizontal DPI scaling ratio. |
| [ryRatio](rxratio) | Returns the vertical DPI scaling ratio. |
| [ScaleX](scalex) | Scales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop. |
| [ScaleY](scaley) | Scales a vertical coordinate according the DPI (dots per pixel) being used by the desktop. |
| [UnscaleX](unscalex) | Unscales an horizontal coordinate according the DPI (dots per pixel) being used by the desktop. |
| [UnscaleY](unscaley) | Unscales a vertical coordinate according the DPI (dots per pixel) being used by the desktop. |
| [ScaleRect](scalerect) | Scales a RECT structure according the DPI (dots per pixel) being used by the desktop. |
| [UnscaleRect](unscalerect) | Unscales a RECT structure according the DPI (dots per pixel) being used by the desktop. |

---

# Layout manager

| Name       | Description |
| ---------- | ----------- |
| [AdjustControls](#adjustcontrols) | Adjust the controls size and/or position. |
| [ControlAnchor](#ControlAnchor) | Anchor the control. |
| [GetAnchorItem](#getanchoritem) | Gets the anchor item. |

---

# Colors

| Name       | Description |
| ---------- | ----------- |
| [ControlSetColor](#controlsetcolor) | Sets the colors of the control. |
| [DialogSetColor](#dialogsetcolor) | Sets the background color of the dialog. |
| [DialogDisableRepaintOnResize](#dialogdisablerepaintonresize) | Enable/disable dialog repainting during resizing. |
| [DialogEnableRepaint](#dialogenablerepaint) | Enable/disable dialog repainting during resizing. |
| [GetColorItem](#getcoloritem) | Gets the color item. |
| [IsDialogRepaintDisabled](#isdialogrepaintdisabled) | Checks if repainting is enabled during resizing. |
| [IsDialogRepaintDisabledOnResize](#isdialogrepaintdisabledonresize) | Checks if repainting is enabled during resizing. |

---

# Keyboard accelerators

| Name       | Description |
| ---------- | ----------- |
| [AccelHandle](#accelhandle) | Gets the accelerator table handle. |
| [AccelAttach](#accelattach) | Attaches an accelerator table handle. |
| [AddAccelerator](#addaccelerator) | Adds an accelerator key to the table. |
| [CreateAcceleratorTable](#createacceleratortable) | Creates the accelerator table. |
| [DestroyAcceleratorTable](#destroyacceleratortable) | Destroys the accelerator table. |

---

# Scrollable dialogs

| Name       | Description |
| ---------- | ----------- |
| [DialogSetViewPort](#dialogsetviewport) | Makes a dialog scrollable by shrinking its client area size. |
| [IsDialogScrollable](#isdialogscrollable) | Flag indicating if a dialog is scrollable. |
| [DialogResetScrollBars](#dialogresetscrollbars) | Resets the dikalog scrolling information. |
| [DialogSetupScrollBars](#dialogsetupscrollbars) | Sets the dialog scroll information. |
| [DialogOnVScroll](#dialogonvscroll) | Handles vertical scrollbar messages. |
| [DialogOnHScroll](#dialogonhscroll) | Handles horizontañ scrollbar messages. |
| [DialogOnSize](#dialogonsize) | Handles WM_SIZE messges. |

---

# Helper function

| Name       | Description |
| ---------- | ----------- |
| [AfxCDialogPtr](#afxcdialogptr) | Returns a pointer to the CDialog class given the handle of the window created with it or the handle of any of it's children. |

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

