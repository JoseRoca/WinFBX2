# Menu Procedures

A *menu* is a list of items that specify options or groups of options (a submenu) for an application. Clicking a menu item opens a submenu or causes the application to carry out a command.

A menu is arranged in a hierarchy. At the top level of the hierarchy is the *menu bar*; which contains a list of *menus*, which in turn can contain *submenus*. A menu bar is sometimes called a top-level menu, and the menus and submenus are also known as pop-up menus.

A menu item can either carry out a command or open a submenu. An item that carries out a command is called a *command item* or a *command*.

An item on the menu bar almost always opens a menu. Menu bars rarely contain command items. A menu opened from the menu bar drops down from the menu bar and is sometimes called a *drop-down menu*. When a drop-down menu is displayed, it is attached to the menu bar. A menu item on the menu bar that opens a drop-down menu is also called a *menu name*.

The menu names on a menu bar represent the main categories of commands that an application provides. Selecting a menu name from the menu bar typically opens a menu whose menu items correspond to the commands in a category. For example, a menu bar might contain a **File** menu name that, when clicked by the user, activates a menu with menu items such as **New**, **Open**, and **Save**.

Only an overlapped or pop-up window can contain a menu bar; a child window cannot contain one. If the window has a title bar, the system positions the menu bar just below it. A menu bar is always visible. A submenu is not visible, however, until the user selects a menu item that activates it.

Each menu must have an owner window. The system sends messages to a menu's owner window when the user selects the menu or chooses an item from the menu.

See more information at [About Menus](https://learn.microsoft.com/en-us/windows/win32/menurc/about-menus)

**Include file**: AfxMenuProcs2.inc.

| Name       | Description |
| ---------- | ----------- |
| [IsMenuHandle](#ismenuhandle) | Determines whether a handle is a menu handle. |
| [IsMenuItemChecked](#ismenuitemchecked) | Returns TRUE if the specified menu item is checked; FALSE otherwise. |
| [IsMenuItemDisabled](#ismenuitemdisabled) | Returns TRUE if the specified menu item is disabled; FALSE otherwise. |
| [IsMenuItemEnabled](#ismenuitemenabled) | Returns TRUE if the specified menu item is enabled; FALSE otherwise. |
| [IsMenuItemGrayed](#ismenuitemgrayed) | Returns TRUE if the specified menu item is grayed; FALSE otherwise. |
| [IsMenuItemHighlighted](#ismenuitemhighlighted) | Returns TRUE if the specified menu item is highlighted; FALSE otherwise. |
| [IsMenuItemOwnerDraw](#ismenuitemownerdraw) | Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise. |
| [IsMenuItemPopUp](#ismenuitempopup) | Returns TRUE if the specified menu item is a submenu; FALSE otherwise. |
| [IsMenuItemSeparator](#ismenuitemseparator) | Returns TRUE if the specified menu item is a separator; FALSE otherwise. |
| [MenuAddPopUp](#menuaddpopup) | Adds a popup child menu to an existing menu. |
| [MenuAddString](#menuaddstring) | Adds a string or separator to an existing menu. |
| [MenuAttach](#menuattach) | Attaches a menu to a window or dialog. |
| [MenuBoldItem](#menubolditem) | Changes the text of a menu item to bold. |
| [MenuCheckItem](#menucheckitem) | Checks a menu item. |
| [MenuCheckRadioButton](#menucheckradiobutton) | Checks a specified menu item and makes it a radio item. |
| [MenuContext](#menucontext) | Creates a floating context menu. |
| [MenuDelete](#menudelete) | Deletes a menu item from an existing menu. |
| [MenuDestroy](#menudestroy) | Destroys the main menu from the window or dialog. |
| [MenuDisableItem](#menudisableitem) | Disables the specified menu item. |
| [MenuDrawBar](#menudrawbar) | Redraws the menu bar of the specified window or dialog. |
| [MenuEnableItem](#menuenableitem) | Enables the specified menu item. |
| [MenuFindItemPosition](#menufinditemposition) | Finds the position of the specified menu item. |
| [MenuGetBarInfo](#menugetbarinfo) | Retrieves information about the specified menu bar. |
| [MenuGetCheckMarkHeight](#menugetcheckmarkheight) | Retrieves the height of the default check-mark bitmap. |
| [MenuGetCheckMarkWidth](#menugetcheckmarkwidth) | Retrieves the width of the default check-mark bitmap. |
| [MenuGetContextHelpID](#menugetcontexthelpid) | Retrieves the Help context identifier associated with the specified menu. |
| [MenuGetDefaultItem](#menugetdefaultitem) | Determines the default menu item on the specified menu. |
| [MenuGetFont](#menugetfont) | Retrieves information about the font used in menu bars. |
| [MenuGetFontPointSize](#menugetfontpointsize) | Retrieves the point size of the font used in menu bars. |
| [MenuGetHandle](#menugethandle) | Retrieves a handle to the menu assigned to the specified window or dialog. |
| [MenuGetItemCount](#menugetitemcount) | Determines the number of items in the specified menu. |
| [MenuGetItemFromPoint](#menugetitemfrompoint) | Determines which menu item, if any, is at the specified location. |
| [MenuGetItemID](#menugetitemid) | Retrieves the menu item ID of a menu item located at the specified position in a menu. |
| [MenuGetRect](#menugetrect) | Calculates the size of a menu bar or a drop-down menu. |
| [MenuGetState](#menugetstate) | Retrieves the state of the specified menu item. |
| [MenuGetSubMenu](#menugetsubmenu) | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
| [MenuGetSubMenusCount](#menugetsubmenuscount) | Get the number of submenus of a menu. |
| [MenuGetSytemMenuHandle](#menugetsytemmenuhandle) | Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying. |
| [MenuGetText](#menugettext) | Retrieves the text of the specified menu item. |
| [MenuGetTextLen](#menugettextlen) | Returns the lengnth of the text of the specified menu item. |
| [MenuGetWindowOwner](#menugetwindowowner) | Retrieves the window owner of the specified menu. |
| [MenuGrayItem](#menugrayitem) | Grays the specified menu item. |
| [MenuHiliteItem](#menuhiliteitem) | Highlights the specified menu item. |
| [MenuItemToggleCheckState](#menuitemtogglecheckstate) | Toggles the checked state of a menu item. |
| [MenuNewBar](#menunewbar) | Creates a new menu bar. |
| [MenuNewPopUp](#menunewpopup) | Creates a new popup menu. |
| [MenuRemoveCloseOption](#menuremovecloseoption) | Removes the system menu close option and disables the X button. |
| [MenuRestoreCloseOption](#menurestorecloseoption) | Restores the system menu close option and enables Alt+F4 and the X button. |
| [MenuRightJustifyItem](#menurightjustifyitem) | Right justifies a top level menu item. |
| [MenuSetContextHelpID](#menusetcontexthelpid) | Associates a Help context identifier with a menu. |
| [MenuSetDefaultItem](#menusetdefaultitem) | Sets the default menu item for the specified menu. |
| [MenuSetDefaultItem](#menusetdefaultitem) | Sets the default menu item for the specified menu. |
| [MenuSetItemBitmaps](#menusetitembitmaps) | Associates the specified bitmap with a menu item. Whether the menu item is selected or clear, the system displays the appropriate bitmap next to the menu item. |
| [MenuSetState](#menusetstate) | Sets the state of the specified menu item. |
| [MenuUnCheckItem](#menuuncheckitem) | Unchecks a menu item. |

---

### <a name="ismenuhandle"></a>IsMenuHandle

Determines whether a handle is a menu handle.
```
FUNCTION IsMenuHandle (BYVAL hMenu AS HMENU) AS BOOLEAN
```
#### Return value

If the function succeeds, the return value is a TRUE. If the function fails, the return value is FALSE.

---

### <a name="menunewbar"></a>MenuNewBar

Creates a menu bar. The menu is initially empty, but it can be filled with menu items by using the **MenuAddPopUp** and **MenuAddString** functions.
```
FUNCTION MenuNewBar () AS HMENU
```
#### Return value

If the function succeeds, the return value is a handle to the newly created menu.

If the function fails, the return value is NULL. To get extended error information, call GetLastError.

#### Remarks

Resources associated with a menu that is assigned to a window are freed automatically. If the menu is not assigned to a window, an application must free system resources associated with the menu before closing. An application frees menu resources by calling the **MenuDestroy** function.

#### Usage example
```
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
```
---
### <a name="menugetdefaultitem"></a>MenuGetDefaultItem

Determines the default menu item on the specified menu.
```
FUNCTION MenuGetDefaultItem (BYVAL hMenu AS HMENU, BYVAL gmdiFlags AS UINT = 0, BYVAL _
   fByPosition AS BOOLEAN = TRUE) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu for which to retrieve the default menu item. |
| *gmdiFlags* | Indicates how the function should search for menu items. This parameter can be zero or more of the following values.<br>GMDI_GOINTOPOPUPS (&H0002): If the default item is one that opens a submenu, the function is to search recursively in the corresponding submenu. If the submenu has no default item, the return value identifies the item that opens the submenu. By default, the function returns the first default item on the specified menu, regardless of whether it is an item that opens a submenu.<br>GMDI_USEDISABLED (&h0001): The function is to return a default item, even if it is disabled. By default, the function skips disabled or grayed items. |
| *fByPosition* | Indicates whether to retrieve the menu item's identifier or its position. If this parameter is FALSE, the identifier is returned. Otherwise, the position is returned. |

#### Return value

If the function succeeds, the return value is the identifier or position of the menu item. If the function fails, the return value is -1. To get extended error information, call **GetLastError**.

#### Usage example
```
MenuGetDefaultItem(hMenu, 1)
```
---

### <a name="menugethandle"></a>MenuGetHandle

Retrieves a handle to the menu assigned to the specified window.
```
FUNCTION MenuGetHandle (BYVAL hwnd AS HWND) AS HMENU
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | A handle to the window whose menu handle is to be retrieved. |

#### Return value

The return value is a handle to the menu. If the specified window has no menu, the return value is NULL. If the window is a child window, the return value is undefined.

#### Remarks

**MenuGetHandle** does not work on floating menu bars. Floating menu bars are custom controls that mimic standard menus; they are not menus. To get the handle on a floating menu bar, use the [Active Accessibility APIs](https://learn.microsoft.com/en-us/previous-versions/ms971350(v=msdn.10)).

---

### <a name="menugetsystemmenuhandle"></a>MenuGetSytemMenuHandle

Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying.
```
FUNCTION MenuGetSytemMenuHandle (BYVAL hwnd AS HWND, BYVAL bRevert AS BOOLEAN = FALSE) AS HMENU
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | A handle to the window that will own a copy of the window menu. |
| *bRevert* | The action to be taken. If this parameter is FALSE, **MenuGetSytemMenuHandle** returns a handle to the copy of the window menu currently in use. The copy is initially identical to the window menu, but it can be modified. If this parameter is TRUE, **MenuGetSytemMenuHandle** resets the window menu back to the default state. The previous window menu, if any, is destroyed. |

#### Return value

If the *bRevert* parameter is FALSE, the return value is a handle to a copy of the window menu. If the *bRevert* parameter is TRUE, the return value is NULL.

#### Remarks

Any window that does not use the **MenuGetSytemMenuHandle** function to make its own copy of the window menu receives the standard window menu.

The window menu initially contains items with various identifier values, such as **SC_CLOSE**, **SC_MOVE**, and **SC_SIZE**.

Menu items on the window menu send **WM_SYSCOMMAND** messages.

All predefined window menu items have identifier numbers greater than &hF000. If an application adds commands to the window menu, it should use identifier numbers less than &hF000.

The system automatically grays items on the standard window menu, depending on the situation. The application can perform its own checking or graying by responding to the **WM_INITMENU** message that is sent before any menu is displayed.

---

### <a name="menugetsubmenu"></a>MenuGetSubMenu

Retrieves a handle to the drop-down menu or submenu activated by the specified menu item.
```
FUNCTION MenuGetSubMenu (BYVAL hMenu AS HMENU, BYVAL nPos AS LONG) AS HMENU
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu. |
| *nPos* | The one-based relative position in the specified menu of an item that activates a drop-down menu or submenu. |

#### Return value

If the function succeeds, the return value is a handle to the drop-down menu or submenu activated by the menu item. If the menu item does not activate a drop-down menu or submenu, the return value is NULL.

---

### <a name="Menugetsubmenuscount"></a>MenuGetSubMenusCount

Retrieves the number of submenus of a menu.
```
FUNCTION MenuGetSubmenusCount (BYVAL hMenu AS HMENU) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu. |

#### Return value

If the function succeeds, the return value is the numbers of submenus of a menu. If the menu item does not have submenus, the return value is zero.

---

### <a name="menugetitemid"></a>MenuGetItemID

Retrieves the menu item identifier of a menu item located at the specified position in a menu.
```
FUNCTION MenuGetItemID (BYVAL hMenu AS HMENU, BYVAL nPos AS LONG) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the item whose identifier is to be retrieved. |
| *nPos* | The one-based relative position of the menu item whose identifier is to be retrieved. |

#### Return value

The return value is the identifier of the specified menu item. If the menu item identifier is NULL or if the specified item opens a submenu, the return value is -1.

---

### <a name="menugetwindowowner"></a>MenuGetWindowOwner

Retrieves the window owner of the menu.
```
FUNCTION MenuGetWindowOwner (BYVAL hMenu AS HMENU) AS HWND
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu. |

#### Return value

The handle of the window that owns the menu.

---

### <a name="menunewpopup"></a>MenuNewPopUp

Creates a drop-down menu, submenu, or shortcut menu. The menu is initially empty. You can insert or append menu items by using the InsertMenuItem function. You can also use the **MenuAddString** function to append menu items.
```
FUNCTION MenuNewPopUp () AS HMENU
```
#### Return value

If the function succeeds, the return value is a handle to the newly created menu.

If the function fails, the return value is NULL. To get extended error information, call GetLastError.

---

### <a name="menuaddpopup"></a>MenuAddPopUp

Adds a popup child menu to an existing menu. A popup menu is a small window that "pops up" when a menu item is highlighted. This allows nesting, and gives the user an opportunity to choose from "sub-menu" items.
```
FUNCTION MenuAddPopUp (BYVAL hMenu AS HMENU, BYREF wszText AS WSTRING, BYVAL hPopUp AS HMENU, _
   BYVAL fState AS UINT, BYVAL item AS LONG = 0, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu in which the new menu item is inserted. |
| *wszText* | Text displayed in the parent menu. An ampersand (&) may be used in the string to make the following letter into a control accelerator (hot-key). The letter appears underscored to signify that it is an accelerator. |
| *hPopUp* | Handle of the child popup menu to be added. |
| *fState* | The initial state of the menu item. It can be one of the following:<br>MFS_DISABLED: Disable the item so that it cannot be selected.<br>MFS_ENABLED: Enable the item so that it can be selected. |
| *item* | The identifier or position of the menu item before which to insert the new item. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | Controls the meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier and the popup menu is inserted prior to it. Otherwise, the popup menu is inserted at the physical position within the parent menu, where position = 1 for the first position, position = 2 for the second, and so on. If position is not specified then the popup menu is appended to the end of the menu. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

---

### <a name="menufinditemposition"></a>MenuFindItemPosition

Finds the position of the specified menu item.
```
FUNCTION MenuFindItemPosition (BYVAL hMenu AS HMENU, BYVAL itemID AS UINT, BYREF itemPos AS LONG) AS BOOLEAN
FUNCTION MenuFindItemPosition OVERLOAD (BYVAL hMenu AS HMENU, BYVAL itemID AS UINT) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu. |
| *itemID* | The item identifier. |
| *itemPos* | A long variable where the item position will be returned. |

#### Return value

First overloaded function: If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

Second overloaded function: If the function succeeds, the return value is the one-based position of the item identifier. If the function fails, it returns 0.

#### Usage examples
```
DIM nPos AS LONG
MenuFindItemPosition(hMenu, ID_EXIT, nPos)
```
```
DIM nPos AS LONG = MenuFindItemPosition(hMenu, ID_EXIT)
```
---

### <a name="menuaddstring"></a>MenuAddString

Adds a string or separator to an existing menu.
```
FUNCTION MenuAddString (BYVAL hMenu AS HMENU, BYREF wszText AS WSTRING, BYVAL id AS LONG, _
   BYVAL fState AS UINT, BYVAL item AS LONG = 0, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu in which the string is added. |
| *wszText* | Text to display in the parent menu. An ampersand (&) may be used in the string to make the following letter into a command accelerator (hot-key). The letter is underscored to signify that it is an accelerator. To create a horizontal separator instead of a text string, set wszText = "-", id = 0, fstate = 0. To include a keyboard accelerator description in a menu string, separate it from the menu item text with a TAB {CHR(9)} character |
| *id* | The unique numeric identifier for the menu item. |
| *fState* | The initial state of the menu item. It can be one or more of the following, combined together with the OR operator to form a bitmask:<br>MFS_CHECKED: Place a checkmark next to the item.<br>MFS_DEFAULT: The default menu item, displayed in bold.  Only one item may be the default.<br>MFS_DISABLED: Disable the menu item so that it cannot be selected.<br>MFS_ENABLED:Enable the menu item so that it can be selected.<br>MFS_GRAYED: Disable the menu item so that it cannot be selected, and draw it in a "grayed" state to indicate this.<br>MFS_HILITE: Highlight the menu item.<br>MFS_UNCHECKED:Do not place a checkmark next to the item.<br>MFS_UNHILITE: Item is not highlighted.<br>A state value of zero (0) provides MFS_ENABLED OR MFS_UNCHECKED OR MFS_UNHILITE. |
| *item* | The menu item to be added, as determined by the *fByPosition* parameter. If the *fByPosition* parameter is FALSE, the menu item is inserted prior to the menu item ID specified by *item*. Otherwise, the menu item is inserted at the physical position within the parent menu, where position = 1 for the first position, position = 2 for the second, and so on. If *item* is not specified then the menu item appended to the end of the menu. |
| *fByPosition* | Controls the meaning of *position*. If this parameter is FALSE, *position* is a menu item identifier; otherwise, it is a physical position within the parent menu. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

---

### <a name="menuattach"></a>MenuAttach

Attaches a menu to a window or dialog.
```
FUNCTION MenuAttach (BYVAL hMenu AS HMENU, BYVAL hwnd AS HWND) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the new menu. If this parameter is NULL, the window's current menu is removed. |
| *hwnd* | A handle to the window to which the menu is to be assigned. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menucheckradiobutton"></a>MenuCheckRadioButton

Checks a specified menu item and makes it a radio item. At the same time, the function clears all other menu items in the associated group and clears the radio-item type flag for those items.

```
FUNCTION MenuCheckRadioButton (BYVAL hMenu AS HMENU, BYVAL first AS LONG, BYVAL last AS LONG, _
   BYVAL check AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the group of menu items. |
| *first* | The identifier or position of the first menu item in the group. |
| *last* | The identifier or position of the last menu item in the group. |
| *check* | The identifier or position of the menu item to check. |
| *fByPosition* | Controls the meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, *item* is the position of the menu item, where position = 1 for the first position, position = 2 for the second, and so on.  |

#### Usage example:
```
MenuCheckRadioButton(hMenu, ID_OPEN, ID_EXIT, ID_EXIT)      ' By item identifier
MenuCheckRadioButton(GetSubMenu(hMenu, 0), 1, 2, 2, TRUE)   ' By position
```
#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menucontext"></a>MenuContext

Creates a floating context menu. A context menu is a floating popup menu which is shown until the user makes a selection or dismisses it.  It can appear anywhere on the screen.

```
FUNCTION MenuContext (BYVAL hDlg AS HWND, BYVAL hPopUpMenu AS HMENU, _
   BYVAL x AS LONG, BYVAL y AS LONG, BYVAL flags AS UINT) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hDLg* | A handle to the dialog. |
| *hPopUpMenu* | Handle of a menu created with **MenuNewPopup**. |
| *fByPosition* | Controls the meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, *item* is the position of the menu item, where position = 1 for the first position, position = 2 for the second, and so on.  |
| *x, y* | Location of the context menu, in pixels, relative to the upper left corner of the desktop. |
| *flags* | Use zero of more of these flags to specify function options. |

Use one of the following flags to specify how the function positions the shortcut menu horizontally.

| Value | Meaning |
| ------| ------- |
| **TPM_CENTERALIGN** | Centers the shortcut menu horizontally relative to the coordinate specified by the x parameter. |
| **TPM_LEFTALIGN** | Positions the shortcut menu so that its left side is aligned with the coordinate specified by the x parameter. |
| **TPM_RIGHTALIGN** | Positions the shortcut menu so that its right side is aligned with the coordinate specified by the x parameter. |

Use one of the following flags to specify how the function positions the shortcut menu vertically.

| Value | Meaning |
| ------| ------- |
| **TPM_BOTTOMALIGN** | Positions the shortcut menu so that its bottom side is aligned with the coordinate specified by the y parameter. |
| **TPM_TOPALIGN** | Positions the shortcut menu so that its top side is aligned with the coordinate specified by the y parameter. |
| **TPM_VCENTERALIGN** | Centers the shortcut menu vertically relative to the coordinate specified by the y parameter. |

Use one of the following flags to specify which mouse button the shortcut menu tracks.

| Value | Meaning |
| ------| ------- |
| **TPM_LEFTBUTTON** | The user can select menu items with only the left mouse button. |
| **TPM_RIGHTBUTTON** | The user can select menu items with both the left and right mouse buttons. |
 
---

### <a name="menudelete"></a>MenuDelete

Deletes a menu item or detaches a submenu from the specified menu. If the menu item opens a drop-down menu or submenu, **MenuDelete** does not destroy the menu or its handle, allowing the menu to be reused. Before this function is called, the **MenuGetSubMenu** function should retrieve a handle to the drop-down menu or submenu.
```
FUNCTION MenuDelete (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu to be changed. |
| *item* | The menu item to be deleted, as determined by the *fByPosition* parameter. |
| *fByPosition* | Controls the meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, *item* is the position of the menu item, where position = 1 for the first position, position = 2 for the second, and so on.  |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

#### Remarks

The application must call the **MenuDrawBar** function whenever a menu changes, whether the menu is in a displayed window.

---

### <a name="menudestroy"></a>MenuDestroy

Destroys the specified menu and frees any memory that the menu occupies.
```
FUNCTION MenuDestroy (BYVAL hMenu AS HMENU) AS BOOLEAN
FUNCTION MenuDestroy (BYVAL hwnd AS HWND) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu to be destroyed. |
| *hwnd* | A handle to the window or dialog to which the menu is attached. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

#### Remarks

Before closing, an application must use the **MenuDestroy** function to destroy a menu not assigned to a window. A menu that is assigned to a window is automatically destroyed when the application closes.

**MenuDestroy** is recursive, that is, it will destroy the menu and all its submenus.

---

### <a name="menudrawbar"></a>MenuDrawBar

Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
```
FUNCTION MenuDrawBar (BYVAL hwnd AS HWND) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | A handle to the window whose menu bar is to be redrawn. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugetbarinfo"></a>MenuGetBarInfo

Retrieves information about the specified menu bar.
```
FUNCTION MenuGetBarInfo (BYVAL hwnd AS HWND, BYVAL idObject AS LONG, BYVAL idItem AS LONG) AS MENUBARINFO
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | A handle to the window or dialog whose information is to be retrieved. |
| *idObject* | The menu object. This parameter can be one of the following values:<br>OBJID_CLIENT: &hFFFFFFFC - The popup menu associated with the window.<br>OBJID_MENU: &hFFFFFFFD - The menu bar associated with the window<br>OBJID_SYSMENU: &hFFFFFFFF - The system menu associated with the window |
| *idItem* | The item for which to retrieve information. If this parameter is zero, the function retrieves information about the menu itself. If this parameter is 1, the function retrieves information about the first item on the menu, and so on. |

#### Return value

A **MENUBARINFO** structure.

#### Result code

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugetcheckmarkheight"></a>MenuGetCheckMarkHeight

Retrieves the height of the default check-mark bitmap.
```
FUNCTION FUNCTION MenuGetCheckMarkHeight () AS LONG
```

#### Remarks

The system displays this bitmap next to selected menu items. Before calling the **MenuSetItemBitmaps** function to replace the default check-mark bitmap for a menu item, an application must determine the correct bitmap size by calling **MenuGetCheckMarkWidth** and **MenuGetCheckMarkHeight**.

---

### <a name="menugetcheckmarkwidth"></a>MenuGetCheckMarkWidth

Retrieves the width of the default check-mark bitmap.
```
FUNCTION FUNCTION MenuGetCheckMarkWidth () AS LONG
```

#### Remarks

The system displays this bitmap next to selected menu items. Before calling the **MenuSetItemBitmaps** function to replace the default check-mark bitmap for a menu item, an application must determine the correct bitmap size by calling **MenuGetCheckMarkWidth** and **MenuGetCheckMarkHeight**.

---

### <a name="menugetitemcount"></a>MenuGetItemCount

Determines the number of items in the specified menu.
```
FUNCTION MenuGetItemCount (BYVAL hMenu AS HMENU) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu to be examined. |

#### Return value

If the function succeeds, the return value specifies the number of items in the menu.

If the function fails, the return value is -1. To get extended error information, call **GetLastError**.

---

### <a name="menugetitemfrompoint"></a>MenuGetItemFromPoint

Determines which menu item, if any, is at the specified location.
```
FUNCTION MenuGetItemFromPoint (NYVAL hwnd AS HWND, BYVAL hMenu AS HMENU, BYVAL ptScreen AS LONG) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | A handle to the window containing the menu. If this value is NULL and the *hMenu* parameter represents a popup menu, the function will find the menu window. |
| *hMenu* | A handle to the menu containing the menu items to hit test. |
| *ptScreen* | A structure that specifies the location to test. If *hMenu* specifies a menu bar, this parameter is in window coordinates. Otherwise, it is in client coordinates. |

#### Return value

Returns the zero-based position of the menu item at the specified location or -1 if no menu item is at the specified location.

---

### <a name="menugettextlen"></a>MenuGetTextLen

Returns the lengnth of the text of the specified menu item.
```
FUNCTION MenuGetTextLen (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugettext"></a>MenuGetText

Retrieves the text of the specified menu item.
```
FUNCTION MenuGetText (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS DWSTRING
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is the retrieved text. If the function fails, the return value is an empty string.

---

### <a name="menusettext"></a>MenuSetText

Sets the text of the specified menu item.

```
FUNCTION MenuSetText (BYVAL hMenu AS HMENU, BYVAL item AS UINT, BYREF wszText AS WSTRING, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *wszText* | The text to set. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugetstate"></a>MenuGetState

Retrieves the state of the specified menu item.

```
FUNCTION MenuGetState (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS UINT
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Zero on failure or one or more of the following values:

| Value | Meaning |
| ----- | ------- |
| **MFS_CHECKED** | The item is checked. |
| **MFS_DEFAULT** | The item is the default. |
| **MFS_DISABLED** | The item is disabled. |
| **MFS_ENABLED** | The item is enabled. |
| **MFS_GRAYED** | The item is grayed. |
| **MFS_HILITE** | The item is highlighted. |
| **MFS_UNCHECKED** | The item is unchecked. |
| **MFS_UNHILITE** | The item is not highlighted. |

Note: To get extended error information, use the **GetLastError** function.

---

### <a name="menuitemtogglecheckstate"></a>MenuItemToggleCheckState

Toggles the checked state of a menu item.

```
FUNCTION MenuItemToggleCheckState (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS UINT
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="ismenuitemchecked"></a>IsMenuItemChecked

Checks if a menu item state is checked.
```
FUNCTION IsMenuItemChecked (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is checked; FALSE otherwise.

---

### <a name="ismenuitemenabled"></a>IsMenuItemEnabled

Checks if a menu item state is enabled.
```
FUNCTION IsMenuItemEnabled (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is enabled; FALSE otherwise.

---

### <a name="ismenuitemdisabled"></a>IsMenuItemDisabled

Checks if a menu item state is disabled.
```
FUNCTION IsMenuItemDisabled (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is disabled; FALSE otherwise.

---

### <a name="ismenuitemgrayed"></a>IsMenuItemGrayed

Checks if a menu item state is gayed.
```
FUNCTION IsMenuItemGrayed (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is grayed; FALSE otherwise.

---

### <a name="ismenuitemhighlighted"></a>IsMenuItemHghlighted

Checks if a menu item state is highlighted.
```
FUNCTION IsMenuItemHighlighted (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is highlighted; FALSE otherwise.

---

### <a name="ismenuitemseparator"></a>IsMenuItemSeparator

Checks if a menu item state is a separator.
```
FUNCTION IsMenuItemSeparator (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is a separator; FALSE otherwise.

---

### <a name="ismenuitempopup"></a>IsMenuItemPopUp

Checks if a menu item state is a popup.
```
FUNCTION IsMenuItemPopUp (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is a popup; FALSE otherwise.

---

### <a name="ismenuitemownerdraw"></a>IsMenuItemOwnerDraw

Checks if a menu item state is ownerdras.
```
FUNCTION IsMenuItemOwnerDraw (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise.

---

### <a name="menubolditem"></a>MenuBoldItem

Changes the text of a menu item to bold.
```
FUNCTION MenuBoldItem (BYVAL hMenu AS HMENU, BYVAL item AS UINT, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugetsefaultitem"></a>MenuSetDefaultItem

Sets the default menu item for the specified menu.
```
MenuSetDefaultItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = TRUE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu for which to set the default menu item. |
| *item* | The identifier or position of the new default menu item or -1 for no default item. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of item. If this parameter is FALSE, item is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Usage example:
```
MenuSetDefaultItem(hMenu, 1)
```
---
### <a name="menusetitembitmaps"></a>MenuSetItemBitmaps

Associates the specified bitmap with a menu item. Whether the menu item is selected or clear, the system displays the appropriate bitmap next to the menu item.
```
FUNCTION MenuSetItemBitmaps (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL hBitmapUnchecked AS HBITMAP, _
   BYVAL hBitmapChecked AS HBITMAP, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu containing the item to receive new check-mark bitmaps. |
| *item* | The menu item to be changed, as determined by the *fByPosition* parameter. |
| *fByPosition* | The meaning of item. If this parameter is FALSE, item is a menu item identifier. Otherwise, it is a menu item position. |
| *hBitmapUnchecked* | A handle to the bitmap displayed when the menu item is not selected. |
| *hBitmapChecked* | A handle to the bitmap displayed when the menu item is selected. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

#### Remarks

If either the *hBitmapUnchecked* or *hBitmapChecked* parameter is NULL, the system displays nothing next to the menu item for the corresponding check state. If both parameters are NULL, the system displays the default check-mark bitmap when the item is selected, and removes the bitmap when the item is not selected.

When the menu is destroyed, these bitmaps are not destroyed; it is up to the application to destroy them.

The selected and clear bitmaps should be monochrome. The system uses the Boolean AND operator to combine bitmaps with the menu so that the white part becomes transparent and the black part becomes the menu-item color. If you use color bitmaps, the results may be undesirable.

Use the **GetSystemMetrics** function with the **SM_CXMENUCHECK** and **SM_CYMENUCHECK** values to retrieve the bitmap dimensions.

#### Usage example
```
MenuCheckItem(hMenu, ID_OPEN)
MenuSetItemBitmaps(hMenu, ID_OPEN, NULL, NULL)
```
---

### <a name="menusetstate"></a>MenuSetState

Sets the state of the specified menu item.
```
FUNCTION MenuSetState (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fState AS UINT, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fState* | The menu item state. It can be one or more of these values (see table below): |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

| fState | Meaning |
| ----- | ------- |
| **MFS_CHECKED** | The item is checked. |
| **MFS_DEFAULT** | The item is the default. |
| **MFS_DISABLED** | The item is disabled. |
| **MFS_ENABLED** | The item is enabled. |
| **MFS_GRAYED** | The item is grayed. |
| **MFS_HILITE** | The item is highlighted. |
| **MFS_UNCHECKED** | The item is unchecked. |
| **MFS_UNHILITE** | The item is not highlighted. |

#### Return value

Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise.

---

### <a name="menucheckitem"></a>MenuCheckItem

Checks a menu item.
```
FUNCTION MenuCheckItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menuuncheckitem"></a>MenuUnCheckItem

Unchecks a menu item.
```
FUNCTION MenuUnCheckItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menuenableitem"></a>MenuEnableItem

Enables the specified menu item.
```
FUNCTION MenuEnableItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menudisableitem"></a>MenuDisableItem

Disables the specified menu item.
```
FUNCTION MenuDisableItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugrayitem"></a>MenuGrayItem

Grays the specified menu item.
```
FUNCTION MenuGrayItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menuhiliteitem"></a>MenuHiliteItem

Grays the specified menu item.
```
FUNCTION MenuHiliteItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menurightjustifyitem"></a>MenuRightJustiyItem

Grays the specified menu item.
```
FUNCTION MenuHiliteItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to set information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

---

### <a name="menugetfont"></a>MenuGetFont

Retrieves information about the font used in menu bars.
```
FUNCTION MenuGetFont () AS LOGFONTW
```

#### Return value

A [LOGFONTW](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logfontw) structure containing the information.

---

### <a name="menugetfontpointsize"></a>MenuGetFontPointSize

Retrieves information about the font used in menu bars.
```
FUNCTION MenuGetFontPointSize () AS LONG
```

#### Return value

If the function fails, the return value is 0.

---

### <a name="menugetrect"></a>MenuGetRect

Calculates the size of a menu bar or a drop-down menu.
```
FUNCTION MenuGetRect (BYVAL hwnd AS HWND, BYVAL hmenu AS HMENU, BYVAL prcmenu AS RECT PTR) AS LONG
FUNCTION MenuGetRect (BYVAL hwnd AS HWND, BYVAL hmenu AS HMENU) AS RECT
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | Handle of the window that owns the menu. |
| *hMenu* | A handle to the menu. |
| *prcmenu* | Pointer to a variable of type **RECT** where to return the retrieved values.. |

#### Return value

First overloaded function: If the function succeeds, the return value is 0. If the function fails, the return value is a system error code.

Second overloaded function: Returns a **RECT** structure.

---

### <a name="menugetcontexthelpid"></a>MenuGetContextHelpId

Retrieves the Help context identifier associated with the specified menu.
```
FUNCTION MenuGetContextHelpId (BYVAL hMenu AS HMENU) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu for which the Help context identifier is to be retrieved. |

#### Return value

Returns the Help context identifier if the menu has one, or zero otherwise.

---

### <a name="menusetcontexthelpid"></a>MenuSetContextHelpId

Associates a Help context identifier with a menu.
```
FUNCTION MenuSetContextHelpId (BYVAL hMenu AS HMENU, BYVAL helpID AS DWORD) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu with which to associate the Help context identifier. |
| *helpID* | The help context identifier. |

#### Return value

Returns TRUE if successful, or FALSE otherwise. To retrieve extended error information, call **GetLastError**.

#### Remarks

All items in the menu share this identifier. Help context identifiers can't be attached to individual menu items.

---

### <a name="menuremovecloseoption"></a>MenuRemoveCloseOptiom

Removes the system menu close option and disables the X button.
```
FUNCTION MenuRemoveCloseOptiom (BYVAL hwnd AS HWND) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | Handle of the window or dialog that owns the menu. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

#### Remarks

To restore the close menu, call **MenuRestoreCloseOption**.

The function calls **MenuDrawBar** to reflect the changes.

---

### <a name="menurestorecloseoptiom"></a>MenuRestoreCloseOption

Restores the system menu close option and enables Alt+F4 and the X button.
```
FUNCTION MenuRestoreCloseOptiom (BYVAL hwnd AS HWND) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hwnd* | Handle of the window or dialog that owns the menu. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

#### Remarks

The function calls **MenuDrawBar** to reflect the changes.

To remove the system menu close option and disable Alt+F4 and the X button call **MenuRemoveCloseOptiom**.

---

### <a name="menuaddbitmaptoitem"></a>MenuAddBitmapToItem

Restores the system menu close option and enables Alt+F4 and the X button.
```
FUNCTION MenuAddBitmapToItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN, BYVAL hbmp AS HBITMAP) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | Handle of the window or dialog that owns the menu. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |
| *hbmp* | The bitmap handle. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

---

### <a name="menuaddicontoitem"></a>MenuAddIconToItem

Converts an icon to a bitmap and adds it to the specified hbmpItem field of HMENU item. The caller is responsible for destroying the bitmap generated. The icon will be destroyed if **fAutoDestroy** is set to true. The **hbmpItem** field of the menu item can be used to keep track of the bitmap by passing NULL to the **phbmp** parameter.

```
FUNCTION MenuAddIconToItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN, _
   BYVAL hIcon AS HICON, BYVAL fAutoDestroy AS BOOLEAN = TRUE, BYVAL phbmp AS HBITMAP PTR = NULL) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | Handle of the window or dialog that owns the menu. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *item*. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |
| *hIcon* | The icon handle. |
| *phbmp* | Location where the bitmap representation of the icon is stored. Can be NULL. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

---
