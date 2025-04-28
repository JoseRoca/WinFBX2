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
| [MenuCheckItem](#menucheckitem) | Checks a menu item. |
| [MenuDelete](#menudelete) | Deletes a menu item from an existing menu. |
| [MenuDestroy](#menudestroy) | Destroys the main menu from the window or dialog. |
| [MenuDisableItem](#menudisableitem) | Disables the specified menu item. |
| [MenuDrawBar](#menudrawbar) | Redraws the menu bar of the specified window or dialog. |
| [MenuEnableItem](#menuenableitem) | Enables the specified menu item. |
| [MenuFindItemPosition](#menufinditemposition) | Finds the position of the specified menu item. |
| [MenuGetFont](#menugetfont) | Retrieves information about the font used in menu bars. |
| [MenuGetFontPointSize](#menugetfontpointsize) | Retrieves the point size of the font used in menu bars. |
| [MenuGetHandle](#menugethandle) | Retrieves a handle to the menu assigned to the specified window or dialog. |
| [MenuGetItemID](#menugetitemid) | Retrieves the menu item ID of a menu item located at the specified position in a menu. |
| [MenuGetRect](#menugetrect) | Calculates the size of a menu bar or a drop-down menu. |
| [MenuGetState](#menugetstate) | Retrieves the state of the specified menu item. |
| [MenuGetSubmenusCount](#menugetsubmenuscount) | Get the number of submenus of a menu. |
| [MenuGetSubMenu](#menugetsubmenu) | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
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
| [MenuSetItemBold](#menusetitembold) | Changes the text of a menu item to bold. |
| [MenuSetText](#menusettext) | Sets the text of the specified menu item. |
| [MenuSetState](#menusetstate) | Sets the state of the specified menu item. |
| [MenuUnCheckItem](#menuuncheckitem) | Unchecks a menu item. |

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

### <a name="Menugetsubmenuscount"></a>MenuGetSubmenusCount

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

### <a name="menuaddstring"></a>MenuAddString

Adds a string or separator to an existing menu.
```
FUNCTION MenuAddString (BYVAL hMenu AS HMENU, BYREF wszText AS WSTRING, BYVAL id AS LONG, _
   BYVAL fState AS UINT, BYVAL position AS LONG = 0, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *hMenu* | A handle to the menu in which the string is added. |
| *wszText* | Text to display in the parent menu. An ampersand (&) may be used in the string to make the following letter into a command accelerator (hot-key). The letter is underscored to signify that it is an accelerator. To create a horizontal separator instead of a text string, set wszText = "-", id = 0, fstate = 0. To include a keyboard accelerator description in a menu string, separate it from the menu item text with a TAB {CHR(9)} character |
| *id* | The unique numeric identifier for the menu item. |
| *fState* | The initial state of the menu item. It can be one or more of the following, combined together with the OR operator to form a bitmask:<br>MFS_CHECKED: Place a checkmark next to the item.<br>MFS_DEFAULT: The default menu item, displayed in bold.  Only one item may be the default.<br>MFS_DISABLED: Disable the menu item so that it cannot be selected.<br>MFS_ENABLED:Enable the menu item so that it can be selected.<br>MFS_GRAYED: Disable the menu item so that it cannot be selected, and draw it in a "grayed" state to indicate this.<br>MFS_HILITE: Highlight the menu item.<br>MFS_UNCHECKED:Do not place a checkmark next to the item.<br>MFS_UNHILITE: Item is not highlighted.<br>A state value of zero (0) provides MFS_ENABLED OR MFS_UNCHECKED OR MFS_UNHILITE. |
| *position* | Optional position in the parent menu, where the menu item should be inserted. If the *fByPosition* parameter is FALSE, the menu item is inserted prior to the menu item ID specified by *position*. Otherwise, the menu item is inserted at the physical *position* within the parent menu, where position = 1 for the first position, position = 2 for the second, and so on. If *position* is not specified then the popup menu is appended to the end of the menu. |
| *fByPosition* | Controls the meaning of *position". If this parameter is FALSE, *position* is a menu item identifier; otherwise, it is a physical position within the parent menu. |

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

