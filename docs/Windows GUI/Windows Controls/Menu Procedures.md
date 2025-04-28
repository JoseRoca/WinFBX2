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

#### Example

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
```
---

