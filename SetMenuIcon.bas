Attribute VB_Name = "SetMenuIconModule"
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Function SetMenuIcon(FrmHwnd As Long, MainMenuNumber As Long, MenuItemNumber As Long, Flags As Long, BitmapUncheckedHandle As Long, BitmapCheckedHandle As Long)
    On Error Resume Next
    Dim lngMenu As Long
    Dim lngSubMenu As Long
    Dim lngMenuItemID As Long
    lngMenu = GetMenu(FrmHwnd)
    lngSubMenu = GetSubMenu(lngMenu, MainMenuNumber)
    lngMenuItemID = GetMenuItemID(lngSubMenu, MenuItemNumber)
    SetMenuIcon = SetMenuItemBitmaps(lngMenu, lngMenuItemID, Flags, BitmapUncheckedHandle, BitmapCheckedHandle)
End Function

