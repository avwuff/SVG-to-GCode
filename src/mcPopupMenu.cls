VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mcPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
' Exposed Enumeration
Public Enum mceItemStates
    mceDisabled = 1
    mceGrayed = 2
End Enum

' Property variables
Private psCaption As String ' Caption of menu item (with the arrow >) if this is submenu
Private piHwnd As Long ' Handle to Menu

' Supporting API code
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_RIGHTBUTTON = &H2

Private Type POINT
    x As Long
    y As Long
End Type
Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, lpNewItem As String) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal uFlags As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal Hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Property Let Caption(ByVal sCaption As String)

    psCaption = sCaption
   
End Property

Public Property Get Caption() As String

    Caption = psCaption
   
End Property


Public Sub Remove(ByVal iMenuPosition As Long)
   
    DeleteMenu piHwnd, iMenuPosition, MF_BYPOSITION
   
End Sub

Private Sub Class_Initialize()
    piHwnd = CreatePopupMenu()
End Sub

Private Sub Class_Terminate()
    DestroyMenu piHwnd
End Sub

Public Property Get Hwnd() As Long
   
    Hwnd = piHwnd

End Property

Public Sub Add(ByVal iMenuID As Long, vMenuItem As Variant, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional eItemState As mceItemStates, Optional ByVal imgUnchecked As Long = 0, Optional ByVal imgChecked As Long = 0)
   
    ' Check to see if it's a menu item (a string) or a submenu (a class).
    If typeName(vMenuItem) = "String" Then
       
        If vMenuItem = "-" Then ' Make a seperator
            AppendMenu piHwnd, MF_STRING Or MF_SEPARATOR, iMenuID, ByVal vbNullString
        Else
            AppendMenu piHwnd, MF_STRING Or -bChecked * MF_CHECKED, iMenuID, ByVal vMenuItem
        End If
   
        ' Menu Icons
        If imgChecked = 0 Then imgChecked = imgUnchecked ' Need a value for both
        SetMenuItemBitmaps piHwnd, iMenuID, MF_BYCOMMAND, imgUnchecked, imgChecked
       
        ' Default item
        If bDefault Then SetMenuDefaultItem piHwnd, iMenuID, 0
        ' Disabled (Regular color text)
        If eItemState = mceDisabled Then EnableMenuItem piHwnd, iMenuID, MF_BYCOMMAND Or MF_DISABLED
        ' Disabled (disabled color text)
        If eItemState = mceGrayed Then EnableMenuItem piHwnd, iMenuID, MF_BYCOMMAND Or MF_GRAYED
   
    ' Add a submenu
    ElseIf TypeOf vMenuItem Is mcPopupMenu Then
        Dim oSubmenu As mcPopupMenu: Set oSubmenu = vMenuItem
        AppendMenu piHwnd, MF_STRING Or MF_POPUP Or -bChecked * MF_CHECKED, oSubmenu.Hwnd, ByVal oSubmenu.Caption
        
        If imgChecked = 0 Then imgChecked = imgUnchecked ' Need a value for both
        SetMenuItemBitmaps piHwnd, oSubmenu.Hwnd, MF_BYCOMMAND, imgUnchecked, imgChecked
        
        ' Default item
        If bDefault Then SetMenuDefaultItem piHwnd, oSubmenu.Hwnd, 0
        ' Disabled (Regular color text)
        If eItemState = mceDisabled Then EnableMenuItem piHwnd, oSubmenu.Hwnd, MF_BYCOMMAND Or MF_DISABLED
        ' Disabled (disabled color text)
        If eItemState = mceGrayed Then EnableMenuItem piHwnd, oSubmenu.Hwnd, MF_BYCOMMAND Or MF_GRAYED
        
        
        Set oSubmenu = Nothing
    End If

End Sub

Public Function Show(Optional ByVal iFormHwnd As Long = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional ByVal iControlHwnd As Long = -1) As Long
Dim iHwnd As Long, iX As Long, iY As Long
   
    ' If no form is passed, use the current window
    If iFormHwnd = -1 Or iFormHwnd = 0 Then
       
        Dim iDesktopHwnd As Long, iChildHwnd As Long, iCurrentID As Long, iChildID As Long
       
        iDesktopHwnd = GetDesktopWindow()
        iChildHwnd = GetWindow(iDesktopHwnd, GW_CHILD)
        iCurrentID = GetCurrentProcessId()
        Do While iChildHwnd
            GetWindowThreadProcessId iChildHwnd, iChildID
            If iChildID = iCurrentID Then Exit Do ' Snagged
            iChildHwnd = GetWindow(iChildHwnd, GW_HWNDNEXT)
        Loop
       
        If iChildHwnd = 0 Then ' Can't resolve a form handle. Bail out.
            Show = -1
            Exit Function
        End If
        iHwnd = iChildHwnd
    Else
        iHwnd = iFormHwnd
    End If
   
    ' If passed a control handle, left-bottom orient to the control.
    If iControlHwnd <> -1 Then
        Dim rt As RECT
        GetWindowRect iControlHwnd, rt
        iX = rt.left
        iY = rt.Bottom
    Else
        Dim pt As POINT
        GetCursorPos pt
        If x = -1 Then iX = pt.x Else: iX = x
        If y = -1 Then iY = pt.y Else: iY = y
    End If
    Show = TrackPopupMenuEx(piHwnd, TPM_RETURNCMD Or TPM_RIGHTBUTTON, iX, iY, iHwnd, ByVal 0&)
   
End Function

