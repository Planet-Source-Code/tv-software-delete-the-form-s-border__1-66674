VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Border_                 As Boolean
Dim Sysmenu_                As Boolean
Dim Minimize_               As Boolean
Dim Maximize_               As Boolean
Dim Sizebox_                As Boolean
Dim Popup_                  As Boolean
Dim hWnd_                   As Long
Dim Wstyle                  As Long

Public Property Let hwnd(ByVal hWnd__ As Long)
hWnd_ = hWnd__
Update
End Property

Public Property Get hwnd() As Long
hwnd = hWnd_
End Property

Public Property Let Border(ByVal Border__ As Boolean)
Border_ = Border__
Update
End Property
Public Property Get Border() As Boolean
Border = Border_
End Property

Public Property Let Sysmenu(ByVal Sysmenu__ As Boolean)
Sysmenu_ = Sysmenu__
Update
End Property
Public Property Get Sysmenu() As Boolean
Sysmenu = Sysmenu_
End Property

Public Property Let Minimize(ByVal Minimize__ As Boolean)
Minimize_ = Minimize__
Update
End Property
Public Property Get Minimize() As Boolean
Minimize = Minimize_
End Property

Public Property Let Maximize(ByVal Maximize__ As Boolean)
Maximize_ = Maximize__
Update
End Property
Public Property Get Maximize() As Boolean
Maximize = Maximize_
End Property

Public Property Let Sizebox(ByVal Sizebox__ As Boolean)
Sizebox_ = Sizebox__
Update
End Property
Public Property Get Sizebox() As Boolean
Sizebox = Sizebox_
End Property

Public Property Let Popup(ByVal Popup__ As Boolean)
Popup_ = Popup__
Update
End Property
Public Property Get Popup() As Boolean
Popup = Popup_
End Property

Private Function Update()
Wstyle = WS_OVERLAPPED Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS
If Border_ = True Then Wstyle = Wstyle Or WS_BORDER Or WS_CAPTION
If Sysmenu_ = True Then Wstyle = Wstyle Or WS_SYSMENU
If Minimize_ = True Then Wstyle = Wstyle Or WS_MINIMIZEBOX
If Maximize_ = True Then Wstyle = Wstyle Or WS_MAXIMIZEBOX
If Sizebox_ = True Then Wstyle = Wstyle Or WS_SIZEBOX
If Popup = True Then Wstyle = Wstyle Or WS_POPUP
Update_Style
End Function

Public Function Update_Style()
If hwnd <> 0 Then
SetWindowLong hwnd, GWL_STYLE, Wstyle
ShowWindow hwnd, 1
End If
End Function
