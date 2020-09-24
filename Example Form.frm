VERSION 5.00
Begin VB.Form frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Current Style"
      Height          =   2895
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   2295
      Begin VB.Label lbl1 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sysmenu"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Popup"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MaximizeBox"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MinimizeBox"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sizebox"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Border"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Wstyle As WindowStyle

Private Sub Command1_Click()
Wstyle.Border = Not Wstyle.Border
Update_Text
End Sub

Private Sub Command2_Click()
Wstyle.Sizebox = Not Wstyle.Sizebox
Update_Text
End Sub

Private Sub Command3_Click()
Wstyle.Minimize = Not Wstyle.Minimize
Update_Text
End Sub

Private Sub Command4_Click()
Wstyle.Maximize = Not Wstyle.Maximize
Update_Text
End Sub

Private Sub Command5_Click()
Wstyle.Popup = Not Wstyle.Popup
Update_Text
End Sub

Private Sub Command6_Click()
Wstyle.Sysmenu = Not Wstyle.Sysmenu
Update_Text
End Sub

Private Sub Form_Load()
Set Wstyle = New WindowStyle
Wstyle.Maximize = True
Wstyle.Minimize = True
Wstyle.Border = True
Wstyle.Sizebox = True
Wstyle.Popup = False
Wstyle.Sysmenu = True
Wstyle.hwnd = Me.hwnd
Update_Text
End Sub

Public Function Update_Text()
lbl1.Caption = "Border = " & IIf(Wstyle.Border, "On", "Off") & vbNewLine & _
                "Popup = " & IIf(Wstyle.Popup, "On", "Off") & vbNewLine & _
                "SizeBox = " & IIf(Wstyle.Sizebox, "On", "Off") & vbNewLine & _
                "MinimizeBox = " & IIf(Wstyle.Minimize, "On", "Off") & vbNewLine & _
                "MaximizeBox = " & IIf(Wstyle.Maximize, "On", "Off") & vbNewLine & _
                "Sysmenu = " & IIf(Wstyle.Sysmenu, "On", "Off") & vbNewLine
End Function

