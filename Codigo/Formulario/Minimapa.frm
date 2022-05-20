VERSION 5.00
Begin VB.Form verpicture 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MiniMap"
   ClientHeight    =   1995
   ClientLeft      =   8745
   ClientTop       =   1095
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1995
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtViewMetric 
      Height          =   375
      Index           =   0
      Left            =   8190
      TabIndex        =   6
      Text            =   "0"
      Top             =   45
      Width           =   735
   End
   Begin VB.TextBox txtViewMetric 
      Height          =   375
      Index           =   1
      Left            =   8190
      TabIndex        =   5
      Text            =   "0"
      Top             =   465
      Width           =   735
   End
   Begin VB.TextBox txtViewMetric 
      Height          =   375
      Index           =   2
      Left            =   8190
      TabIndex        =   4
      Text            =   "0"
      Top             =   885
      Width           =   735
   End
   Begin VB.TextBox txtViewMetric 
      Height          =   375
      Index           =   3
      Left            =   8190
      TabIndex        =   3
      Text            =   "0"
      Top             =   1305
      Width           =   735
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   30
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1980
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      Height          =   255
      Left            =   7470
      TabIndex        =   10
      Top             =   105
      Width           =   675
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      Height          =   255
      Left            =   7470
      TabIndex        =   9
      Top             =   525
      Width           =   675
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Left            =   7470
      TabIndex        =   8
      Top             =   945
      Width           =   675
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      Height          =   255
      Left            =   7470
      TabIndex        =   7
      Top             =   1365
      Width           =   675
   End
   Begin VB.Label lblImageWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   5370
      TabIndex        =   2
      Top             =   30
      Width           =   735
   End
   Begin VB.Label lblImageHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   5370
      TabIndex        =   1
      Top             =   450
      Width           =   735
   End
End
Attribute VB_Name = "verpicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    WindowMoving = True
    ClickMouseX = X
    ClickMouseY = Y

End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If WindowMoving Then
        If (GetAsyncKeyState(vbLeftButton)) Then
            verpicture.Left = verpicture.Left + X - ClickMouseX
            verpicture.Top = verpicture.Top + Y - ClickMouseY
        Else
            WindowMoving = False

        End If

    End If

End Sub
