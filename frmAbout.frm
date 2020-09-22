VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CoinCollector"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4080
      Picture         =   "frmAbout.frx":04FE
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":0DC8
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1692
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1740
      TabIndex        =   2
      Top             =   840
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CoinCollector"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   825
      TabIndex        =   1
      Top             =   120
      Width           =   3285
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Tile the grass picture on the form
For X = 0 To Int(Me.ScaleWidth / 20)
    For Y = 0 To Int(Me.ScaleHeight / 20)
        BitBlt frmAbout.hdc, X * 20, Y * 20, 20, 20, picGrass.hdc, 0, 0, SRCCOPY
    Next Y
Next X
'Refresh the form
frmAbout.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'To save memory
frmAbout.Cls
End Sub
