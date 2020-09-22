VERSION 5.00
Begin VB.Form frmHowToPlay 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to play"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmHowToPlay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is you."
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   795
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   120
      Picture         =   "frmHowToPlay.frx":000C
      Top             =   720
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   1185
      Left            =   120
      Picture         =   "frmHowToPlay.frx":04FE
      Top             =   3480
      Width           =   1725
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use the arrow keys to steer."
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Policemen."
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   780
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   120
      Picture         =   "frmHowToPlay.frx":70A4
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "While collecting the coins you will have to watch out for the policemen, since you are a thief they wont mind arresting you."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHowToPlay.frx":7596
      Height          =   555
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   4620
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   120
      Picture         =   "frmHowToPlay.frx":7638
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Regulair coin."
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   120
      Picture         =   "frmHowToPlay.frx":7B2A
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHowToPlay.frx":801C
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmHowToPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Tile the grass picture on the form
For X = 0 To Int(Me.ScaleWidth / 20)
    For Y = 0 To Int(Me.ScaleHeight / 20)
        BitBlt Me.hdc, X * 20, Y * 20, 20, 20, frmAbout.picGrass.hdc, 0, 0, SRCCOPY
    Next Y
Next X
'Refresh the form
Me.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Cls
'To save memory
End Sub

