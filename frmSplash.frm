VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9000
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Interval        =   1000
      Left            =   3960
      Top             =   3240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   6240
      Width           =   3135
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Picture by my sister Laura"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   1800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(FixPath(App.Path) & "graphics\CoinCollector.jpg")
End Sub

Private Sub tmrLoad_Timer()
lblMsg.Caption = "Loading tiles..."
DoEvents
dcTiles = GenerateDC(FixPath(App.Path) & "Graphics\Tiles.bmp")
DoEvents
lblMsg.Caption = "Loading player..."
DoEvents
dcPlayer = GenerateDC(FixPath(App.Path) & "Graphics\Player.bmp")
DoEvents
lblMsg.Caption = "Loading coins..."
DoEvents
dcCoin = GenerateDC(FixPath(App.Path) & "Graphics\Coin.bmp")
dcStopper = GenerateDC(FixPath(App.Path) & "Graphics\Stopper.bmp")
DoEvents
lblMsg.Caption = "Loading messages..."
DoEvents
dcCompleted = GenerateDC(FixPath(App.Path) & "Graphics\Completed.bmp")
dcGameOver = GenerateDC(FixPath(App.Path) & "Graphics\Game Over.bmp")
dcCaught = GenerateDC(FixPath(App.Path) & "Graphics\Caught.bmp")
DoEvents
lblMsg.Caption = "Loading police..."
DoEvents
dcPolice = GenerateDC(FixPath(App.Path) & "Graphics\Police.bmp")
DoEvents
lblMsg.Caption = "Loading temp pictures..."
DoEvents
dcLast = GenerateDC(FixPath(App.Path) & "Graphics\Coin.bmp")
DoEvents
Unload Me
frmMain.Show
End Sub
