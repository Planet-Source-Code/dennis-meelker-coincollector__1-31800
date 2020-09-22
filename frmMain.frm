VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CoinCollector"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   120
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&Start a new game"
      End
      Begin VB.Menu mnuGameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameCustom 
         Caption         =   "&Play a custom level"
      End
      Begin VB.Menu mnuGameSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHowTo 
         Caption         =   "&How to play"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'End the game
TimeToEnd = True
'Delete the dc's
DeleteDCs
End
End Sub

Private Sub mnuGameCustom_Click()
'Show the play a custom level form
frmCustom.Show vbModal
End Sub

Private Sub mnuGameNew_Click()
'Start level 1 with 3 lives
iLives = 3
TimeToEnd = True
'Set the setting for the new game
frmMain.Tag = FixPath(App.Path) & "Maps\Level1.map"
tmrStart.Tag = False
tmrStart = True
End Sub

Private Sub mnuGameQuit_Click()
'Quit the game
TimeToEnd = True
'Delete the dc's
DeleteDCs
End
End Sub

Private Sub mnuHelpAbout_Click()
'Show the about form
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpHowTo_Click()
'Show the how to play form
frmHowToPlay.Show vbModal
End Sub

Private Sub tmrStart_Timer()
'Start the game, i used this timer because if i didn't, there would be a lot of mainloop's
'wich would result in a lot of memory usage
tmrStart.Enabled = False
MainLoop frmMain.Tag, tmrStart.Tag
End Sub
