VERSION 5.00
Begin VB.Form frmCustom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Play a custom level"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2370
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   158
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrEnable 
      Interval        =   1
      Left            =   120
      Top             =   3240
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   3840
      Pattern         =   "*.map"
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstMaps 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlay_Click()
'Hide the form so the listbox stay's filled
Me.Hide
'Stop the current game
TimeToEnd = True
'Set the lives
iLives = -1
'Set the setting for the new game
frmMain.Tag = FixPath(File1.Path) & lstMaps.List(lstMaps.ListIndex) & ".map"
frmMain.tmrStart.Tag = True
frmMain.tmrStart = True
End Sub

Private Sub Form_Load()
'Tile the grass picture on the form
For X = 0 To Int(Me.ScaleWidth / 20)
    For Y = 0 To Int(Me.ScaleHeight / 20)
        BitBlt Me.hdc, X * 20, Y * 20, 20, 20, frmAbout.picGrass.hdc, 0, 0, SRCCOPY
    Next Y
Next X
'Refresh the form
Me.Refresh
'Set the path
File1.Path = FixPath(App.Path) & "Custom Maps"
'Clear the list
lstMaps.Clear
'Check if there are any files
If File1.ListCount = 0 Then Exit Sub
'Add all the files to the listbox
For i = 0 To File1.ListCount - 1
    lstMaps.AddItem Left(File1.List(i), Len(File1.List(i)) - 4)
Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Cls
'To save memory
End Sub

Private Sub tmrEnable_Timer()
'Check if there is anything selected
If lstMaps.ListIndex > -1 Then
    'Check if the selected one is empty
    If lstMaps.List(lstMaps.ListIndex) = "" Then
        'Disable the button
        cmdPlay.Enabled = False
    Else
        'Enable the button
        cmdPlay.Enabled = True
    End If
End If
End Sub
