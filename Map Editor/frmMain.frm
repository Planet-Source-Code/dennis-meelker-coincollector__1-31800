VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Editor"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFill 
      Caption         =   "&Fill"
      Height          =   375
      Left            =   3120
      TabIndex        =   30
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdCoinsEver 
      Caption         =   "&Place coins everywhere"
      Height          =   375
      Left            =   4800
      TabIndex        =   29
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Current Tile"
      Height          =   1335
      Left            =   6240
      TabIndex        =   27
      Top             =   3720
      Width           =   2295
      Begin VB.Label lblCur 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdPlaces 
      Caption         =   "&Remove Wrong Placed Coins"
      Height          =   375
      Left            =   4800
      TabIndex        =   26
      Top             =   7080
      Width           =   1575
   End
   Begin VB.PictureBox picStopperOrg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2280
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picPoliceOrg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1560
      Picture         =   "frmMain.frx":09A2
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   3240
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCountCoins 
      Caption         =   "&Count Coins"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   7080
      Width           =   1575
   End
   Begin VB.PictureBox picMuntOrg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2280
      Picture         =   "frmMain.frx":1344
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picPlayerOrg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1560
      Picture         =   "frmMain.frx":1CE6
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selectie"
      Height          =   1095
      Left            =   6240
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
      Begin VB.PictureBox picSpec 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1680
         Picture         =   "frmMain.frx":2688
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   12
         Tag             =   "1"
         Top             =   480
         Width           =   300
      End
      Begin VB.PictureBox picCur 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   360
         Picture         =   "frmMain.frx":2B7A
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   9
         Tag             =   "0"
         Top             =   480
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speciaal"
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
      Begin VB.PictureBox picStopper 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1200
         Picture         =   "frmMain.frx":306C
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   25
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picPolice 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1560
         Picture         =   "frmMain.frx":355E
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   22
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picDel 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   120
         Picture         =   "frmMain.frx":3A50
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   16
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picMunt 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   840
         Picture         =   "frmMain.frx":3F42
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   14
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picPlayer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   480
         Picture         =   "frmMain.frx":4434
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   10
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles"
      Height          =   1575
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   6
         Left            =   1200
         Picture         =   "frmMain.frx":4926
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   21
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   5
         Left            =   840
         Picture         =   "frmMain.frx":4E18
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   20
         Top             =   1080
         Width           =   300
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   4
         Left            =   840
         Picture         =   "frmMain.frx":530A
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   19
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   3
         Left            =   480
         Picture         =   "frmMain.frx":57FC
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   18
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   2
         Left            =   480
         Picture         =   "frmMain.frx":5CEE
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   17
         Top             =   1080
         Width           =   300
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":61E0
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   6
         Top             =   240
         Width           =   300
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":66D2
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   5
         Top             =   1080
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Openen"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGrass 
      Caption         =   "&New Map"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Opslaan"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   399
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Shape Shape 
         BorderStyle     =   3  'Dot
         Height          =   330
         Left            =   2520
         Top             =   3120
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCoinsEver_Click()
For X = 0 To iX
    For Y = 0 To iY
        If Tile(X, Y).Picture = 0 Or Tile(X, Y).Picture = 3 Or Tile(X, Y).Picture = 4 Or Tile(X, Y).Picture = 6 Then
            BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picTile(Tile(X, Y).Picture), 0, 0, SRCCOPY
            Tile(X, Y).SpecialData1 = 2
            DrawSpecial 2, X, Y
            If Tile(X, Y).SpecialData2 > 9 Then
                DrawSpecial Tile(X, Y).SpecialData2, X, Y
            End If
        End If
    Next Y
Next X
picMap.Refresh
End Sub

Private Sub cmdCountCoins_Click()
Dim iCounter As Integer
For X = 0 To iX
    For Y = 0 To iY
        If Tile(X, Y).SpecialData1 = 2 Or Tile(X, Y).SpecialData1 = 3 Then
            iCounter = iCounter + 1
        End If
    Next Y
Next X
MsgBox iCounter
End Sub

Private Sub cmdFill_Click()
For X = 0 To iX
    For Y = 0 To iY
        Tile(X, Y).Picture = picCur.Tag
        BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picTile(picCur.Tag), 0, 0, SRCCOPY
        DrawSpecial Tile(X, Y).SpecialData1, X, Y
        If Tile(X, Y).SpecialData2 > 9 Then
            DrawSpecial Tile(X, Y).SpecialData2, X, Y
        End If
    Next Y
Next X
picMap.Refresh
End Sub

Private Sub cmdGrass_Click()
ReDim Tile(19, 19) As TileInfo
iX = 19
iY = 19

For X = 0 To iX
    For Y = 0 To iY
        Tile(X, Y).Picture = 0
        Tile(X, Y).SpecialData1 = 0
        BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picTile(0).hDC, 0, 0, SRCCOPY
    Next Y
Next X
picMap.Refresh
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Canceled
Dim FF As Long

cDlg.DialogTitle = "Open map"
cDlg.Filter = "All Files (*.*)|*.*|Maps (*.map)|*.map"
cDlg.FilterIndex = 2
cDlg.ShowOpen

FF = FreeFile

Me.Caption = "Map Editor - " & cDlg.FileName

'Kill "c:\test.map"

Open cDlg.FileName For Binary As #FF
    Line Input #FF, textline
    iX = textline
    Line Input #FF, textline
    iY = textline
    
    ReDim Tile(iX, iY) As TileInfo
    
    For X = 0 To iX
        For Y = 0 To iY
            Get #FF, , Tile(X, Y)
            BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picTile(Tile(X, Y).Picture).hDC, 0, 0, SRCCOPY
            If Tile(X, Y).SpecialData1 > 0 Then
                DrawSpecial Tile(X, Y).SpecialData1, X, Y
            End If
            
            If Tile(X, Y).SpecialData2 > 0 Then
                DrawSpecial Tile(X, Y).SpecialData2, X, Y
            End If
        Next Y
    Next X
Close #FF

picMap.Refresh

Canceled:

End Sub

Private Sub cmdPlaces_Click()
For X = 0 To iX
    For Y = 0 To iY
        If Tile(X, Y).SpecialData1 = 2 Or Tile(X, Y).SpecialData1 = 3 Then
            If Tile(X, Y).Picture = 1 Or Tile(X, Y).Picture = 2 Or Tile(X, Y).Picture = 5 Then
                Shape.Top = Y * 20 - 1
                Shape.Left = X * 20 - 1
                Tile(X, Y).SpecialData1 = 0
                Tile(X, Y).SpecialData2 = 0
                BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picTile(Tile(X, Y).Picture).hDC, 0, 0, SRCCOPY
                picMap.Refresh
            End If
        End If
    Next Y
Next X
MsgBox "Done removing wrong placed coins.", vbInformation
End Sub

Private Sub cmdSave_Click()
On Error GoTo Canceled
Dim FF As Long

cDlg.DialogTitle = "Save map as..."
cDlg.Filter = "All Files (*.*)|*.*|Map (*.map)|*.map"
cDlg.FilterIndex = 2
cDlg.ShowSave

FF = FreeFile

If Len(Dir(cDlg.FileName, vbNormal)) > 0 Then
    If MsgBox("File allready exists, do you want to overwrite it?", vbQuestion + vbYesNo) = vbYes Then
        Kill cDlg.FileName
    Else
        Exit Sub
    End If
End If

Open cDlg.FileName For Binary As #FF
    Put #FF, , iX & vbCrLf & iY & vbCrLf
    
    For X = 0 To iX
        For Y = 0 To iY
            Put #FF, , Tile(X, Y)
        Next Y
    Next X
Close #FF

Canceled:

End Sub



Private Sub Form_Load()
iX = 19
iY = 19
cmdGrass_Click
End Sub

Private Sub picDel_Click()
picSpec.Tag = 0
picSpec.Picture = picDel.Picture
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Int(X / 20)
Y = Int(Y / 20)

If X > 19 Then
    X = 19
End If

If Y > 19 Then
    Y = 19
End If

If X < 0 Then
    X = 0
End If

If Y < 0 Then
    Y = 0
End If

lblCur.Caption = X & "," & Y

Shape.Left = X * 20 - 1
Shape.Top = Y * 20 - 1

If Button = 1 Then
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picCur.hDC, 0, 0, SRCCOPY
    
    If Tile(X, Y).SpecialData1 > 0 Then
        DrawSpecial Tile(X, Y).SpecialData1, X, Y
    End If
    
    Tile(X, Y).Picture = picCur.Tag
    
    picMap.Refresh
ElseIf Button = 2 Then
    'BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picCur.hDC, 0, 0, SRCCOPY
    
    DrawSpecial picSpec.Tag, X, Y
    If picSpec.Tag > 9 Then
        Tile(X, Y).SpecialData2 = picSpec.Tag
    Else
        Tile(X, Y).SpecialData1 = picSpec.Tag
        Tile(X, Y).SpecialData2 = picSpec.Tag
    End If
    
    picMap.Refresh
End If
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Int(X / 20)
Y = Int(Y / 20)

If X > 19 Then
    X = 19
End If

If Y > 19 Then
    Y = 19
End If

If X < 0 Then
    X = 0
End If

If Y < 0 Then
    Y = 0
End If

lblCur.Caption = X & "," & Y

Shape.Left = X * 20 - 1
Shape.Top = Y * 20 - 1

If Button = 1 Then
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picCur.hDC, 0, 0, SRCCOPY
    
    If Tile(X, Y).SpecialData1 > 0 Then
        DrawSpecial Tile(X, Y).SpecialData1, X, Y
    End If
    
    Tile(X, Y).Picture = picCur.Tag
    
    picMap.Refresh
ElseIf Button = 2 Then

    DrawSpecial picSpec.Tag, X, Y
    If picSpec.Tag > 9 Then
        Tile(X, Y).SpecialData2 = picSpec.Tag
    Else
        Tile(X, Y).SpecialData1 = picSpec.Tag
        Tile(X, Y).SpecialData2 = picSpec.Tag
    End If
    
    picMap.Refresh
End If
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Int(X / 20)
Y = Int(Y / 20)

If X > 19 Then
    X = 19
End If

If Y > 19 Then
    Y = 19
End If

If X < 0 Then
    X = 0
End If

If Y < 0 Then
    Y = 0
End If

lblCur.Caption = X & "," & Y

Shape.Left = X * 20 - 1
Shape.Top = Y * 20 - 1

If Button = 2 Then

    DrawSpecial picSpec.Tag, X, Y
    If picSpec.Tag > 9 Then
        Tile(X, Y).SpecialData2 = picSpec.Tag
    Else
        Tile(X, Y).SpecialData1 = picSpec.Tag
        Tile(X, Y).SpecialData2 = picSpec.Tag
    End If
    
    picMap.Refresh
End If
End Sub

Private Sub picMunt_Click()
picSpec.Tag = 2
picSpec.Picture = picMunt.Picture
End Sub

Private Sub picPlayer_Click()
picSpec.Tag = 1
picSpec.Picture = picPlayer.Picture
End Sub

Private Sub picPolice_Click()
picSpec.Tag = 10
picSpec.Picture = picPolice.Picture
End Sub

Private Sub picStopper_Click()
picSpec.Tag = 3
picSpec.Picture = picStopper.Picture
End Sub

Private Sub picTile_Click(Index As Integer)
picCur.Picture = picTile(Index).Picture
picCur.Tag = picTile(Index).Index
End Sub

Public Function DrawSpecial(iIndex As Variant, X As Variant, Y As Variant)
Select Case iIndex
Case 0
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picTile(Tile(X, Y).Picture).hDC, 0, 0, SRCCOPY
Case 1
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picPlayerOrg.hDC, 20, 0, SRCAND
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picPlayerOrg.hDC, 0, 0, SRCPAINT
Case 2
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picMuntOrg.hDC, 20, 0, SRCAND
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picMuntOrg.hDC, 0, 0, SRCPAINT
Case 3
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picStopperOrg.hDC, 20, 0, SRCAND
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picStopperOrg.hDC, 0, 0, SRCPAINT
Case 10
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picPoliceOrg.hDC, 20, 0, SRCAND
    BitBlt picMap.hDC, X * 20, Y * 20, 20, 20, picPoliceOrg.hDC, 0, 0, SRCPAINT
End Select
End Function

