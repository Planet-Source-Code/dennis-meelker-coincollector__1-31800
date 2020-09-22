Attribute VB_Name = "modPacMan"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_DOWN As Integer = &H1000

'Dc's
Public dcTiles As Long
Public dcPlayer As Long
Public dcCoin As Long
Public dcCompleted As Long
Public dcLast As Long
Public dcPolice As Long
Public dcStopper As Long
Public dcCaught As Long
Public dcGameOver As Long

'-----------------------------
Public TimeToEnd As Boolean
Public NextLevel As Boolean
Public Restart As Boolean

'Booleans for the buttons
Public UpPressed As Boolean
Public DownPressed As Boolean
Public LeftPressed As Boolean
Public RightPressed As Boolean

Public UpKey As Long
Public DownKey As Long
Public LeftKey As Long
Public RightKey As Long
'-----------------------------
'Police booleans
Public bPolice As Boolean

Public iLastDirection As Integer

Public PoliceCounter As Integer

Public iLives As Integer

Public iCoins As Integer
Public iCollectedCoins As Integer

Public Xtmp As Integer
Public Ytmp As Integer

Public iX As Integer
Public iY As Integer

Public PoliceStoppedFor As Integer

Public Tile() As TileInfo

Public Player As CharacterType
Public Police() As CharacterType

Public Type TileInfo
    X As Variant
    Y As Variant
    Picture As Variant
    Walkable As Boolean
    SpecialData1 As Variant
    SpecialData2 As Variant
End Type

Public Type CharacterType
    X As Integer
    Y As Integer
    TmpX As Variant
    TmpY As Variant
    Last As Long
    Direction As Integer
End Type

Dim TmpX As Integer
Dim TmpY As Integer

Public Function GetUserInput()
'Get all the user input
UpPressed = GetKeyState(UpKey) And KEY_DOWN
DownPressed = GetKeyState(DownKey) And KEY_DOWN
LeftPressed = GetKeyState(LeftKey) And KEY_DOWN
RightPressed = GetKeyState(RightKey) And KEY_DOWN
End Function
Public Sub InitButtons()
'Load the keys
UpKey = vbKeyUp
DownKey = vbKeyDown
LeftKey = vbKeyLeft
RightKey = vbKeyRight
End Sub


Public Function MainLoop(sFilename As String, bCustom As Boolean)
Const TickDifference As Long = 10
Dim LastTick As Long
Dim CurrentTick As Long
Dim CurrentTile As Integer
'Make sure the form is loaded
'frmMain.Show
'Get the last tick
LastTick = GetTickCount
'Set some variables
TimeToEnd = False
NextLevel = False
iCollectedCoins = 0
'Police things
PoliceStoppedFor = 0
Restart = False
bPolice = False
PoliceCounter = 0
'Redim the police array so it is empty
ReDim Police(0) As CharacterType
'Set the default start position
Player.X = 0
Player.Y = 0
TmpX = 0
TmpY = 0
'Set the default direction
Player.Direction = 0
'Load the buttons
InitButtons
'Load the given map
LoadMap sFilename
'Check if this is a custom game
If iLives = -1 Then
    'It is a custom game, don't display any lives
    frmMain.Caption = "CoinCollector"
Else
    'It ain't a custom game, display the lives
    frmMain.Caption = "CoinCollector - " & iLives & " lives left"
End If
'Draw the temp picture of the player
BitBlt dcLast, 0, 0, 20, 20, frmMain.hdc, Player.X, Player.Y, SRCCOPY

Do
    'Check if the loop should stop
    If Not TimeToEnd Then
        'Get the last tick
        CurrentTick = GetTickCount()
        'Check if there are 10 miliseconds ellapsed
        If GetTickCount() - LastTick > TickDifference Then
            'There are, get the last tick and store it
            LastTick = GetTickCount()
            'Get the input
            GetUserInput
            'Set the player's direction
            SetDirections
            'Check if the player collides with a wall
            CheckPlayerCollition
            'Check if the player is still in the level
            SeeIfPlayerIn
            'Check if the player is on a coin
            SeeIfOnCoin
            'Blit the temp picture on the form
            BitBlt frmMain.hdc, TmpX, TmpY, 20, 20, dcLast, 0, 0, SRCCOPY
            'Check if there is any police
            If bPolice = True Then
                'Check if they should move allready
                If PoliceStoppedFor = 0 Then
                    'They may move, move them
                    MovePolice
                    'Draw them
                    DrawPolice
                    'Save the last position
                    FixTmps
                Else
                    'Decrease the counter with one
                    PoliceStoppedFor = PoliceStoppedFor - 1
                End If
            End If
            'Draw on the temp picture again
            BitBlt dcLast, 0, 0, 20, 20, frmMain.hdc, Player.X, Player.Y, SRCCOPY
            
            'Check if there is any police
            If bPolice = True Then
                'See if the police can arrest you
                If PoliceStoppedFor = 0 Then
                    'They can, loop trough all the policemen
                    For i = 1 To PoliceCounter
                        'Check if you collide with the current policeman
                        If IsIn(Police(i).X, Police(i).Y, Player.X + 10, Player.Y + 10) = True Then
                            'You do, redraw the player
                            DrawStuff
                            'Set the restart variable to true
                            Restart = True
                            'Check how much lives you have
                            If iLives = 0 Then
                                '0 lives left, display the game over picture
                                BitBlt frmMain.hdc, 71, 180.5, 258, 39, dcGameOver, 0, 39, SRCAND
                                BitBlt frmMain.hdc, 71, 180.5, 258, 39, dcGameOver, 0, 0, SRCPAINT
                                'The level doesn't need to be restarted
                                Restart = False
                            ElseIf iLives = -1 Then
                                'This is a custom game, show the Caught picture
                                'The level should be restarted
                                Restart = True
                                BitBlt frmMain.hdc, 37, 176, 326, 48, dcCaught, 0, 48, SRCAND
                                BitBlt frmMain.hdc, 37, 176, 326, 48, dcCaught, 0, 0, SRCPAINT
                            Else
                                'You still have lives, display the Caught picture
                                BitBlt frmMain.hdc, 37, 176, 326, 48, dcCaught, 0, 48, SRCAND
                                BitBlt frmMain.hdc, 37, 176, 326, 48, dcCaught, 0, 0, SRCPAINT
                            End If
                            'Refresh the form so changes will appear
                            frmMain.Refresh
                            'Loop until the user presses space
                            Do Until (GetKeyState(vbKeySpace) And KEY_DOWN)
                                'Let the computer do it's things
                                DoEvents
                            Loop
                            'If you have more than zero lives decrease with one
                            If iLives > 0 Then
                                iLives = iLives - 1
                                'Display the changes
                                frmMain.Caption = "CoinCollector - " & iLives & " lives left"
                            End If
                            'The mainloop should stop
                            TimeToEnd = True
                            
                            Exit Do
                        End If
                    Next i
                End If
            End If
            'Draw the player
            DrawStuff
            'Update the last position variable
            TmpX = Player.X
            TmpY = Player.Y
            'Refresh the form so changes will appear
            frmMain.Refresh
        End If
    
    End If
    
    'Let the computer do it's things
    DoEvents
    
Loop Until TimeToEnd
'Check if the next level should be loaded
If NextLevel = True Then
    'Check if it's a custom map
    If bCustom = False Then
        'It isn't, Start the next level
        frmMain.Tag = NextMap(sFilename)
        frmMain.tmrStart.Tag = bCustom
        frmMain.tmrStart = True
        Exit Function
    End If
End If
'Check if the level should be restarted
If Restart = True Then
    'It should, do so
    frmMain.Tag = sFilename
    frmMain.tmrStart.Tag = bCustom
    frmMain.tmrStart = True
End If
End Function

Public Sub DrawStuff()
'Check wich direction the player is at and draw him
Select Case Player.Direction
Case 0
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 20, 0, SRCAND
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 0, 0, SRCPAINT
Case 1
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 20, 0, SRCAND
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 0, 0, SRCPAINT
Case 2
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 60, 0, SRCAND
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 40, 0, SRCPAINT
Case 3
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 20, 20, SRCAND
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 0, 20, SRCPAINT
Case 4
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 60, 20, SRCAND
    BitBlt frmMain.hdc, Player.X, Player.Y, 20, 20, dcPlayer, 40, 20, SRCPAINT
End Select
End Sub

Public Sub LoadMap(sMap As String)
Dim FF As Long
Dim tmpTile As TileInfo
'Get a free file number
FF = FreeFile
'Open the file
Open sMap For Binary As #FF
    'Read the x max
    Line Input #FF, textline
    iX = textline
    'Read the y max
    Line Input #FF, textline
    iY = textline
    'Redim the tile array
    ReDim Tile(iX, iY) As TileInfo
    'Loop trough all the tiles and load them from the file
    For X = 0 To iX
        For Y = 0 To iY
            Get #FF, , tmpTile
            Tile(X, Y) = tmpTile
        Next Y
    Next X
Close #FF
'Loop trough all the files
For X = 0 To iX
    For Y = 0 To iY
        'Draw the tile
        DrawTile X, Y
        'See if is the start position
        If Tile(X, Y).SpecialData1 = 1 Then
            'It is, set the variables
            Player.X = X * 20
            Player.Y = Y * 20
            TmpX = X * 20
            TmpY = Y * 20
        'See if it is a coin
        ElseIf Tile(X, Y).SpecialData1 = 2 Then
            'It is, draw it
            DrawSpecial 2, X, Y
        'See if it is a stopper coin
        ElseIf Tile(X, Y).SpecialData1 = 3 Then
            'It is, draw it
            DrawSpecial 3, X, Y
        End If
        'See if there is a policeman on the tile
        If Tile(X, Y).SpecialData2 = 10 Then
            'There is, increase the counter by one
            PoliceCounter = PoliceCounter + 1
            'Redim the array
            ReDim Preserve Police(PoliceCounter) As CharacterType
            'Set the policeman's position and direction
            Police(PoliceCounter).X = X * 20
            Police(PoliceCounter).Y = Y * 20
            Police(PoliceCounter).TmpX = X * 20
            Police(PoliceCounter).TmpY = Y * 20
            Police(PoliceCounter).Direction = 4
            'Create a dc for the policeman
            Police(PoliceCounter).Last = GenerateDC(FixPath(App.Path) & "graphics\Coin.bmp")
            'Draw on it
            BitBlt Police(PoliceCounter).Last, 0, 0, 20, 20, frmMain.hdc, X * 20, Y * 20, SRCCOPY
            'Set the boolean to show that there is police
            bPolice = True
        End If
    Next Y
Next X
'Count the coins in the map
CountCoins
End Sub

Public Function CheckForCollition(bDown As Boolean, bRight As Boolean) As Boolean
'See if the collition should be checked on the right side
If bRight = True Then
    'It is, get the tile's x number
    Xtmp = Int((Player.X + 19) / 20)
Else
    'it aint
    Xtmp = Int(Player.X / 20)
End If
'See if the collition should be checked downwards
If bDown = True Then
    'It should, get the tile's y number
    Ytmp = Int((Player.Y + 19) / 20)
Else
    'it aint
    Ytmp = Int(Player.Y / 20)
End If
'Set the function to false
CheckForCollition = False
'See if the tile is inside the form
If Xtmp > iX Or Ytmp > iY Or Xtmp < 0 Or Ytmp < 0 Then
    'it aint, exit from the function
    CheckForCollition = True
    Exit Function
End If
'See if the tile is a wall
If Tile(Xtmp, Ytmp).Picture = 1 Or Tile(Xtmp, Ytmp).Picture = 2 Or Tile(Xtmp, Ytmp).Picture = 5 Then
    'It is return true
    CheckForCollition = True
End If
End Function
Public Function DrawSpecial(iIndex As Variant, X As Variant, Y As Variant)
'See what should be drawn
Select Case iIndex
Case 1
    'The start position
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcPlayer, 20, 0, SRCAND
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcPlayer, 0, 0, SRCPAINT
Case 2
    'A coin
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcCoin, 20, 0, SRCAND
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcCoin, 0, 0, SRCPAINT
Case 3
    'A stopper coin
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcStopper, 20, 0, SRCAND
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcStopper, 0, 0, SRCPAINT
End Select
End Function
Public Sub CheckPlayerCollition()
'See which direction the player is facing
Select Case Player.Direction
    Case 1
        'Up, check for collition
        Player.Y = Player.Y - 2
                    
        If CheckForCollition(False, False) = True Then
            Player.Y = Player.Y + 2
        End If
                    
        If CheckForCollition(False, True) = True Then
            Player.Y = Player.Y + 2
        End If
    Case 2
        'Right, check for collition
        Player.X = Player.X + 2
                    
        If CheckForCollition(False, True) = True Then
            Player.X = Player.X - 2
        End If
                    
        If CheckForCollition(True, True) = True Then
            Player.X = Player.X - 2
        End If
    Case 3
        'Down, check for collition
        Player.Y = Player.Y + 2
                    
        If CheckForCollition(True, False) = True Then
            Player.Y = Player.Y - 2
        End If
                    
        If CheckForCollition(True, True) = True Then
            Player.Y = Player.Y - 2
        End If
    Case 4
        'Left, check for collition
        Player.X = Player.X - 2
                    
        If CheckForCollition(False, False) = True Then
            Player.X = Player.X + 2
        End If
                    
            If CheckForCollition(True, False) = True Then
            Player.X = Player.X + 2
        End If
End Select
End Sub

Public Sub SetDirections()
'See if up is pressed
If UpPressed Then
    'It is, decrease the y variable
    Player.Y = Player.Y - 2
    'Check if there is collition
    If CheckForCollition(False, False) = True Then
        Player.Y = Player.Y + 2
    Else
        If CheckForCollition(False, True) = True Then
            Player.Y = Player.Y + 2
        Else
            'There aint collition, set the direction
            Player.Direction = 1
            Player.Y = Player.Y + 2
        End If
    End If
End If
'See if down is pressed
If DownPressed Then
    'It is, increase the y variable
    Player.Y = Player.Y + 2
    'Check if there is collition
    If CheckForCollition(True, False) = True Then
        Player.Y = Player.Y - 2
    Else
        If CheckForCollition(True, True) = True Then
            Player.Y = Player.Y - 2
        Else
            'There aint collition, set the direction
            Player.Y = Player.Y - 2
            Player.Direction = 3
        End If
    End If
                
End If
'See if Left is pressed
If LeftPressed Then
    'It is, decrease the x variable
    Player.X = Player.X - 2
    'Check if there is collition
    If CheckForCollition(False, False) = True Then
        Player.X = Player.X + 2
    Else
        If CheckForCollition(True, False) = True Then
            Player.X = Player.X + 2
        Else
            'There aint collition, set the direction
            Player.Direction = 4
            Player.X = Player.X + 2
        End If
    End If
End If
'See if Right is pressed
If RightPressed Then
    'It is, increase the x variable
    Player.X = Player.X + 2
    'Check if there is collition
    If CheckForCollition(False, True) = True Then
        Player.X = Player.X - 2
    Else
        If CheckForCollition(True, True) = True Then
            Player.X = Player.X - 2
        Else
            'There aint collition, set the direction
            Player.Direction = 2
            Player.X = Player.X - 2
        End If
    End If
End If
End Sub

Public Sub SeeIfPlayerIn()
'Check if the player is still inside the field
If Player.X < 0 Then
    Player.X = 0
    Player.Direction = 0
End If
If Player.Y < 0 Then
    Player.Y = 0
    Player.Direction = 0
End If
If Player.X > iX * 20 Then
    Player.X = iX * 20
    Player.Direction = 0
End If
If Player.Y > iY * 20 Then
    Player.Y = iY * 20
    Player.Direction = 0
End If
End Sub

Public Sub DrawTile(X As Variant, Y As Variant)
'See what tile it is and draw it
Select Case Tile(X, Y).Picture
Case 0
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 0, 0, SRCCOPY
Case 1
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 20, 0, SRCCOPY
Case 2
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 40, 0, SRCCOPY
Case 3
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 60, 0, SRCCOPY
Case 4
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 80, 0, SRCCOPY
Case 5
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 100, 0, SRCCOPY
Case 6
    BitBlt frmMain.hdc, X * 20, Y * 20, 20, 20, dcTiles, 120, 0, SRCCOPY
End Select
End Sub

Public Sub CountCoins()
'Set the counter to zero
iCoins = 0
'Loop trough all the tiles
For X = 0 To iX
    For Y = 0 To iY
        'Check if it is a coin
        If Tile(X, Y).SpecialData1 = 2 Or Tile(X, Y).SpecialData1 = 3 Then
            'Update the counter
            iCoins = iCoins + 1
        End If
    Next Y
Next X
End Sub

Public Sub InitPictures()
'Load all the used pictures into the memory
dcTiles = GenerateDC(FixPath(App.Path) & "Tiles.bmp")
dcPlayer = GenerateDC(FixPath(App.Path) & "Player.bmp")
dcCoin = GenerateDC(FixPath(App.Path) & "Coin.bmp")
dcCompleted = GenerateDC(FixPath(App.Path) & "Completed.bmp")
dcGameOver = GenerateDC(FixPath(App.Path) & "Game Over.bmp")
dcCaught = GenerateDC(FixPath(App.Path) & "Caught.bmp")
dcLast = GenerateDC(FixPath(App.Path) & "Coin.bmp")
dcPolice = GenerateDC(FixPath(App.Path) & "Police.bmp")
dcStopper = GenerateDC(FixPath(App.Path) & "Stopper.bmp")
End Sub

Public Function FixPath(sPath As Variant) As String
'Check if the first character on the right is a \
If Right(sPath, 1) = "\" Then
    'It is, ok
    FixPath = sPath
Else
    'It aint, put a \ at the end
    FixPath = sPath & "\"
End If
End Function

Public Sub DeleteDCs()
'Delete all the pictures from the memory
DeleteGeneratedDC dcPlayer
DeleteGeneratedDC dcTiles
DeleteGeneratedDC dcCoin
DeleteGeneratedDC dcCompleted
DeleteGeneratedDC dcLast
DeleteGeneratedDC dcPolice
DeleteGeneratedDC dcStopper
DeleteGeneratedDC dcCaught
DeleteGeneratedDC dcGameOver
'Delete all the polices temp pictures from the memory
For i = 1 To PoliceCounter
    DeleteGeneratedDC Police(i).Last
Next i
DoEvents
End Sub

Public Sub SeeIfOnCoin()
'Get the tile
Xtmp = Int((Player.X + 10) / 20)
Ytmp = Int((Player.Y + 10) / 20)
'Check if it is a coin
If Tile(Xtmp, Ytmp).SpecialData1 = 2 Or Tile(Xtmp, Ytmp).SpecialData1 = 3 Then
    'It is, increase the counter
    iCollectedCoins = iCollectedCoins + 1
    'Check if it is a stopper coin
    If Tile(Xtmp, Ytmp).SpecialData1 = 3 Then
        'It is, make the police stop
        PoliceStoppedFor = 500
        'Play the sound
        PlaySound "Pickup Stopper.wav"
    Else
        'It ain't, Play the sound
        PlaySound "Pickup.wav"
    End If
    'Set the specialdata variable to zero
    Tile(Xtmp, Ytmp).SpecialData1 = 0
    'Draw the temp picture on the form
    BitBlt frmMain.hdc, TmpX, TmpY, 20, 20, dcLast, 0, 0, SRCCOPY
    'Loop trough all the policemen, if there are any
    If bPolice = True Then
        For i = 1 To PoliceCounter
            'Draw their temp pictures
            BitBlt frmMain.hdc, Police(i).TmpX, Police(i).TmpY, 20, 20, Police(i).Last, 0, 0, SRCCOPY
        Next i
    End If
    'Draw the tile
    DrawTile Xtmp, Ytmp
    'Check if there are any policemen
    If bPolice = True Then
        'Loop trough them
        For i = 1 To PoliceCounter
            'Draw on their temp picture
            BitBlt Police(i).Last, 0, 0, 20, 20, frmMain.hdc, Police(i).X, Police(i).Y, SRCCOPY
        Next i
        'Redraw them
        DrawPolice
    End If
    'Draw the temp picture
    BitBlt dcLast, 0, 0, 20, 20, frmMain.hdc, TmpX, TmpY, SRCCOPY
    'Check if all the coins are picked up
    If iCollectedCoins = iCoins Then
        'They are, show the Level Completed picture
        BitBlt frmMain.hdc, 10.5, 176, 379, 48, dcCompleted, 0, 48, SRCAND
        BitBlt frmMain.hdc, 10.5, 176, 379, 48, dcCompleted, 0, 0, SRCPAINT
        'Refresh the form to show the changes
        frmMain.Refresh
        'Loop until space is pressed
        Do Until (GetKeyState(vbKeySpace) And KEY_DOWN)
            'let the computer do what het has to do
            DoEvents
        Loop
        'See if the lives should be displayed
        If iLives > -1 Then
            'They should, increase the counter with one because the level is completed
            iLives = iLives + 1
            'Show the new value
            frmMain.Caption = "CoinCollector - " & iLives & " lives left"
        End If
        'Set the booleans
        TimeToEnd = True
        NextLevel = True
    End If
End If
End Sub

Public Function NextMap(sFile As Variant) As String
Dim sPath As String
Dim sFileOnly As String
'Get the filetitle
sFileOnly = ExtractFileTitle(sFile)
'Get the path
sPath = FixPath(ExtractPath(sFile))
'Get the number of the current map
sFileOnly = CInt(Right(sFileOnly, Len(sFileOnly) - 5))
'Return the filename of the next map
NextMap = sPath & "level" & (sFileOnly + 1) & ".map"
End Function
Public Function ExtractFileTitle(File As Variant) As String
Dim pos As Integer
Dim ef As String
'Get rid of the .map part
File = Left(File, Len(File) - 4)
'Get rid of the path
For i = 0 To Len(File)
    ef = Right(File, i)
    If Left(ef, 1) = "/" Or Left(ef, 1) = "\" Then
        ExtractFileTitle = Right(ef, (Len(ef) - 1))
        Exit Function
    End If
Next i
End Function
Public Function ExtractPath(File As Variant) As String
Dim pos As Integer
Dim ef As String
'Get the path
For i = 0 To Len(File)
    ef = Right(File, i)
    If Left(ef, 1) = "/" Or Left(ef, 1) = "\" Then
        ExtractPath = Left(File, Len(File) - Len(ef))
        Exit Function
    End If
Next i
End Function
Public Sub DrawPolice()
'Loop trough the policemen
For i = 1 To PoliceCounter
    'Draw there temp picture on the form
    BitBlt frmMain.hdc, Police(i).TmpX, Police(i).TmpY, 20, 20, Police(i).Last, 0, 0, SRCCOPY
Next i

For i = 1 To PoliceCounter
    'Draw there temp picture
    BitBlt Police(i).Last, 0, 0, 20, 20, frmMain.hdc, Police(i).X, Police(i).Y, SRCCOPY
Next i
'Check what direction the policemen is facing and draw him
For i = 1 To PoliceCounter
    Select Case Police(i).Direction
    Case 1
        'Up
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 20, 0, SRCAND
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 0, 0, SRCPAINT
    Case 2
        'Right
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 60, 0, SRCAND
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 40, 0, SRCPAINT
    Case 3
        'Down
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 20, 20, SRCAND
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 0, 20, SRCPAINT
    Case 4
        'Left
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 60, 20, SRCAND
        BitBlt frmMain.hdc, Police(i).X, Police(i).Y, 20, 20, dcPolice, 40, 20, SRCPAINT
    End Select
Next i
End Sub

Public Sub MovePolice()
'Check if the policemen should move
If PoliceStoppedFor > 0 Then Exit Sub
'They should, loop trough all of them
For i = 1 To PoliceCounter
    'See what's there direction and de/in crease the variable
    Select Case Police(i).Direction
    Case 1
        'Up
        Police(i).Y = Police(i).Y - 1
    Case 2
        'Right
        Police(i).X = Police(i).X + 1
    Case 3
        'Down
        Police(i).Y = Police(i).Y + 1
    Case 4
        'Left
        Police(i).X = Police(i).X - 1
    End Select
    'See if he is still inside the form
    If Police(i).X < 0 Then
        Police(i).X = 0
        PoliceRandomDirection i
    End If
    
    If Police(i).Y < 0 Then
        Police(i).Y = 0
        PoliceRandomDirection i
    End If
    
    If Police(i).X > iX * 20 Then
        Police(i).X = iX * 20
        PoliceRandomDirection i
    End If
    
    If Police(i).Y > iY * 20 Then
        Police(i).Y = iY * 20
        PoliceRandomDirection i
    End If
    'Check if there is collition
    CheckForPoliceCollition i
Next i
End Sub

Public Sub FixTmps()
'Check if the policemen should move
If PoliceStoppedFor > 0 Then Exit Sub
'They should, check if there are any
If bPolice = True Then
    'Loop trough all of them and update the variables
    For i = 1 To PoliceCounter
        Police(i).TmpX = Police(i).X
        Police(i).TmpY = Police(i).Y
    Next i
End If
End Sub

Public Function CheckForPoliceCollition(iIndex As Variant)
Dim X As Integer
Dim Y As Integer
Dim i As Integer
'Set i as the number
i = iIndex
'Check the policeman's direction
Select Case Police(i).Direction
Case 1
    'Get the tile
    X = Int(Police(i).X / 20)
    Y = Int(Police(i).Y / 20)
    'See if it is a wall
    If Tile(X, Y).Picture = 1 Or Tile(X, Y).Picture = 2 Or Tile(X, Y).Picture = 5 Then
        'It is move the policeman back
        Police(i).Y = Police(i).Y + 1
        'Get a random direction
        PoliceRandomDirection i
    End If
Case 2
    'Get the tile
    X = Int((Police(i).X + 19) / 20)
    Y = Int(Police(i).Y / 20)
    'See if it is a wall
    If Tile(X, Y).Picture = 1 Or Tile(X, Y).Picture = 2 Or Tile(X, Y).Picture = 5 Then
        'It is move the policeman back
        Police(i).X = Police(i).X - 1
        'Get a random direction
        PoliceRandomDirection i
    End If
Case 3
    'Get the tile
    X = Int(Police(i).X / 20)
    Y = Int((Police(i).Y + 19) / 20)
    'See if it is a wall
    If Tile(X, Y).Picture = 1 Or Tile(X, Y).Picture = 2 Or Tile(X, Y).Picture = 5 Then
        'It is move the policeman back
        Police(i).Y = Police(i).Y - 1
        'Get a random direction
        PoliceRandomDirection i
    End If
Case 4
    'Get the tile
    X = Int(Police(i).X / 20)
    Y = Int(Police(i).Y / 20)
    'See if it is a wall
    If Tile(X, Y).Picture = 1 Or Tile(X, Y).Picture = 2 Or Tile(X, Y).Picture = 5 Then
        'It is move the policeman back
        Police(i).X = Police(i).X + 1
        'Get a random direction
        PoliceRandomDirection i
    End If
End Select
End Function

Public Function PoliceRandomDirection(i As Variant)
Dim iDir As Integer
Dim X As Integer
Dim Y As Integer
'Turn on randomize
Randomize
'Start a loop
Do
    'Get a direction by random
    iDir = Int((4 * Rnd) + 1)
    'Check wich direction it is
    Select Case iDir
    Case 1
        'Up, check if the policeman is still on the form
        If Not Police(i).Y = 0 Then
        'Get the tile
        X = Int(Police(i).X / 20)
        Y = Int((Police(i).Y - 19) / 20)
        'Make sure there aint a wall
        If Tile(X, Y).Picture <> 1 And Tile(X, Y).Picture <> 2 And Tile(X, Y).Picture <> 5 Then
            'Set the new direction
            Police(i).Direction = 1
            Exit Do
        End If
        End If
    Case 2
        'Right, check if the policeman is still on the form
        If Not Police(i).X = iX * 20 Then
            'Get the tile
            X = Int((Police(i).X + 39) / 20)
            Y = Int(Police(i).Y / 20)
            'Make sure there aint a wall
            If Tile(X, Y).Picture <> 1 And Tile(X, Y).Picture <> 2 And Tile(X, Y).Picture <> 5 Then
                'Set the new direction
                Police(i).Direction = 2
                Exit Do
            End If
        End If
    Case 3
        'Down, check if the policeman is still on the form
        If Not Police(i).Y = iY * 20 Then
            'Get the tile
            X = Int(Police(i).X / 20)
            Y = Int((Police(i).Y + 39) / 20)
            'Make sure there aint a wall
            If Tile(X, Y).Picture <> 1 And Tile(X, Y).Picture <> 2 And Tile(X, Y).Picture <> 5 Then
                'Set the new direction
                Police(i).Direction = 3
                Exit Do
            End If
        End If
    Case 4
        'Left, check if the policeman is still on the form
        If Not Police(i).X = 0 Then
            'Get the tile
            X = Int((Police(i).X - 19) / 20)
            Y = Int(Police(i).Y / 20)
            'Make sure there aint a wall
            If Tile(X, Y).Picture <> 1 And Tile(X, Y).Picture <> 2 And Tile(X, Y).Picture <> 5 Then
                'Set the new direction
                Police(i).Direction = 4
                Exit Do
            End If
        End If
    End Select
Loop
End Function

Public Function IsIn(X As Integer, Y As Integer, XPos As Integer, YPos As Integer) As Boolean
'Check if the given position is inside the other one
If XPos > X And XPos < (X + 20) And YPos > Y And YPos < (Y + 20) Then
    'It is, return true
    IsIn = True
Else
    'It aint, return false
    IsIn = False
End If
End Function

Public Function PlaySound(sFile As Variant)
'Play the given sound
sndPlaySound FixPath(App.Path) & "sound\" & sFile, SND_NODEFAULT + SND_ASYNC
End Function
