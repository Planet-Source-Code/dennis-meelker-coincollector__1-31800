Attribute VB_Name = "modMapEditor"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Tile() As TileInfo

Public iX As Integer
Public iY As Integer

'public

Public Type TileInfo
    X As Variant
    Y As Variant
    Picture As Variant
    Walkable As Boolean
    SpecialData1 As Variant
    SpecialData2 As Variant
End Type
