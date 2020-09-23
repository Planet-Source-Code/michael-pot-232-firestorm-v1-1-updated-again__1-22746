Attribute VB_Name = "MdlBitmap"
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As Size) As Long
Public Declare Function SetBitmapDimensionEx Lib "gdi32" (ByVal hbm As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Type Size
        cx As Long
        cy As Long
End Type
Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function GetBPP(Picture As Object) As Long
Dim PI As BITMAP
GetObject Picture.Picture, Len(PI), PI
GetBPP = PI.bmBitsPixel
End Function

Public Function Distance(sx, sy, Ex, Ey) As Long
Distance = Sqr((Ex - sx) ^ 2 + (Ey - sy) ^ 2)
End Function


'32 bit - Add 4 to counter
'24 bit - Add 3 to counter

Sub WriteINI(Path As String, Section As String, Nam As String, Vaule As String)
Dim V As String
V = Vaule
WritePrivateProfileString Section, Nam, V, Path
DoEvents
End Sub
Function ReadINI(Path As String, Section As String, Nam As String) As String
Static R As String * 200
R = ""
GetPrivateProfileString Section, Nam, "Error Reading INI", R, 200, Path
ReadINI = Trim(R)
If Asc(Right(ReadINI, 1)) = 0 Then
ReadINI = Mid(ReadINI, 1, Len(ReadINI) - 1)
End If
End Function

