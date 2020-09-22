Attribute VB_Name = "modColors"
Option Explicit

Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffbits As Long
End Type
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImport As Long
End Type
Dim XRes As Long
Dim YRes As Long
Global PicFileName As String
Global PaletteSize As Long
Global h1 As BITMAPFILEHEADER
Global h2 As BITMAPINFOHEADER
Global PicPalette(256, 2) As Integer
Global pic() As Long
Public XOld, YOld, XStart, YStart As Single
Public Const Key_LButton = 1
Public Const Key_RButton = 2
Public ThisX, ThisY As Single
Public AvgColor(2) As Long
Public OrigColor(2) As Long
Public StainColor(2) As Long
Public OKtoCheck As Boolean

Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function GetPixelFormat Lib "gdi32" (ByVal hDC As Long) As Long


Function AddZero(Num As Integer) As String
Dim i As Integer
Dim temp As String
For i = 1 To Num
    temp = temp & "0"
Next i
AddZero = temp
End Function

Function GetRed(MyColor As String) As Integer
Dim TheRed As String
MyColor = AddZero(6 - Len(MyColor)) & MyColor
TheRed = "&H" & Right$(MyColor, 2)
GetRed = CDec(TheRed)
End Function

Function GetBlue(MyColor As String) As Integer
Dim TheBlue As String
MyColor = AddZero(6 - Len(MyColor)) & MyColor
TheBlue = "&H" & Left$(Right$(MyColor, 6), 2)
GetBlue = CDec(TheBlue)
End Function

Function GetGreen(MyColor As String) As Integer
Dim TheGreen As String
MyColor = AddZero(6 - Len(MyColor)) & MyColor
TheGreen = "&H" & Left$(Right$(MyColor, 4), 2)
GetGreen = CDec(TheGreen)

End Function
Function RGB2Hex(R As Integer, G As Integer, B As Integer) As String
Dim temp As String
temp = Hex(RGB(R, G, B))
RGB2Hex = AddZero(6 - Len(temp)) & temp
End Function

