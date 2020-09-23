Attribute VB_Name = "modCaps"

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public CapVal(255) As Long 'the cap's value
Public CapSp(255) As Long

Type Color
r As Integer
g As Integer
b As Integer
End Type

Public OldVal(520, 1 To 10) As Byte
Public OldColor(1 To 10) As Color

Public OldI As Long

Public Const MaxFade = 10

Public FPS, TFPS

