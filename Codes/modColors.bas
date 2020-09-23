Attribute VB_Name = "modColors"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Function GetHexColor(ByVal PC)
r = Hex(RedC(PC))
If Len(r) < 2 Then r = String(2 - Len(r), "0") & r
g = Hex(GreenC(PC))
If Len(g) < 2 Then g = String(2 - Len(g), "0") & g
b = Hex(BlueC(PC))
If Len(b) < 2 Then b = String(2 - Len(b), "0") & b
GetHexColor = r & g & b
End Function

'------------------------------------------
'The folowing three functions will return
'the RED, BLUE and GREEN (RGB) values from
'any color.
'Simply they are the inverse function of RGB

Function RedC(c)
pix& = c
RedC = pix& Mod 256
End Function
Function BlueC(c)
pix& = c
BlueC = (pix& And &HFF0000) / 65536
End Function
Function GreenC(c)
pix& = c
GreenC = ((pix& And &HFF00FF00) / 256&)
End Function
'------------------------------------------

Function GetRandom() As String
GetRandom = Rnd * 10
i = InStr(1, GetRandom, ".")
If i > 0 Then
GetRandom = Left(GetRandom, i - 1) & Mid(GetRandom, i + 1)
End If
GetRandom = Format(GetRandom, "00000000") & Timer
i = InStr(1, GetRandom, ".")
If i > 0 Then
GetRandom = Left(GetRandom, i - 1) + Mid(GetRandom, i + 1)
End If
End Function

