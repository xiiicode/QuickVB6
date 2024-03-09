Attribute VB_Name = "mCRC32"
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------

'纯算法版
Dim crc32Table(255) As Long
Dim bool_init As Boolean

Public Function ComputeCrc32(buf() As Byte) As Long
    If bool_init = False Then Call InitCrc32
    Dim i As Long, iCRC As Long, lngA As Long, ret As Long
    Dim b() As Byte
    Dim bytT As Byte, bytC As Byte
    b = buf 'StrConv(item, vbFromUnicode)
    iCRC = &HFFFFFFFF
    InitCrc32
    For i = 0 To UBound(b)
        bytC = b(i)
        bytT = (iCRC And &HFF) Xor bytC
        lngA = ((iCRC And &HFFFFFF00) / &H100) And &HFFFFFF
        iCRC = lngA Xor crc32Table(bytT)
    Next
    ret = iCRC Xor &HFFFFFFFF
    Crc32_Byte = ret
End Function

Public Function ComputeCrc32_String(str As String) As Long
    Dim b() As Byte
    b = StrConv(str, vbFromUnicode)
    Crc32_String = Crc32_Byte(b)
End Function

'初始化CRC32表
Private Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    Dim i As Integer, j As Integer, Crc32 As Long, Temp As Long
    For i = 0 To 255
        Crc32 = i
        For j = 0 To 7
            Temp = ((Crc32 And &HFFFFFFFE) / &H2) And &H7FFFFFFF
            If (Crc32 And &H1) Then Crc32 = Temp Xor Seed Else Crc32 = Temp
        Next
        crc32Table(i) = Crc32
    Next
    InitCrc32 = Precondition
End Function

' Public Function LongToHex(lng As Long) As String
'     Dim buff(3) As Byte
'     CopyMemory buff(0), lng, 4
'     Dim i As Byte
'     For i = 0 To UBound(buff)
'         If buff(i) > 15 Then
'             LongToHex = LongToHex & Hex(buff(i))
'         Else
'             LongToHex = LongToHex & "0" & Hex(buff(i))
'         End If
'     Next i
' End Function