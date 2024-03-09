Attribute VB_Name = "mCRC32API"
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------

'API版
Private Declare Function RtlComputeCrc32 Lib "ntdll.dll" (ByVal dwInitial As Long, ByVal pData As Long, ByVal iLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
 

Public Function ComputeCrc32(byt() As Byte) As Long
    Dim i As Long
    i = RtlComputeCrc32(0, VarPtr(byt(0)), UBound(byt) + 1)
    GetBytesCrc32 = i
End Function
 
Public Function ComputeCrc32_String(str As String) As Long
    Dim b() As Byte
    Dim i As Long
    b = StrConv(str, vbFromUnicode)
    i = RtlComputeCrc32(0, VarPtr(b(0)), UBound(b) + 1)
    GetStringCrc32 = i
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