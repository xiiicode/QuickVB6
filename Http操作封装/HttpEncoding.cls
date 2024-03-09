VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
'GZIP
'--------------------------------------------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function InitDecompression Lib "gzip.dll" () As Long
Private Declare Function CreateDecompression Lib "gzip.dll" (ByRef context As Long, ByVal flags As Long) As Long
Private Declare Function DestroyDecompression Lib "gzip.dll" (ByRef context As Long) As Long
Private Declare Function Decompress Lib "gzip.dll" (ByVal context As Long, inBytes As Any, ByVal input_size As Long, outBytes As Any, ByVal output_size As Long, ByRef input_used As Long, ByRef output_used As Long) As Long
Private Const offset As Long = &H8
Private Const GZIP_LVL As Long = 1
'--------------------------------------------------
'--------------------------------------------------
'UTF-8
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
Dim Ware As New HttpWare


'=============================================================
'=============================================================
'=============================================================
'utf8转unicode
Public Function UTF8_Decode_Bytes(ByRef Utf() As Byte) As Byte()
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim BT() As Byte
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2 - 1
    ReDim BT(lBufferSize)
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, VarPtr(BT(0)), lBufferSize + 1)
    If lRet <> 0 Then
        ReDim Preserve BT(lRet - 1)
    Else
        ReDim BT(0)
    End If
    UTF8_Decode_Bytes = BT
End Function

'utf8转unicode
Public Function UTF8_Decode(ByRef Utf() As Byte) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    On Error GoTo errline:
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    UTF8_Decode = String$(lBufferSize, Chr(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(UTF8_Decode), lBufferSize)
    If lRet <> 0 Then
        UTF8_Decode = Left(UTF8_Decode, lRet)
    End If
    Exit Function
errline:
    UTF8_Decode = ""
End Function
'unicode转utf8
Public Function UTF8_Encode_Bytes(ByVal UCS As String) As Byte()
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    lLength = Len(UCS)
    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)
    If lResult <> 0 Then
    lResult = lResult - 1
    ReDim Preserve abUTF8(lResult)
    UTF8_Encode_Bytes = abUTF8
    End If
End Function



'----------------------------------------------------------
'下面列出了Javascript中三个URL编码函数的安全字符 (即函数不会对这些字符进行编码)
'escape（69个）：*/@+-._0-9a-zA-Z
'encodeURI（82个）：!#$&'()*+,/:;=?@-._~0-9a-zA-Z
'encodeURIComponent（71个）：!'()*-._~0-9a-zA-Z
'----------------------------------------------------------

'url编码,利用js escape方式
Public Function Escape(str As String) As String
    Dim i As Long
    Dim t As String
    Dim js As String
    i = InStr(str, "'")
    If i > 0 Then
        t = Chr(34)
    Else
        t = "'"
    End If
    js = "function enc(){var s = " & t & str & t & ";return escape(s);}enc();"
    Escape = Ware.JavaScriptExecute(js)
End Function
'url解码,利用js unescape方式
Public Function Unescape(str As String) As String
    Dim i As Long
    Dim t As String
    Dim js As String
    i = InStr(str, "'")
    If i > 0 Then
        t = Chr(34)
    Else
        t = "'"
    End If
    js = "function enc(){var s = " & t & str & t & ";return unescape(s);}enc();"
    Unescape = Ware.JavaScriptExecute(js)
End Function
'url编码,利用js encodeURIComponent方式
Public Function EncodeURIComponent(str As String) As String
    Dim i As Long
    Dim t As String
    Dim js As String
    i = InStr(str, "'")
    If i > 0 Then
        t = Chr(34)
    Else
        t = "'"
    End If
    js = "function enc(){var s = " & t & str & t & ";return encodeURIComponent(s);}enc();"
    EncodeURIComponent = Ware.JavaScriptExecute(js)
End Function
'url解码,利用js decodeURIComponent方式
Public Function DecodeURIComponent(str As String) As String
    Dim i As Long
    Dim t As String
    Dim js As String
    i = InStr(str, "'")
    If i > 0 Then
        t = Chr(34)
    Else
        t = "'"
    End If
    js = "function enc(){var s = " & t & str & t & ";return decodeURIComponent(s);}enc();"
    DecodeURIComponent = Ware.JavaScriptExecute(js)
End Function

'=============================================================
'=============================================================
'=============================================================
'CJK编码,也叫做Unicode编码或中日韩字符集,中国GBK和GB2312标准
Public Function CJK_Encode(str As String, Optional isSlashStart As Boolean = True) As String
    Dim i As Long
    Dim j As Integer
    Dim rs As String
    Dim Si As String
    Dim hx As String
    Dim uSign As String
    If isSlashStart Then
        uSign = "\U"
    Else
        uSign = "%U"
    End If
    For i = 1 To Len(str)
        Si = Mid(str, i, 1)
        j = Asc(Si)
        If j < 0 Or j > 126 Then
            hx = Hex(j)
            rs = rs & uSign & hx
        Else
            rs = rs & Si
        End If
    Next i
    CJK_Encode = rs
End Function
'CJK解码
Public Function CJK_Decode(str As String, Optional isSlashStart As Boolean = True) As String
    Dim bs() As Byte
    Dim br() As Byte
    Dim i As Long
    Dim rn As Long
    Dim uSign As Byte
    If isSlashStart Then
        uSign = 92 '\
    Else
        uSign = 32 '%
    End If
    bs = StrConv(str, vbFromUnicode)
    ReDim br(Len(str))
    For i = 0 To UBound(bs)
        If bs(i) = uSign Then
            If bs(i + 1) = 85 Or bs(i + 1) = 117 Then 'U or u
                br(rn) = CByte("&h" & Chr(bs(i + 2)) & Chr(bs(i + 3)))
                rn = rn + 1
                br(rn) = CByte("&h" & Chr(bs(i + 4)) & Chr(bs(i + 5)))
                rn = rn + 1
                i = i + 5
            End If
        Else
            br(rn) = bs(i)
            rn = rn + 1
        End If
    Next i
    CJK_Decode = Trim(StrConv(br, vbUnicode))
End Function



'=============================================================
'=============================================================
'=============================================================
'gzip解压,解压结果在原数据
'解压缩数组
Public Function UnGzip(arrBit() As Byte) As Boolean
    On Error GoTo errline
    Dim SourceSize  As Long
    Dim Buffer()    As Byte
    Dim lReturn    As Long
    Dim outUsed    As Long
    Dim inUsed      As Long
    Dim chandle As Long
    If arrBit(0) <> &H1F Or arrBit(1) <> &H8B Or arrBit(2) <> &H8 Then
        Exit Function '不是GZIP数据的字节流
    End If
    '获取原始长度
    'GZIP格式的最后4个字节表示的是原始长度
    '与最后4个字节相邻的4字节是CRC32位校验,用于比对是否和原数据匹配
    lReturn = UBound(arrBit) - 3
    CopyMemory SourceSize, arrBit(lReturn), 4
'    这里的判断是因为:(维基)一个压缩数据集包含一系列的block（块），只要未压缩数据大小不超过65535字节，块的大小是任意的。
'    GZIP基本头是10字节
    If SourceSize > 65535 Or 65535 < 10 Then
        SourceSize = 102400 '应该是跳出,但是个人还是想申请100KB空间尝试一下
    '    Exit Function
    Else
        SourceSize = SourceSize + 1
    End If
    ReDim Buffer(SourceSize) As Byte
    '创建解压缩进程
    InitDecompression
    CreateDecompression chandle, GZIP_LVL  '创建
    '解压缩数据
    Decompress ByVal chandle, arrBit(0), UBound(arrBit) + 1, Buffer(0), SourceSize + 1, inUsed, outUsed
    If outUsed <> 0 Then
        DestroyDecompression chandle
        ReDim arrBit(outUsed - 1)
        CopyMemory arrBit(0), Buffer(0), outUsed
        UnGzip = True
    End If
    Exit Function
errline:
    Debug.Print "UnGzip ERROR:" & Err.Number & "," & Err.Description
End Function


'Base64 编码
Public Function Base64_Encode(bytes() As Byte) As String
    On Error GoTo over
    Dim buf() As Byte, length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(bytes) + 1) Mod 3 '除以3的余数
    length = UBound(bytes) + 1 - mods
    ReDim buf(length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To length - 1 Step 3
        buf(i / 3 * 4) = (bytes(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (bytes(i) And &H3) * &H10 + (bytes(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (bytes(i + 1) And &HF) * &H4 + (bytes(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = bytes(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(length / 3 * 4) = (bytes(length) And &HFC) / &H4
        buf(length / 3 * 4 + 1) = (bytes(length) And &H3) * &H10
        buf(length / 3 * 4 + 2) = 64
        buf(length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(length / 3 * 4) = (bytes(length) And &HFC) / &H4
        buf(length / 3 * 4 + 1) = (bytes(length) And &H3) * &H10 + (bytes(length + 1) And &HF0) / &H10
        buf(length / 3 * 4 + 2) = (bytes(length + 1) And &HF) * &H4
        buf(length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        Base64_Encode = Base64_Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function

'Base64 解码
Public Function Base64_Decode(str As String) As Byte()
    On Error GoTo over
    Dim OutStr() As Byte, i As Long, j As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    If InStr(1, str, "=") <> 0 Then str = Left(str, InStr(1, str, "=") - 1) '判断Base64真实长度,除去补位
    Dim length As Long, mods As Long
    mods = Len(str) Mod 4
    length = Len(str) - mods
    ReDim OutStr(length / 4 * 3 - 1 + Switch(mods = 0, 0, mods = 2, 1, mods = 3, 2))
    For i = 1 To length Step 4
        Dim buf(3) As Byte
        For j = 0 To 3
            buf(j) = InStr(1, B64_CHAR_DICT, Mid(str, i + j, 1)) - 1 '根据字符的位置取得索引值
        Next
        OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
        OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
        OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
    Next
    If mods = 2 Then
        OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(str, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(str, length + 2, 1)) - 1) And &H30) / 16
    ElseIf mods = 3 Then
        OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(str, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(str, length + 2, 1)) - 1) And &H30) / 16
        OutStr(length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(str, length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(str, length + 3, 1)) - 1) And &H3C) / &H4
    End If
    Base64_Decode = OutStr
over:
End Function

