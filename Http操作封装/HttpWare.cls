VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpWare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Dim objScript1
Dim objScript2
Dim init1 As Boolean
Dim init2 As Boolean


'=============================================================
'JavaScript执行
'=============================================================
Public Function JavaScriptExecute(code As String) As String
    On Error GoTo errline
    If init1 = False Then '初始化不再放进Class_Initialize()
        init1 = True
        Set objScript1 = CreateObject("MSScriptControl.ScriptControl")
        objScript1.AllowUI = True '检测是否允许运行用户的接口元素。如果为False，则诸如消息框之类的界面元素不可见。
        objScript1.language = "JavaScript"
    End If
    JavaScriptExecute = objScript1.eval(code)
    Exit Function
errline:
    Debug.Print "【Err.JavaScriptExecute" & Err.Description
End Function

Public Function JsonParse(json As String, exp As String) As String
    On Error GoTo errline
    If init2 = False Then '初始化不再放进Class_Initialize()
        init2 = True
        Set objScript2 = CreateObject("MSScriptControl.ScriptControl")
        objScript2.language = "JScript"
    End If
    JsonParse = objScript2.eval("JSON_DATA=" & json & ";JSON_DATA." & exp & ";")
    Exit Function
errline:
    Debug.Print "【Err.JsonParse】"
End Function

Public Function TimestampToDate(ts As Long) As String
    TimestampToDate = DateAdd("s", ts, "1970/01/01 08:00:00")
End Function

Public Function DateToTimestamp(Optional strDate As String = "") As Long
    Dim ts As Long

    If IsDate(strDate) Then
        ts = DateDiff("s", "1970/01/01 08:00:00", strDate)
    Else
        ts = DateDiff("s", "1970/01/01 08:00:00", Now)
    End If
    DateToTimestamp = ts
End Function


'=============================================================
'字符串的提取 如:MidEx("is[12345]ok","[","]") 结果为"12345"
'=============================================================
Public Function MidEx(str As String, strPrefix As String, strSuffix As String, Optional starindex As Long = 1) As String
    Dim i As Long
    Dim j As Long
    If Len(strPrefix) = 0 Then
        i = 1
    Else
        If starindex < 1 Then starindex = 1
        i = InStr(starindex, str, strPrefix)
    End If
    If i = 0 Then Exit Function
    j = InStr(i + Len(strPrefix), str, strSuffix)
    If j = 0 Or j <= i + Len(strPrefix) Then Exit Function
    MidEx = Mid(str, i + Len(strPrefix), j - i - Len(strPrefix))
    starindex = j
End Function

'=============================================================
'产生随机数
'=============================================================
Public Function Random() As String
    Dim D As Double
    D = Rnd * 1
    Random = Format(D, "0.000000000000000")
End Function

Public Function ReadString(file_path As String, OutStr As String) As Boolean
    Dim buff() As Byte
    If ReadBytes(file_path, buff) Then
        OutStr = StrConv(buff, vbUnicode)
        ReadString = True
    End If
End Function


Public Function ReadBytes(file_path As String, outBit() As Byte) As Boolean
    On Error GoTo errline
    Dim fn As Integer
    Dim fz As Long
    fn = FreeFile()
    fz = FileLen(file_path) '获取文件长度
    ReDim outBit(fz - 1)
    Open file_path For Binary As #fn
        Get #fn, , outBit
    Close #fn
    ReadBytes = True
    Exit Function
errline:
End Function

Public Function WriteString(file_path As String, putinStr, Optional isAppend As Boolean = False) As Boolean
    On Error GoTo errline
    Dim fn As Integer
    fn = FreeFile()
    If isAppend = False Then
        Open file_path For Output As #fn
    Else
        Open file_path For Append As #fn
    End If
    Print #fn, putinStr
    Close #fn
    WriteString = True
    Exit Function
errline:
End Function

Public Function WriteBytes(file_path As String, putBit() As Byte, Optional isAppend As Boolean = False) As Boolean
    On Error GoTo errline
    Dim fn As Integer
    fn = FreeFile()
    Open file_path For Binary As #fn
        If isAppend Then
            Put #fn, LOF(fn) + 1, putBit
        Else
            Put #fn, , putBit
        End If
    Close #fn
    WriteBytes = True
    Exit Function
errline:
End Function