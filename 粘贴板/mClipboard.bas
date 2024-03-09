Attribute VB_Name = "mClipboard"
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
'┏┳━━━━━━━━━━━━━━━━━━━━━━━━━
'★┃2013/10/29 4:41:51
'★┃将字符设置到粘贴板
'┗┻━━━━━━━━━━━━━━━━━━━━━━━━━
Public Function SetString(str As String) As Boolean
    On Error GoTo errLine
    Clipboard.Clear
    DoEvents
    Clipboard.SetText str
    DoEvents
    SetString = True
    Exit Function
errLine:
    SetString = False
End Function
'┏┳━━━━━━━━━━━━━━━━━━━━━━━━━
'★┃2013/10/29 4:41:51
'★┃获取粘贴板的字符
'┗┻━━━━━━━━━━━━━━━━━━━━━━━━━
Public Function GetString(str As String) As Boolean
    On Error GoTo errLine
    If Clipboard.GetFormat(vbCFText) Then
        str = Clipboard.GetText
        GetString = True
    End If
    Exit Function
errLine:
    GetString = False
End Function
'┏┳━━━━━━━━━━━━━━━━━━━━━━━━━
'★┃2013/10/29 4:41:51
'★┃将图像复制到粘贴板
'┗┻━━━━━━━━━━━━━━━━━━━━━━━━━
Public Function SetImage(img As IPicture) As Boolean
    On Error GoTo errLine
    Clipboard.Clear
    DoEvents
    Clipboard.SetData img, vbCFBitmap
    DoEvents
    SetImage = True
    Exit Function
errLine:
    SetImage = False
End Function
'┏┳━━━━━━━━━━━━━━━━━━━━━━━━━
'★┃2013/10/29 4:41:51
'★┃获取粘贴板的图像
'┗┻━━━━━━━━━━━━━━━━━━━━━━━━━
Public Function GetImage(img As IPictureDisp) As Boolean
    On Error GoTo errLine
    If Clipboard.GetFormat(vbCFBitmap) Then
        Set img = Clipboard.GetData
        GetImage = True
    End If
    Exit Function
errLine:
    GetImage = False
End Function

