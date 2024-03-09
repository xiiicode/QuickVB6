Attribute VB_Name = "mConfig"
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public CfgFile As String  '配置文件的目录，默认是程序目录下 config.ini
Public CfgName As String  '配置项的应用域名称，默认为 Configure

'初始化检测
Private Sub init_check()
    If Len(CfgFile) = 0 Then CfgFile = App.Path & "\Config.ini"
    If Len(CfgName) = 0 Then CfgName = "Configure"
End Sub


'写入配置文件
Public Sub CfgWrite(key As String, value As String)
    Call init_check
    WritePrivateProfileString CfgName, key, value, CfgFile
End Sub


'读取配置String
Public Function CfgRead(key As String) As String
    Dim retstr As String
    retstr = String(1024, 0)
    Call init_check
    GetPrivateProfileString CfgName, key, "", retstr, 1024, CfgFile
    CfgRead = Trim(Replace(retstr, Chr(0), ""))
End Function

'读取配置Long
Public Function CfgReadLong(key As String) As Long
    On Error GoTo errline
    Call init_check
    CfgReadLong = GetPrivateProfileInt(CfgName, key, 0, CfgFile)
errline:
End Function

'保存窗体中控件信息到配置
Public Sub SaveControls(frm As Form)
    On Error Resume Next
    Dim objX As Object
    Dim k As String
    For Each objX In frm.Controls
        k = CStr(objX.Tag)
        If Len(k) > 1 Then
            If TypeOf objX Is TextBox Then '文本框,排除多行
                If objX.MultiLine = False Then CfgWrite k, objX.Text
            ElseIf TypeOf objX Is CheckBox Then '多选
                CfgWrite k, objX.value
            ElseIf TypeOf objX Is OptionButton Then '单选
                CfgWrite k, objX.value
            ElseIf TypeOf objX Is ComboBox Then '下拉列表
                CfgWrite k, objX.ListIndex
            ElseIf TypeOf objX Is Label Then '标签
                CfgWrite k, objX.Caption
            End If
        End If
    Next
End Sub

'加载窗体中控件的配置
Public Sub LoadControls(frm As Form)
    On Error Resume Next
    Dim objX As Object
    Dim s As String
    Dim k As String
    Dim i As Long
    Dim f As String
    f = frm.Name
    For Each objX In frm.Controls
        k = CStr(objX.Tag)
        If Len(k) > 1 Then
            s = CfgRead(k)
            If TypeOf objX Is TextBox Then '文本框,排除多行
                If objX.MultiLine = False Then
                    objX.Text = s
                End If
            ElseIf TypeOf objX Is CheckBox Then '多选
                If s = "1" Then
                    objX.value = 1
                Else
                    objX.value = 0
                End If
            ElseIf TypeOf objX Is OptionButton Then '单选
                If s = "True" Then
                    objX.value = True
                Else
                    objX.value = False
                End If
            ElseIf TypeOf objX Is ComboBox Then '下拉列表
                i = CfgReadLong(k)
                If objX.ListCount > 0 Then
                    objX.ListIndex = IIf(i < objX.ListCount, i, 0)
                End If
            ElseIf TypeOf objX Is Label Then '标签
                i = CfgReadLong(k)
                objX.Caption = i
            End If
        End If
    Next
End Sub

'保存单个控件的配置，读取的方法也是类似，几乎用不到所以就懒得写了
Public Sub SaveControl(ctl As Variant, vl As String)
    Dim k As String
    If TypeOf ctl Is TextBox Then '文本框,排除多行
        If ctl.MultiLine = False Then
            CfgWrite k, ctl.Text
        End If
    ElseIf TypeOf ctl Is CheckBox Then '多选
        CfgWrite k, ctl.value
    ElseIf TypeOf ctl Is OptionButton Then '单选
        CfgWrite k, ctl.value
    ElseIf TypeOf ctl Is ComboBox Then '下拉列表
        CfgWrite k, ctl.ListIndex
    ElseIf TypeOf ctl Is Label Then '标签
        CfgWrite k, ctl.Caption
    End If
End Sub

