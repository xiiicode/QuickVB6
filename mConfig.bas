Attribute VB_Name = "mConfig"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public CfgFile As String
Public CfgName As String

'�������õ�����ģ�飬�������е�VB������Ŀ��ʹ���ˣ�����Ŀ�浵������ʹ����2009�꣬�򵥷��㡣
'֧�ֵĿؼ���TextBox��CheckBox��OptionButton��ComboBox��Label��
'��ʹ��������c# winform��Ŀ����Ҳ����дһ�����ƵĹ����࣬���ڱ��洰����Ŀؼ����á�

'ע�⣺��Ҫ���������õĿؼ���tag��������һ���������Ʋ���Ч��
'ע�⣺��Ҫ���������õĿؼ���tag��������һ���������Ʋ���Ч��
'ע�⣺��Ҫ���������õĿؼ���tag��������һ���������Ʋ���Ч��


'ʹ�÷ǳ��ļ򵥣�ֻ��Ҫ���ı���:
'Private Sub Form_Load()
'    mConfig.LoadControls Me  '�������õ��ؼ�
'End Sub
'
'Private Sub Command1_Click()
'    mConfig.SaveControls Me  '���ؼ����ݱ��浽����
'End Sub



'��ʼ�����
Private Sub init_check()
    If Len(CfgFile) = 0 Then CfgFile = App.Path & "\Config.ini"
    If Len(CfgName) = 0 Then CfgName = "Configure"
End Sub


'д�������ļ�
Public Sub CfgWrite(key As String, value As String)
    Call init_check
    WritePrivateProfileString CfgName, key, value, CfgFile
End Sub


'��ȡ����String
Public Function CfgRead(key As String) As String
    Dim retstr As String
    retstr = String(1024, 0)
    Call init_check
    GetPrivateProfileString CfgName, key, "", retstr, 1024, CfgFile
    CfgRead = Trim(Replace(retstr, Chr(0), ""))
End Function

'��ȡ����Long
Public Function CfgReadLong(key As String) As Long
    On Error GoTo errline
    Call init_check
    CfgReadLong = GetPrivateProfileInt(CfgName, key, 0, CfgFile)
errline:
End Function

'���洰���пؼ���Ϣ������
Public Sub SaveControls(frm As Form)
    On Error Resume Next
    Dim objX As Object
    Dim k As String
    For Each objX In frm.Controls
        k = CStr(objX.Tag)
        If Len(k) > 1 Then
            If TypeOf objX Is TextBox Then '�ı���,�ų�����
                If objX.MultiLine = False Then CfgWrite k, objX.Text
            ElseIf TypeOf objX Is CheckBox Then '��ѡ
                CfgWrite k, objX.value
            ElseIf TypeOf objX Is OptionButton Then '��ѡ
                CfgWrite k, objX.value
            ElseIf TypeOf objX Is ComboBox Then '�����б�
                CfgWrite k, objX.ListIndex
            ElseIf TypeOf objX Is Label Then '��ǩ
                CfgWrite k, objX.Caption
            End If
        End If
    Next
End Sub

'���ش����пؼ�������
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
            If TypeOf objX Is TextBox Then '�ı���,�ų�����
                If objX.MultiLine = False Then
                    objX.Text = s
                End If
            ElseIf TypeOf objX Is CheckBox Then '��ѡ
                If s = "1" Then
                    objX.value = 1
                Else
                    objX.value = 0
                End If
            ElseIf TypeOf objX Is OptionButton Then '��ѡ
                If s = "True" Then
                    objX.value = True
                Else
                    objX.value = False
                End If
            ElseIf TypeOf objX Is ComboBox Then '�����б�
                i = CfgReadLong(k)
                If objX.ListCount > 0 Then
                    objX.ListIndex = IIf(i < objX.ListCount, i, 0)
                End If
            ElseIf TypeOf objX Is Label Then '��ǩ
                i = CfgReadLong(k)
                objX.Caption = i
            End If
        End If
    Next
End Sub

'���浥���ؼ������ã���ȡ�ķ���Ҳ�����ƣ������ò������Ծ�����д��
Public Sub SaveControl(ctl As Variant, vl As String)
    Dim k As String
    If TypeOf ctl Is TextBox Then '�ı���,�ų�����
        If ctl.MultiLine = False Then
            CfgWrite k, ctl.Text
        End If
    ElseIf TypeOf ctl Is CheckBox Then '��ѡ
        CfgWrite k, ctl.value
    ElseIf TypeOf ctl Is OptionButton Then '��ѡ
        CfgWrite k, ctl.value
    ElseIf TypeOf ctl Is ComboBox Then '�����б�
        CfgWrite k, ctl.ListIndex
    ElseIf TypeOf ctl Is Label Then '��ǩ
        CfgWrite k, ctl.Caption
    End If
End Sub

