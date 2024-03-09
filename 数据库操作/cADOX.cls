VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Public Enum ColLength
    adDefault = 0       '默认值是零（0），既不是 adColFixed 也不是 adColNullable。
    adColFixed = 1      '列长度是固定的,必填字段
    adColNullable = 2   ' 列可以包含 null 值。非必填,允许为空
End Enum
Public Enum FieldType
    adInteger = 3 '四字节的有符号整数 (DBTYPE_I4)。
    adBigInt = 20 '八字节的有符号整数 (DBTYPE_I8)。"
    adUnsignedTinyInt = 17 '一个单字节的无符号整数 (DBTYPE_UI1)。
    adUnsignedSmallInt = 18 '一个双字节的无符号整数 (DBTYPE_UI2)。"
    adUnsignedInt = 19 '四字节的无符号整数 (DBTYPE_UI4)。"
    adUnsignedBigInt = 21 '八字节无符号整数 (DBTYPE_UI8)
    adSingle = 4
    adDouble = 5
    adCurrency = 6
    adDecimal = 14 '具有固定精度和范围的确切数字值 (DBTYPE_DECIMAL)。
    adNumeric = 131 '具有固定精度和范围的确切数字值 (DBTYPE_NUMERIC)。
    adBoolean = 11
    adError = 10
    adFileTime = 64 '64 位的值，表示从 1601 年 1 月 1 日开始的 100 个十亿分之一秒间隔的数量 (DBTYPE_FILETIME)。
    adGuid = 72 '全局唯一标识符 (GUID) (DBTYPE_GUID)。"
    adDate = 7
    adDBDate = 133 '日期值 (yyyymmdd) (DBTYPE_DBDATE)。
    adDBTime = 134 '时间值 (hhmmss) (DBTYPE_DBTIME)。
    adDBTimestamp = 135 '日期/时间戳（yyyymmddhhmmss 加十亿分之一的小数）(DBTYPE_DBTIMESTAMP)。
    adBSTR = 8 '以 Null 终止的 Unicode 字符串 (DBTYPE_BSTR)。
    adWChar = 130 '以 Null 终止的 Unicode 字符串 (DBTYPE_WSTR)。
    adChar = 129 '一个字符串值 (DBTYPE_STR)。
    adVarChar = 200
    adLongVarChar = 201
    adVarWChar = 202
    adLongVarWChar = 203 '这个是对应的备注类型,一个以 Null 终止的长 Unicode 字符串值（仅限于 Parameter 对象）。"
    adBinary = 128
    adVarBinary = 204
    adLongVarBinary = 205
End Enum
Private Type FIELD_PROP
    dwname As String
    dwtype As FieldType
    dwsize As Long
    dwNull As ColLength
End Type
    
Dim dwErr           As String
Dim AddFields()     As FIELD_PROP


Public Property Get LastError() As String
    LastError = dwErr
End Property

'连接数据库
Public Function ConnectionMDB(mdbfile As String, Optional psw As String) As Boolean
    On Error GoTo errline
    Dim mADO            As Object
    Set mADO = CreateObject("ADOX.Catalog")
    Dim dwStrConn As String
    dwStrConn = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & mdbfile
    If Len(psw) > 0 Then
        dwStrConn = dwStrConn & ";Jet OLEDB:Database Password=" & psw
    End If
    mADO.ActiveConnection = dwStrConn
    ConnectionMDB = True
    Exit Function
errline:
    dwErr = Err.number & "," & Err.Description
End Function

'创建数据库
Public Function CreateDatabase(mdbfile As String, Optional psw As String) As Boolean
    On Error GoTo errline
    Dim mADO            As Object
    Set mADO = CreateObject("ADOX.Catalog")
    Dim dwStrConn As String
    dwStrConn = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & mdbfile
    If Len(psw) > 0 Then
        dwStrConn = dwStrConn & ";Jet OLEDB:Database Password=" & psw
    End If
    mADO.Create dwStrConn '如果成功，新的 Connection 对象将被赋给 ActiveConnection 属性。
    DoEvents
'    mADO.ActiveConnection = dwStrConn
    DoEvents
    CreateDatabase = True
    Exit Function
errline:
    dwErr = Err.number & "," & Err.Description
End Function

'添加字段
Public Function AddField(name_ As String, type_ As FieldType, Optional size_ As Long, Optional colNull As ColLength = adDefault)
    Dim i As Long
    i = UBound(AddFields) + 1
    ReDim Preserve AddFields(i)
    AddFields(i).dwname = Trim(name_)
    AddFields(i).dwtype = type_
    AddFields(i).dwsize = size_
    AddFields(i).dwNull = colNull
End Function
'添加表
Public Function CreateTable(mdbfile As String, tb_name As String, Optional psw As String) As Boolean
    Dim Col As Object
    Dim Tbl As Object
    Dim Ado As Object
    On Error GoTo errline
    Dim i As Long
    Dim dwStrConn As String
    Set Ado = CreateObject("ADOX.Catalog")
    Set Col = CreateObject("ADOX.Column") '集合
    Set Tbl = CreateObject("ADOX.Table")
    dwStrConn = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & mdbfile
    If Len(psw) > 0 Then
        dwStrConn = dwStrConn & ";Jet OLEDB:Database Password=" & psw
    End If
    Ado.ActiveConnection = dwStrConn

    Tbl.Name = tb_name
    Col.ParentCatalog = Ado
    '首先创建一个自增字段id
    Col.Type = adInteger '必须先设置字段类型
    Col.Name = "id"
    Col.Properties("Jet OLEDB:Allow Zero Length").value = False
    Col.Properties("AutoIncrement").value = True '自增字段
    Col.Properties("Nullable").value = True
    Tbl.Columns.Append Col, FieldType.adInteger, 0
    '添加其他的字段
    For i = 0 To UBound(AddFields)
        If Len(AddFields(i).dwname) > 0 And AddFields(i).dwtype <> 0 Then
            Tbl.Columns.Append AddFields(i).dwname, AddFields(i).dwtype, AddFields(i).dwsize
        End If
    Next i
    For i = 0 To UBound(AddFields)
        If Len(AddFields(i).dwname) > 0 And AddFields(i).dwtype <> 0 And AddFields(i).dwNull = adColNullable Then
            Tbl.Columns(AddFields(i).dwname).Attributes = ColLength.adColNullable '该字段允许为空
        End If
    Next i
    Ado.Tables.Append Tbl '建立数据表
    CreateTable = True
    Exit Function
errline:
    dwErr = Err.number & "," & Err.Description
End Function
'调试显示
Public Function DebugPrintStructure(mdbfile As String, Optional psw As String) As String
    On Error GoTo errline
    Dim dwT
    Dim dwF
    Dim prop
    Dim Col As Object
    Dim i As Long
    Dim dwStrConn As String
    Dim mADO            As Object
    Set mADO = CreateObject("ADOX.Catalog")
    Set Col = CreateObject("ADOX.Column") '集合
    dwStrConn = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & mdbfile
    If Len(psw) > 0 Then
        dwStrConn = dwStrConn & ";Jet OLEDB:Database Password=" & psw
    End If
    mADO.ActiveConnection = dwStrConn
    Dim s As String
    For Each dwT In mADO.Tables
'        s = s & dwT.Name & vbCrLf  '这里如果去掉就会显示系统的表等
        For Each dwF In dwT.Columns
            s = s & dwT.Name & " ->" & dwF.Type & " > " & dwF.DefinedSize & " > " & dwF.Name & vbCrLf
            '显示字段详细的属性
'            For Each prop In dwF.Properties
'                 s = s & prop.Name & " : " & prop.Value & vbCrLf
'            Next
        Next
    Next
    DebugPrintStructure = s
    Exit Function
errline:
    dwErr = Err.number & "," & Err.Description
End Function
'初始化
Public Function Initialize() As Boolean
    On Error GoTo errline
    ReDim AddFields(0)
    dwErr = ""
    Initialize = True
    Exit Function
errline:
    dwErr = Err.number & "," & Err.Description
End Function
Private Sub Class_Initialize()
    Call Initialize
End Sub

