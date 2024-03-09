VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Public Enum ConnectMode
    adModeUnknown = 0           '默认值。权限尚未设置或不能确定权限。
    adModeRead = 1              '只读权限。
    adModeWrite = 2             '只写权限。
    adModeReadWrite = 3         '读/写权限。
    adModeShareDenyRead = 4     '禁止其他人以读权限打开连接。
    adModeShareDenyWrite = 8    '禁止其他人以写权限打开连接。
    adModeShareExclusive = 12   '禁止其他人打开连接。
    adModeShareDenyNone = 16    '允许其他人以任何权限打开连接。不拒绝其他人的读或写访问。
    adModeRecursive = &H4000000 '与 adModeShareDenyNone, adModeShareDenyWrite 或 adModeShareDenyRead 一起使用，对当前 Record 的所有子记录设置权限。
End Enum
Public Enum ConnectState
    adStateClosed = 0           '对象已关闭。
    adStateOpen = 1             '对象已打开。
    adStateConnecting = 2       '对象正在连接。
    adStateExecuting = 4        '对象正在执行命令。
    adStateFetching = 8         '正在检索对象的行。
End Enum
Public Enum CursorType '打开 Recordset 时应该使用的游标类型
    AdOpenForwardOnly = 0       '（默认值）打开仅向前类型游标。
    AdOpenKeyset = 1            '打开键集类型游标
    AdOpenDynamic = 2           '打开动态类型游标
    AdOpenStatic = 3            '打开静态类型游标
End Enum
Public Enum LockType '打开 Recordset 时应该使用的锁定（并发）类型
    AdLockReadOnly = 1          '（默认值）只读 — 不能改变数据。
    AdLockPessimistic = 2       '保守式锁定（逐个） — 提供者完成确保成功编辑记录所需的工作，通常通过在编辑时立即锁定数据源的记录。
    AdLockOptimistic = 3        '开放式锁定（逐个） — 提供者使用开放式锁定，只在调用Update 方法时才锁定记录。
    AdLockBatchOptimistic = 4   '开放式批更新—用于批更新模式（与立即更新模式相对）。
End Enum
Public Enum OptionsType '可选，长整型值，用于指示提供者如何计算 Source 参数（如果它代表的不是 Command 对象），或从以前保存 Recordset 的文件中恢复 Recordset。可为下列常量之一
    adCmdUnspecifed = -1        '未描述CommandType属性
    adCmdText = 1               '指示提供者应该将 Source 作为命令的文本定义来计算。
    adCmdTable = 2              '指示 ADO 生成 SQL 查询以便从在 Source 中命名的表中返回所有行。
    adCmdTableDirect = 512      '指示提供者更改从在 Source 中命名的表中返回所有行。
    adCmdStoredProc = 4         '指示提供者应该将 Source 视为存储过程。
    adCmdUnknown = 8            '指示 Source 参数中的命令类型为未知。
    adCmdFile = 256             '指示应从在 Source 中命名的文件中恢复保留（保存的）Recordset。
    'adAsyncExecute             '指明指令是否需要异步执行。这个值不能与CommandTypeEnum 之中的adCmdTableDirect一起使用
    'adAsyncFetch               '指示在提取 Initial Fetch Size 属性中指定的初始数量后，应该异步提取所有剩余的行。如果所需的行尚未提取，主要的线程将被堵塞直到行重新可用。
    'adAsyncFetchNonBlocking    '指示主要线程在提取期间从未堵塞。如果所请求的行尚未提取，当前行自动移到文件末尾。
End Enum
Dim dwErr           As String
Dim Conn            As Object
Public Rs    As Object

Public Property Get LastError() As String
    LastError = dwErr
End Property

'连接数据库
Public Function ConnectionMDB(mdbfile As String, Optional psw As String) As Boolean
    On Error GoTo errline
    dwErr = ""
    Call CloseConn '先尝试关闭,避免内存泄漏
    Dim strConn As String
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfile
    If Len(psw) <> 0 Then strConn = strConn & ";Persist Security Info=True;Jet OLEDB:Database Password=" & psw
    Conn.Open strConn
    ConnectionMDB = True
    Exit Function
errline:
    dwErr = Err.Number & "," & Err.Description
End Function
'连接MYSQL数据库
Public Function ConnectionMySQL(uid As String, pwd As String, db_name As String, Optional dwMode As ConnectMode = adModeReadWrite, Optional Host As String = "127.0.0.1", Optional port As Integer = 3306, Optional connTimeout As Long = 15) As Boolean
    On Error GoTo errline
    dwErr = ""
    Call CloseConn '先尝试关闭,避免内存泄漏
    Conn.mode = dwMode 'adModeReadWrite
    Conn.ConnectionString = "Driver={MySQL ODBC 5.3 ANSI Driver}" _
                        & ";Persist Security Info=True" _
                        & ";SERVER=" & Host _
                        & ";PORT=" & port _
                        & ";DATABASE=" & db_name _
                        & ";UID=" & uid _
                        & ";PWD=" & pwd _
                        & ";OPTION=3;Charset=gbk"
    Conn.ConnectionTimeout = connTimeout '默认为15，如果等于0，那么无限等待连接
    Conn.Open
    If Conn.State = adStateOpen Then
        ConnectionMySQL = True
    Else
        dwErr = "数据库连接失败[State=" & Conn.State & "]"
    End If
    Exit Function
errline:
    dwErr = Err.Number & ":" & Err.Description
End Function



Public Function SQLExecute(sql_ As String) As Boolean
    On Error GoTo errline
    dwErr = ""
    Conn.Execute sql_
    SQLExecute = True
    Exit Function
errline:
    dwErr = Err.Number & "," & Err.Description
End Function

'打开数据集 注意已经包含了是否有结果的判断
'正确的使用方式应该是:
'-------------------------------------------------
'If RecordsetOpen ("select * from table1 order by id") = True then
'    showlog "打开完成"
'Else
'    If len(LastError) <> 0 Then
'        showlog "发生错误:" & LastError
'    Else
'        showlog "没有数据"
'    End If
'End If
'-------------------------------------------------
Public Function RecordsetOpen(sql_ As String, Optional dwCursor As CursorType = AdOpenStatic, Optional dwLock As LockType = AdLockOptimistic, Optional dwOpt As OptionsType = adCmdText) As Boolean
    'adOpenDynamic, adLockOptimistic, adCmdText
    On Error GoTo errline
    dwErr = ""
    '----------------------------------------------------
    'recordset.Open Source, ActiveConnection, CursorType, LockType, Options
    '----------------------------------------------------
    Rs.Open sql_, Conn, dwCursor, dwLock, dwOpt
    If Rs.EOF Or Rs.BOF Then Exit Function
    RecordsetOpen = True
    Exit Function
errline:
    dwErr = Err.Number & "," & Err.Description



End Function

'初始化和销毁
Public Function Initialize() As Boolean
    On Error GoTo errline
    dwErr = ""
    Set Conn = CreateObject("ADODB.Connection")
    Set Rs = CreateObject("ADODB.Recordset")
    Initialize = True
    Exit Function
errline:
    dwErr = Err.Number & "," & Err.Description
End Function
Private Sub Class_Initialize()
    Call Initialize
End Sub
Public Sub CloseConn()
    On Error Resume Next
    Conn.Close
End Sub
Private Sub Class_Terminate()
    Call CloseConn
    Set Conn = Nothing
    Set Rs = Nothing
End Sub


'读取较大值的时候,参考下面的方法
'l = rs.Fields(1).ActualSize
's = rs.Fields(1).GetChunk(l)

'MoveFirst      方法    移动到记录首
'MoveLast       方法    移动到记录尾
'MoveNext       方法    移动到下一条记录
'MovePrevious   方法    移动到上一条记录
'Move           方法    向前或向后移动N条记录
'Bookmark       属性    返回标识当前记录的书签或设置到指定的书签位置
'BOF            属性    若当前记录在首条之前,返回TRUE,反之FALSE
'EOF            属性    若当前记录在尾条之后,返回TRUE,反之FALSE
'=============================================
'添加示例:
'=============================================
'Private Sub Command1_Click()
'    Dim db As New cADODB
'    If db.ConnectionMDB(App.Path & "\test.mdb", "abc123") = False Then
'        Debug.Print db.LastError
'        Exit Sub
'    End If
'    If db.RecordsetOpen("select * from mytable") = False Then
'        '由于是添加操作,所以只在意是否错误,不在意是否为空
'        If Len(db.LastError) <> 0 Then Debug.Print db.LastError
'    End If
'    db.Rs.AddNew
'    db.Rs.Fields("nike") = "mike"
'    db.Rs.Fields("uid") = 88888
'    db.Rs.Fields("psw") = "mima123"
'    db.Rs.Update
'    Debug.Print "OK"
'End Sub

'=============================================
'查找显示示例:
'=============================================
'Private Sub Command1_Click()
'    Dim db As New cADODB
'    If db.ConnectionMDB(App.Path & "\test.mdb", "abc123") = False Then
'        Debug.Print db.LastError
'        Exit Sub
'    End If
'    If db.RecordsetOpen("SELECT * FROM mytable WHERE id>70") = False Then
'        If Len(db.LastError) <> 0 Then
'            Debug.Print db.LastError
'        Else
'            Debug.Print "无数据集"
'        End If
'        Exit Sub
'    End If
'    Dim id As Long
'    Dim nike As String
'    Dim uid As Long
'    Dim psw As String
'    db.Rs.MoveFirst
'    Do Until db.Rs.EOF
'        id = db.Rs.Fields("id")
'        nike = db.Rs.Fields("nike")
'        uid = db.Rs.Fields("uid")
'        psw = db.Rs.Fields("psw")
'        Text1.Text = Text1.Text & id & " " & nike & " " & uid & " " & psw & vbCrLf
'        db.Rs.MoveNext
'        DoEvents
'    Loop
'End Sub

'=============================================
'调用数据库过程 带参数示例
'=============================================
'Private Sub Kaijiang(bet_id As Long, bet_count As Long, bet_data As String)
'    '声明可以使用万能对象
'    'CreateObject ("ADODB.Parameter")
'    'CreateObject ("ADODB.Command")
'    Dim DBcommand As New ADODB.Command
'    Dim parm_id As New ADODB.Parameter
'    Dim parm_count As New ADODB.Parameter
'    Dim parm_date As New ADODB.Parameter
'    Dim parm_face As New ADODB.Parameter
'
'    Set parm_id = DBcommand.CreateParameter("@_betId", adInteger, adParamInput, 10, bet_id)
'    Set parm_count = DBcommand.CreateParameter("@_zjCount", adInteger, adParamInput, 10, bet_count)
'    Set parm_date = DBcommand.CreateParameter("@_kjData", adVarChar, adParamInput, 255, bet_data)
'    Set parm_face = DBcommand.CreateParameter("@_kset", adVarChar, adParamInput, 255, "0123456789")
'
'    DBcommand.Parameters.Refresh
'    DBcommand.Parameters.Append parm_id
'    DBcommand.Parameters.Append parm_count
'    DBcommand.Parameters.Append parm_date
'    DBcommand.Parameters.Append parm_face
'
'    DBcommand.ActiveConnection = db.Conn
'    DBcommand.CommandText = "kanJiang"
'    DBcommand.CommandType = adCmdStoredProc
'    DBcommand.Execute
'End Sub


'=============================================
'使用 Microsoft DataGrid Control 6.0(OLEDB)
'=============================================
'Set DataGrid1.DataSource = db.Rs           '绑定
'DataGrid1.Columns(0).Visible = False       '隐藏ID列
'Showlog DataGrid1.Columns(0).Caption       '第0列的列名

'Private Sub DataGrid1_Click()
    'DataGrid1.Col '被选中的单元格的列号
    'DataGrid1.Row '被选中的单元格的行号
    'Text1.Text = DataGrid1.Text
'End Sub

'DataGrid1.Columns(7).Width = 1600 '设置第7列宽度1700
'DataGrid1.Columns(0).Alignment = 1 '设置单元格左对齐
'DataGrid1.Columns(1).Alignment = 2 '设置单元格中心对齐
'DataGrid1.Columns(3).Alignment = 3 '设置单元格右对齐

'获取总行数的方法
'DataGrid1.ApproxCount
'DataGrid1.VisibleRows
'还是推荐使用 rs.recordcount


'DataGrid1.Scroll 0, -DataGrid1.FirstRow '拉动滚动条

'DataGrid1.Row = j
'DataGrid1.Col = i
'Text1.Text = DataGrid1.Text

'下列代码对"id"字段的那一列进行格式化显示,比如值为 "13" 将显示为 "P-0013"
'DataGrid1.Columns("id").NumberFormat = "P-0000"

'DataGrid1.VisibleCols   总列数

'Dim i As Long
'For i = 0 To DataGrid1.Columns.Count - 1
'    Showlog DataGrid1.Columns(i).Caption  '输出所有字段名
'Next i

'DataGrid1.AllowDelete = False          '不允许删除
'DataGrid1.AllowUpdate = False          '不允许更新
'DataGrid1.AllowAddNew = False          '不允许添加新行
'DataGrid1.Columns(0).Locked = True     '第一列只读设置
