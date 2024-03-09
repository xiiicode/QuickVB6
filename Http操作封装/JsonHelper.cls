VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Dim MSSC
Dim IsLoad_JSONstringify As Boolean

Private Sub Class_Terminate()
    Set MSSC = Nothing
End Sub

Public Function LoadJsonData(json_text As String) As Boolean
    On Error GoTo errline
    If IsObject(MSSC) Then
        MSSC.eval ("JDATA=" & json_text & ";")
    Else
        Set MSSC = CreateObject("MSScriptControl.ScriptControl")
        MSSC.AllowUI = False
        MSSC.Timeout = 3000
        MSSC.language = "JavaScript"
        MSSC.AddCode "Array.prototype.eq=function(n){return this[n];};var JDATA=" & json_text & ";" '
    End If
    LoadJsonData = True
    Exit Function
errline:
    Debug.Print "装载错误:" & Err.Number & "," & Err.Description
End Function

Public Function GetObj() As Object
    On Error GoTo errline
    Set GetObj = MSSC.CodeObject.JDATA
errline:
End Function

Public Function GetVal(subitem As String)
    On Error GoTo errline
    GetVal = MSSC.eval("JDATA." & subitem)
errline:
End Function

Public Function GetText(subitem As String) As String
    Dim v
    v = MSSC.eval("JDATA." & subitem & ".toString()")
    If Not IsObject(v) Then
        GetText = CStr(v)
    End If
End Function

Public Function GetInt(subitem As String) As Long
    Dim v
    v = GetVal(subitem)
    If IsNumeric(v) Then
        GetInt = CLng(v)
    End If
End Function

Public Function GetFloat(subitem As String) As Double
    Dim v
    v = GetVal(subitem)
    If IsNumeric(v) Then
        GetFloat = CDbl(v)
    End If
End Function

Public Function GetDate(subitem As String) As Date
    Dim v
    v = GetVal(subitem)
    If IsDate(v) Then
        GetDate = CDate(v)
    End If
End Function

Public Sub SetKV(key As String, v As String)
    Dim txt As String
    If Len(v) = 0 Then
        txt = "''"
    Else
        If IsNumeric(v) Then
            txt = v
        Else
            Dim c As String
            c = Mid(v, 1, 1)
            If c = "{" Or c = "[" Then
                txt = v
            Else
                txt = "'" & v & "'"
            End If
        End If
    End If
    MSSC.eval ("JDATA." & key & "=" & txt)
End Sub

Public Function ToString() As String
    On Error GoTo errline
    If IsLoad_JSONstringify = False Then
        IsLoad_JSONstringify = True
        '由于不支持JSON.stringify()方法,只能使用导入的函数
        '由 https://www.bugbugnow.net/2018/05/wshjscriptjson.html 修改
        MSSC.AddCode "var toString = Object.prototype.toString;"
        MSSC.AddCode "var hasOwnProperty = Object.prototype.hasOwnProperty;"
        MSSC.AddCode "var isArray = Array.isArray || function (a) { return toString.call(a) === '[object Array]'; };"
        MSSC.AddCode "var syh='" & Chr(34) & "';"
        MSSC.AddCode "var escMap = {syh: '\\'+syh, '\\': '\\\\', '\b': '\\b', '\f': '\\f', '\n': '\\n', '\r': '\\r', '\t': '\\t'};"
        MSSC.AddCode "var escFunc = function (m) { return escMap[m] || '\\u' + (m.charCodeAt(0) + 0x10000).toString(16).substr(1); };"
        MSSC.AddCode "var escRE = " & "/[\\" & Chr(34) & "\u0000-\u001F\u2028\u2029]/g;"
        MSSC.AddCode "function JSONstringify(value){if(value==null){return''}else if(typeof value==='number'){return isFinite(value)?value.toString():'null'}else if(typeof value==='boolean'){return value.toString()}else if(typeof value==='object'){if(typeof value.toJSON==='function'){return stringify(value.toJSON())}else if(isArray(value)){var res='[';for(var i=0;i<value.length;i++)res+=(i?', ':'')+JSONstringify(value[i]);return res+']'}else if(toString.call(value)==='[object Object]'){var tmp=[];for(var k in value){if(hasOwnProperty.call(value,k))tmp.push(JSONstringify(k)+': '+JSONstringify(value[k]))}return'{'+tmp.join(', ')+'}'}}return syh+value.toString().replace(eval(escRE),escFunc)+syh};"
    End If
    ToString = MSSC.eval("JSONstringify(JDATA)")
errline:
End Function

'*--------------------------------------------------
'|样本
'*--------------------------------------------------
'{
'    ts: '2000-10-1 8:1',
'    bb: 'jack',
'    items: ['aa', 'bb', 'cc'],
'    objs: [{
'        id: 1,
'        tag: 'xx'
'    },
'    {
'        id: 2,
'        tag: 'yy'
'    }]
'}
'*--------------------------------------------------
'|示例
'*--------------------------------------------------
'Private Sub Command1_Click()
'
'    Dim JH As New JsonHelper
'    Dim json As String
'
'    json = "{ts:'2000-10-1 8:1' ,name: 'jack', items: ['aa', 'bb', 'cc'], objs:[{id:1, info:'xx'},{id:2, info:'yy'}]}"
'
'    If JH.LoadJsonData(json) = False Then
'        Debug.Print "装载JSON失败"
'        Exit Sub
'    End If
'    Debug.Print "-----普通读取"
'    Debug.Print "ts = " & JH.GetDate("ts")
'    Debug.Print "name = " & JH.GetText("name")
'    Dim count As Long, i As Long
'    count = JH.GetVal("items.length")
'    Debug.Print "items总数:" & count
'    For i = 0 To count - 1
'        Debug.Print "items" & i & " = " & JH.GetText("items[" & i & "]")
'    Next i
'
'    Debug.Print "-----对象读取"
'    Dim o
'    Set o = JH.GetObj()
'    Debug.Print "objs1.items3 = " & o.items.eq(2)
'    Debug.Print "objs1.objs2.info = " & o.objs.eq(1).info  '注意这里的eq取下标的方法,通常不建议使用对象方式读取,由于json区分大小写,而VB会对某些名称自动大写,如果name,tag...
'
'    Debug.Print "-----编辑写入"
'    JH.SetKV "city", "beijing"
'    JH.SetKV "items[3]", "dd"
'
'    Debug.Print "-----输出JSON字符串"
'    Debug.Print JH.ToString()
'
'End Sub

'*--------------------------------------------------
'|输出
'*--------------------------------------------------
'-----普通读取
'ts = 2000/10/1 8:01:00
'Name = jack
'items总数:3
'items0 = aa
'items1 = bb
'items2 = cc
'-----对象读取
'objs1.items3 = cc
'objs1.objs2.info = yy
'-----编辑写入
'-----输出JSON字符串
'{"ts": "2000-10-1 8:1", "name": "jack", "items": ["aa", "bb", "cc", "dd"], "objs": [{"id": 1, "info": "xx"}, {"id": 2, "info": "yy"}], "city": "beijing"}
