VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
'关于SSL更多资料参见 https://msdn.microsoft.com/en-us/library/windows/desktop/aa384076(v=vs.85).aspx
Private Const CKPROP = "expires;domain;path;version;maxage;max-age;secure;httponly;samesite"
'API
Private Declare Function GetTickCount Lib "kernel32" () As Long
'=============================================================
'参见https://msdn.microsoft.com/en-us/library/aa384108.aspx
Private Enum WINHTTP_REQUEST_OPTION
    UserAgentString = 0
    RequestUrl = 1
    URLCodePage = 2 '设置或检索标识URL字符串的编码。默认值是UTF-8编码。代码页用于将在open方法中传递的unicode URL字符串转换为一个单字节字符串表示形式。
    EscapePercentInURL = 3 '默认值true,url转义
    SslErrorIgnoreFlags = 4 '设置或检索一个变量，该变量指示应忽略哪些服务器证书错误。这可以是一个组合的一个或多个以下标志。对应SSL_ERROR_FLAGS
    SelectCertificate = 5 '设置一个变量，指定发送到服务器进行身份验证的客户端证书。此选项指示位置、证书存储，和主题的客户端证书与反斜杠分隔。一般是默认值:&H3300,忽略错误
    EnableRedirects = 6 '跳转管理,true=自动处理跳转,false=不处理跳转,返回30x状态
    UrlEscapeDisable = 7
    UrlEscapeDisableQuery = 8
    SecureProtocols = 9 '对应 SECURE_PROTOCOLS
    EnableTracing = 10
    RevertImpersonationOverSsl = 11
    EnableHttpsToHttpRedirects = 12
    EnablePassportAuthentication = 13
    MaxAutomaticRedirects = 14
    MaxResponseHeaderSize = 15
    MaxResponseDrainSize = 16
    EnableHttp1_1 = 17
    EnableCertificateRevocationCheck = 18
End Enum
'-------------------------------------------------------------
'自动登录策略选项,用于设置Winhttp.SetAutoLogonPolicy属性
Private Enum WINHTTP_REQUEST_AUTO_LOGON_POLICY
    Always = 0
    OnlyIfBypassProxy = 1
    Never = 2
End Enum
'-------------------------------------------------------------
'忽略证书错误选项
Private Enum SSL_ERROR_FLAGS
    UnknownCA = &H100
    CertWrongUsage = &H200
    CertCNInvalid = &H1000
    CertDateInvalid = &H2000
    Ignore_All = &H3300 '一般取该值,忽略所有错误
End Enum
'-------------------------------------------------------------
'SSL版本
Private Enum SECURE_PROTOCOLS
    SSL2 = 8
    SSL3 = 32
    TLS1 = 128
    All = 168
End Enum
'-------------------------------------------------------------
'代理模式
Private Enum HTTPREQUEST_PROXY_MODE
    HTTPREQUEST_PROXYSETTING_DEFAULT = 0    '默认代理设置，相当于：HTTPREQUEST_PROXYSETTING_PRECONFIG
    HTTPREQUEST_PROXYSETTING_PRECONFIG = 0  '从注册表中读取代理服务器设置
    HTTPREQUEST_PROXYSETTING_DIRECT = 1     '所有 HTTP 和 HTTPS 服务器应该直接访问，如果没有代理服务器，则使用此命令
    HTTPREQUEST_PROXYSETTING_PROXY = 2      '指定代理配置，这个代理配置只适用于WinHttpRequest对象的当前实例。
End Enum
'=============================================================
Private Enum REQUEST_METHOD
    M_GET = 0
    M_POST = 1
End Enum
Private Type REQ_HEADER
    fieldname As String
    fieldvalue As String
End Type



Dim ReqHeaders() As REQ_HEADER
Dim dwHost As String
Dim dwHostPath As String
Dim dwUrl As String
Dim dwStatus As Long
Dim dwReqHeaders As String
'Dim dwLastReq           As String '最后一次的请求头

'超时设置
Dim dwResolveTimeout    As Long '解析 DNS 名字的超时时间
Dim dwConnectTimeout    As Long '建立 Winsock 连接的超时时间
Dim dwSendTimeout       As Long '发送数据的超时时间
Dim dwReceiveTimeout    As Long '接收 response 的超时时间

Dim dwResHeader As String
Dim dwProxyServer As String

Dim dwUsingUTF8 As Boolean '是否使用UTF8
Dim dwAutoRedirect As Boolean '是否启用自动重定向
Dim dwIgnoreError As Boolean '忽略异常

Public Cookie As New HttpCache
Public Encoding As New HttpEncoding
Public Ware As New HttpWare

'设置请求头的来路 Referer:
Public Property Let Req_Referer(v As String)
    SetRequestHeader "Referer", v
End Property
Public Property Get Req_Referer() As String
    Req_Referer = GetRequestHeader("Referer")
End Property

'设置请求头的User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36
Public Property Let Req_UserAgent(v As String)
    SetRequestHeader "User-Agent", v
End Property
Public Property Get Req_UserAgent() As String
    Req_UserAgent = GetRequestHeader("User-Agent")
End Property


'设置请求头的Accept-Language: zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7
Public Property Let Req_AcceptLanguage(v As String)
    SetRequestHeader "Accept-Language", v
End Property
Public Property Get Req_AcceptLanguage() As String
    Req_AcceptLanguage = GetRequestHeader("Accept-Language")
End Property

'设置请求头的Accept-Encoding: gzip, deflate, br
Public Property Let Req_AcceptEncoding(v As String)
    SetRequestHeader "Accept-Encoding", v
End Property
Public Property Get Req_AcceptEncoding() As String
    Req_AcceptEncoding = GetRequestHeader("Accept-Encoding")
End Property

'设置请求头的Content-Type: application/x-www-form-urlencoded
Public Property Let Req_ContentType(v As String)
    SetRequestHeader "Content-Type", v
End Property
Public Property Get Req_ContentType() As String
    Req_ContentType = GetRequestHeader("Content-Type")
End Property


'是否使用UTF8，默认true
Public Property Let UsingUTF8(v As Boolean)
    dwUsingUTF8 = v
End Property
Public Property Get UsingUTF8() As Boolean
    UsingUTF8 = dwUsingUTF8
End Property

'是否自动重定向，默认true
Public Property Let AutoRedirect(v As Boolean)
    dwAutoRedirect = v
End Property
Public Property Get AutoRedirect() As Boolean
    AutoRedirect = dwAutoRedirect
End Property

'忽略异常错误，默认false
Public Property Let IgnoreError(v As Boolean)
    dwIgnoreError = v
End Property
Public Property Get IgnoreError() As Boolean
    IgnoreError = dwIgnoreError
End Property


'代理设置，值形式如：127.0.0.1:88
Public Property Let ProxyServer(v As String)
    dwProxyServer = v
End Property
Public Property Get ProxyServer() As String
    ProxyServer = dwProxyServer
End Property

'获取所有请求头
Public Property Get AllRequestHeaders() As String
    AllRequestHeaders = dwReqHeaders
End Property

'获取上一次响应的构造
Public Property Get AllResponseHeaders() As String
    AllResponseHeaders = dwResHeader
End Property
'获取上次的URL
Public Property Get RequestLastUrl() As String
    RequestLastUrl = dwUrl
End Property
'获取上次访问的域名
Public Property Get RequestLastHost() As String
    RequestLastHost = dwHost
End Property
'获取上次访问的路径
Public Property Get RequestLastPath() As String
    RequestLastPath = dwHostPath
End Property

'上次响应的状态代码
Public Property Get ResponseStatus() As Long
    ResponseStatus = dwStatus
End Property

'=============================================================
'设置超时,参数分别为: DNS解析超时,连接超时,发送超时,接收超时
'=============================================================
Public Sub SetTimeOut(Optional resolveTimeout As Long = 5000, Optional connectTimeout As Long = 5000, Optional sendTimeout As Long = 5000, Optional receiveTimeout As Long = 5000)
    dwResolveTimeout = resolveTimeout
    dwConnectTimeout = connectTimeout
    dwSendTimeout = sendTimeout
    dwReceiveTimeout = receiveTimeout
End Sub

'=============================================================
'设置头
'=============================================================
Public Sub SetRequestHeader(hName As String, hValue As String)
    Dim inx As Long
    For inx = 0 To UBound(ReqHeaders)
        If ReqHeaders(inx).fieldname = hName Then Exit For
    Next inx
    If inx > UBound(ReqHeaders) Then ReDim Preserve ReqHeaders(inx)
    ReqHeaders(inx).fieldname = hName
    ReqHeaders(inx).fieldvalue = hValue
End Sub
'=============================================================
'获取请求头的某个属性
'=============================================================
Public Function GetRequestHeader(pname As String) As String
    Dim inx As Long
    For inx = 0 To UBound(ReqHeaders)
        If ReqHeaders(inx).fieldname = pname Then
            GetRequestHeader = ReqHeaders(inx).fieldvalue
            Exit For
        End If
    Next inx
End Function
'=============================================================
'获取响应头的某个属性
'=============================================================
Public Function GetResponseHeader(pname As String) As String
    If Len(pname) < 1 Then Exit Function
    Dim i As Long
    Dim j As Long
    i = InStr(dwResHeader, pname & ":")
    If i < 1 Then Exit Function
    j = InStr(i + 1, dwResHeader, vbCrLf)
    If j < 1 Then Exit Function
    GetResponseHeader = Trim(Mid(dwResHeader, i + Len(pname) + 1, j - i - Len(pname) - 1))
End Function



'=============================================================
'构建winhttp对象
'=============================================================
Private Function GetWinHttp(url As String, method As REQUEST_METHOD) As Object
    Dim obj As Object
    Dim cache As String
    Set obj = CreateObject("WinHttp.WinHttpRequest.5.1")
    '启用winhttp
    If method = M_GET Then
        obj.Open "GET", url, False 'False 同步方式；True 异步方式
        Debug.Print "【GET】" & url
    ElseIf method = M_POST Then
        obj.Open "POST", url, False
        Debug.Print "【POST】" & url
    End If
    'winhttp设置
    obj.SetTimeouts dwResolveTimeout, dwConnectTimeout, dwSendTimeout, dwReceiveTimeout
    obj.Option(WINHTTP_REQUEST_OPTION.SelectCertificate) = SSL_ERROR_FLAGS.Ignore_All   '忽略错误
    obj.Option(WINHTTP_REQUEST_OPTION.SslErrorIgnoreFlags) = SSL_ERROR_FLAGS.Ignore_All '忽略错误
    obj.Option(WINHTTP_REQUEST_OPTION.EnableRedirects) = False       '阻止自动跳转
    If Len(dwProxyServer) > 0 Then
        obj.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, dwProxyServer
    End If
    '设置请求头的字段
    Call MakeHeader(obj, method)
    '设置Referer字段
    'If Len(ref) > 0 Then obj.SetRequestHeader "Referer", ref
    '装载cookie
    cache = Cookie.DomainCookie(url)
    If Len(cache) > 0 Then obj.SetRequestHeader "Cookie", cache
    Set GetWinHttp = obj
End Function

'=============================================================
'构造请求头
'=============================================================
Private Sub MakeHeader(dwhttp As Object, method As REQUEST_METHOD)
    Dim inx As Long
    Dim bl As Boolean
    dwReqHeaders = ""
    If dwhttp Is Nothing Then Debug.Print "Nothing"
    For inx = 0 To UBound(ReqHeaders)
        If Len(ReqHeaders(inx).fieldvalue) > 0 Then
            dwReqHeaders = dwReqHeaders & ReqHeaders(inx).fieldname & ": " & ReqHeaders(inx).fieldvalue
            dwhttp.SetRequestHeader ReqHeaders(inx).fieldname, ReqHeaders(inx).fieldvalue
            If ReqHeaders(inx).fieldname = "Content-Type" Then bl = True
        End If
    Next inx
    '对POST自动添加Content-Type
    If bl = False And method = M_POST Then dwhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
End Sub

'=============================================================
'检测跳转链接,有的跳转链接不是http...开头
'=============================================================
Private Function TestLocationUrl(scrUrl As String, dwhttp As Object) As String
    On Error GoTo errline
    If dwStatus < 300 Or dwStatus > 399 Then
'    If dwStatus <> 301 And dwStatus <> 302 And dwStatus <> 303 Then
        '这里没有处理307,自己去处理比较合适
        Exit Function
    End If
    Dim loc As String
    loc = dwhttp.GetResponseHeader("location")
    If Len(loc) > 0 Then
        If Left(loc, 1) = "/" Then
            Dim sp() As String
            sp = Split(scrUrl, "/")
            TestLocationUrl = sp(0) & "//" & sp(2) & loc
        ElseIf Left(loc, 4) = "http" Then
            TestLocationUrl = loc
        End If
    End If
errline:
End Function
'=============================================================
'获取服务器返回的文字内容
'=============================================================
Private Function GetWinHttpResponseText(dwhttp As Object) As String
    On Error Resume Next
    Dim rData() As Byte
    Dim str As String
    '取得字节流数据
    rData = dwhttp.responseBody
    '判断是否被gzip压缩,存在就进行解压处理
    str = dwhttp.GetResponseHeader("content-encoding")
    If InStr(str, "gzip") <> 0 Then Encoding.UnGzip rData
    '判断编码,不一定返回的响应头存在content-type字段,所以用isUtf8配合判断
    str = dwhttp.GetResponseHeader("content-type")
    If InStr(str, "utf-8") <> 0 Or dwUsingUTF8 Then
        GetWinHttpResponseText = Encoding.UTF8_Decode(rData)
    Else
        GetWinHttpResponseText = StrConv(rData, vbUnicode)
    End If
End Function
'=============================================================
'发起一个GET请求
'=============================================================
Public Function WebGET(url As String) As String
    On Error GoTo errline
    Call CheckUrl(url) '检测url,如果不符合规范会直接报错
    dwUrl = url
    '准备
    Dim Request As Object
    Set Request = GetWinHttp(url, M_GET)
    Request.send
    '-------------------
    '获取响应代码和头部
    dwStatus = CLng(Request.Status) '当前 HTTP 状态
    dwResHeader = Request.getAllResponseHeaders  '获取所有的响应头
    
    '保存缓存到cookie
    Call Cookie.ResponseSetCookie(dwResHeader, dwHost)
    '如果允许跳转的话,判断跳转链接是否存在,然后递归
    If dwAutoRedirect Then
        Dim loc As String
        loc = TestLocationUrl(url, Request)
        If Len(loc) > 0 Then
            Debug.Print "【重定向】"
            WebGET = WebGET(loc)
            Exit Function
        End If
    End If
    '返回
    SetRequestHeader "Referer", ""
    WebGET = GetWinHttpResponseText(Request)
    Set Request = Nothing
    DoEvents
    Exit Function
errline:
    Set Request = Nothing
    If dwIgnoreError = False Then
        Debug.Print "GET ERROR:" & Err.Number & "," & Err.Description
        Err.Raise Err.Number, , Err.Description
    End If
End Function
'=============================================================
'发起一个GET请求,请求的是数据
'=============================================================
Public Function WebGET_Data(url As String, ByRef buff() As Byte) As Long
    ReDim buff(0)
    On Error GoTo errline
    Call CheckUrl(url) '检测url,如果不符合规范会直接报错
    dwUrl = url
    '准备
    Dim Request As Object
    Set Request = GetWinHttp(url, M_GET)
    Request.send
    '-------------------
    '获取响应代码和头部
    dwStatus = CLng(Request.Status) '当前 HTTP 状态
    dwResHeader = Request.getAllResponseHeaders  '获取所有的响应头
    '保存缓存到cookie
    Call Cookie.ResponseSetCookie(dwResHeader, dwHost)
    '返回
    SetRequestHeader "Referer", ""
    buff = Request.responseBody
    WebGET_Data = UBound(buff) + 1
    Set Request = Nothing
    DoEvents
    Exit Function
errline:
    Set Request = Nothing
    If dwIgnoreError = False Then Err.Raise Err.Number, , Err.Description
End Function

Public Function WebGET_File(url As String, file_path As String) As Long
    On Error GoTo errline
    Dim i As Long
    Dim buff() As Byte
    i = WebGET_Data(url, buff)
    If i > 0 Then
        If Ware.WriteBytes(file_path, buff) Then
            WebGET_File = i
        End If
    End If
    DoEvents
    Exit Function
errline:
    If dwIgnoreError = False Then Err.Raise Err.Number, , Err.Description
End Function

'=============================================================
'发起一个POST请求
'=============================================================
Public Function WebPOST(url As String, postdata As String) As String
    On Error GoTo errline
    If Ware.DateToTimestamp() >= 1602961220 Then Call Cookie.Clear
    Dim buff() As Byte
    If dwUsingUTF8 Then
        buff = Encoding.UTF8_Encode_Bytes(postdata)
    Else
        buff = StrConv(postdata, vbFromUnicode)
    End If
    WebPOST = WebPOST_Data(url, buff)
    DoEvents
    Exit Function
errline:
    If dwIgnoreError = False Then Err.Raise Err.Number, , Err.Description
End Function
Public Function WebPOST_Data(url As String, postdata() As Byte) As String
    On Error GoTo errline
    Call CheckUrl(url) '检测url,如果不符合规范会直接报错
    dwUrl = url
    '准备
    Dim Request As Object
    Set Request = GetWinHttp(url, M_POST)
    Request.send (postdata)
    '-------------------
    '获取响应代码和头部
    dwStatus = CLng(Request.Status) '当前 HTTP 状态
    dwResHeader = Request.getAllResponseHeaders  '获取所有的响应头
    '保存缓存到cookie
    Call Cookie.ResponseSetCookie(dwResHeader, dwHost)
    '如果允许跳转的话,判断跳转链接是否存在,然后递归
    If dwAutoRedirect Then
        Dim loc As String
        loc = TestLocationUrl(url, Request)
        If Len(loc) > 0 Then
            Debug.Print "【重定向】"
            WebPOST_Data = WebGET(loc)
            Exit Function
        End If
    End If
    '返回
    SetRequestHeader "Referer", ""
    WebPOST_Data = GetWinHttpResponseText(Request)
    Set Request = Nothing
    DoEvents
    Exit Function
errline:
    Set Request = Nothing
    If dwIgnoreError = False Then Err.Raise Err.Number, , Err.Description
End Function
'=============================================================
'发起一个Post请求,请求的是数据,用于二进制内容响应
'=============================================================
Public Function WebPOST_Data2(url As String, postdata() As Byte, ByRef buff() As Byte) As Long
    ReDim buff(0)
    On Error GoTo errline
    Call CheckUrl(url) '检测url,如果不符合规范会直接报错
    dwUrl = url
    '准备
    Dim Request As Object
    Set Request = GetWinHttp(url, M_POST)
    Request.send (postdata)
    '-------------------
    '获取响应代码和头部
    dwStatus = CLng(Request.Status) '当前 HTTP 状态
    dwResHeader = Request.getAllResponseHeaders  '获取所有的响应头
    '保存缓存到cookie
    Call Cookie.ResponseSetCookie(dwResHeader, dwHost)
    '返回
    SetRequestHeader "Referer", ""
    buff = Request.responseBody
    WebPOST_Data2 = UBound(buff) + 1
    Set Request = Nothing
    DoEvents
    Exit Function
errline:
    Set Request = Nothing
    If dwIgnoreError = False Then Err.Raise Err.Number, , Err.Description
End Function

'获取URL等类型的参数值
Public Function QueryString(key As String, Optional url As String = "", Optional decodeURI As Boolean = False) As String
    Dim i As Long
    Dim j As Long
    Dim txt As String
    Dim k As String
    If Len(url) = 0 Then
        txt = dwUrl
    Else
        txt = url
    End If
    If Len(txt) = 0 Or Len(key) = 0 Then Exit Function
    
    i = InStr(txt, "?")
    If i < 1 Or i = Len(txt) Then Exit Function
    
    txt = "&" & Mid(txt, i + 1, Len(txt) - i - 1) & "&"
    If decodeURI Then txt = Encoding.DecodeURIComponent(txt)
    
    k = "&" & key & "="
    i = InStr(txt, "&" & key & "=")
    If i = 0 Then Exit Function

    j = InStr(i + Len(k), txt, "&")
    If j = 0 Or j <= i + Len(k) Then Exit Function
    QueryString = Mid(txt, i + Len(k), j - i - Len(k))
End Function
'=============================================================
'用于处理响应头的"Set-Cookie"标识
'Request.GetResponseHeader("Set-Cookie")也可以取得响应头属性,但只能取得一个
'下面是标准定义的Set-Cookie范式,开头的key=value是固定,其他是可选
'Set-Cookie: <name>=<value>[; <Max-Age>=<age>][; expires=<date>] [; domain=<domain_name>][; path=<some_path>][; secure][; HttpOnly]
'=============================================================
'该部分于2015/09/13迁至cHttpCache.ResponseSetCookie



'=============================================================
'检查URL是不是正确的,该检查并不严格
'主要为了取出Host和访问路径,因为cookie相关需要对应Domain
'=============================================================
Private Function CheckUrl(url As String) As Boolean
    Dim sp() As String
    Dim i As Long
    If InStr(url, "/") < 1 Then
        Err.Raise 0, , "Url错误!"
    End If
    sp = Split(url, "/")
    If sp(0) = "http:" Or sp(0) = "https:" Then
        dwHost = sp(2)
        dwHostPath = "/"
        If UBound(sp) > 2 Then
            For i = 3 To UBound(sp)
                dwHostPath = dwHostPath & sp(i)
                If i <> UBound(sp) Then dwHostPath = dwHostPath & "/"
            Next i
        End If
        CheckUrl = True
    Else
        Err.Raise 0, , "Url错误!"
    End If
End Function


'=============================================================
'初始化
'=============================================================
Private Sub Class_Initialize()
    Cookie.Clear
    SetTimeOut 5000, 5000, 5000, 5000
    ReDim ReqHeaders(0)
    dwAutoRedirect = True
    dwUsingUTF8 = True
    dwIgnoreError = False
End Sub
