VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Private Type COOKIE_TYPE
    Cdomain As String
    Cpath As String
    Cname As String
    Cvalue As String
    Csecure As Boolean
End Type

Dim iCookie() As COOKIE_TYPE
    Private Const Q_DOMAIN = "domain="
    Private Const Q_PATH = "path="

'=============================================================
'字符串的提取 如:MidEx("is[12345]ok","[","]") 结果为"12345"
'=============================================================
Private Function MidEx(str As String, Lstr As String, Rstr As String, Optional starindex As Long = 1) As String
    Dim i As Long
    Dim j As Long
    If Len(Lstr) = 0 Then
        i = 1
    Else
        If starindex < 1 Then starindex = 1
        i = InStr(starindex, str, Lstr)
    End If
    If i = 0 Then Exit Function
    j = InStr(i + Len(Lstr), str, Rstr)
    If j = 0 Or j <= i + Len(Lstr) Then Exit Function
    MidEx = Mid(str, i + Len(Lstr), j - i - Len(Lstr))
    starindex = j
End Function

'=============================================================
'用于处理响应头的"Set-Cookie"标识
'Request.GetResponseHeader("Set-Cookie")也可以取得响应头属性,但只能取得一个
'下面是标准定义的Set-Cookie范式,开头的key=value是固定,其他是可选
'Set-Cookie: <name>=<value>[; <Max-Age>=<age>][; expires=<date>] [; domain=<domain_name>][; path=<some_path>][; secure][; HttpOnly]
'=============================================================
'2020/09/13
'修复由samesite引起的bug
'cookie只提取key,value,domain,path四项
'其他的expires/maxage/samesite... ...不再处理,再模拟提交中没有实际价值
Public Function ResponseSetCookie(resp_headers As String, host As String) As Long
    Dim offset As Long
    Dim strLine As String
    Dim i As Long, j As Long, l As Long, count As Long

    
    offset = 1
    strLine = MidEx(resp_headers, "Set-Cookie: ", vbCrLf, offset)
    If Len(strLine) = 0 Then Exit Function
    Do
        strLine = strLine & "; "
        i = InStr(strLine, "=")
        If i < 1 Then Exit Do
        j = InStr(strLine, "; ")
        If j < i Then Exit Do
        
        Dim dwkey As String
        Dim dwvalue As String
        Dim dwdomain As String
        Dim dwpath As String
        
        dwkey = Left(strLine, i - 1)
        l = j - i - 1
        If l > 0 Then
            dwvalue = Mid(strLine, i + 1, l)
        Else
            dwvalue = ""
        End If
        
        i = InStr(strLine, Q_DOMAIN)
        If i < 1 Then
            dwdomain = host
        Else
            i = i + Len(Q_DOMAIN)
            j = InStr(i, strLine, "; ")
            If j > i Then
                l = j - i
                dwdomain = Mid(strLine, i, l)
            Else
                dwdomain = host
            End If
        End If
        
        i = InStr(strLine, Q_PATH)
        If i < 1 Then
            dwpath = "/"
        Else
            i = i + Len(Q_PATH)
            j = InStr(i, strLine, "; ")
            If j > i Then
                l = j - i
                dwpath = Mid(strLine, i, l)
            Else
                dwpath = "/"
            End If
        End If
        
        AddItem dwkey, dwvalue, dwdomain, dwpath
        count = count + 1
        
        strLine = MidEx(resp_headers, "Set-Cookie: ", vbCrLf, offset)
        If Len(strLine) = 0 Then Exit Do
    Loop
    ResponseSetCookie = count
End Function

Public Function SaveCache(filepath As String, Optional isAppend As Boolean = False) As Long
    On Error GoTo errline
    Dim fn As Integer
    Dim i As Long
    Dim s As String
    Dim n As Long
    fn = FreeFile()
    If isAppend = False Then
        Open filepath For Output As #fn
    Else
        Open filepath For Append As #fn
    End If
    For i = 0 To UBound(iCookie)
        If Len(iCookie(i).Cname) <> 0 Then
            s = iCookie(i).Cname & ";" & iCookie(i).Cvalue & ";" & iCookie(i).Cpath & ";" & iCookie(i).Cdomain & ";" & iCookie(i).Csecure
            n = n + 1
            Print #fn, s
        End If
    Next i
    Close #fn
    SaveCache = n
    Exit Function
errline:
    SaveCache = -1
End Function

Public Function LoadCache(filepath As String) As Long
    On Error GoTo errline
    Dim fn As Integer
    Dim i As Long
    Dim s As String
    Dim sp() As String
    Dim bool As Boolean
    Dim n As Long
    fn = FreeFile()
    Open filepath For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, s
            If InStr(s, ";") <> 0 Then
                sp = Split(s, ";")
                If UBound(sp) = 4 Then
                    bool = True
                    For i = 0 To UBound(sp)
                        sp(i) = Trim(sp(i))
                        If i <> 1 And Len(sp(i)) = 0 Then bool = False
                    Next i
                    If bool Then
                        i = UBound(iCookie) + 1
                        ReDim Preserve iCookie(i)
                        iCookie(i).Cname = sp(0)
                        iCookie(i).Cvalue = sp(1)
                        iCookie(i).Cpath = sp(2)
                        iCookie(i).Cdomain = sp(3)
                        iCookie(i).Csecure = IIf(sp(4) = "True", True, False)
                        n = n + 1
                    End If
                End If
            End If
        Loop
    Close #fn
    LoadCache = n
    Exit Function
errline:
    LoadCache = -1
End Function
'获取与一个域名相关的COOKIE
Public Function DomainCookie(url As String) As String
Dim i As Long
Dim j As Long
Dim m As Long
Dim n As Long
Dim dom As String
Dim s As String
Dim sec As Boolean
Dim bool As Boolean
m = InStr(url, "//")
If m = 0 Then Exit Function
n = InStr(m + 2, url, "/")
If n = m + 2 Then Exit Function
If Left(url, 5) = "https" Then sec = True
If n = 0 Then
    dom = Right(url, Len(url) - m - 1)
Else
    dom = Mid(url, m + 2, n - m - 2)
End If

Dim samedomain As Boolean
For i = 1 To UBound(iCookie)
    samedomain = False
    If Len(iCookie(i).Cdomain) = Len(dom) Then
        If iCookie(i).Cdomain = dom Then samedomain = True
    ElseIf Len(iCookie(i).Cdomain) > Len(dom) Then
        j = InStrRev(iCookie(i).Cdomain, dom)
        If Len(iCookie(i).Cdomain) - j + 1 = Len(dom) Then samedomain = True
    Else
        j = InStrRev(dom, iCookie(i).Cdomain)
        If Len(dom) - j + 1 = Len(iCookie(i).Cdomain) Then samedomain = True
    End If
    If samedomain Then
        bool = False
'        If j + Len(iCookie(i).Cdomain) - 1 = Len(dom) Then
            If iCookie(i).Cpath = "/" Then
                If Len(iCookie(i).Cvalue) > 0 Then bool = True
            Else
                If n <> 0 Then
                    j = InStr(n, url, iCookie(i).Cpath)
                    If j = n Then
                        If Len(iCookie(i).Cvalue) > 0 Then bool = True
                    End If
                End If
            End If
'        End If
        If bool Then
            If iCookie(i).Csecure Then
                If sec = True Then s = s & "; " & iCookie(i).Cname & "=" & iCookie(i).Cvalue
            Else
                s = s & "; " & iCookie(i).Cname & "=" & iCookie(i).Cvalue
            End If
        End If
    End If
Next i
If Len(s) <> 0 Then DomainCookie = Right(s, Len(s) - 2)
End Function
'获取已知名字的COOKIE值,可选匹配参数为由域名或PATH值决定
Public Function GetValue(dwName As String, Optional dwdomain As String = "", Optional dwpath As String = "")
Dim i As Long
For i = 0 To UBound(iCookie)
    If iCookie(i).Cname = dwName Then
        If Len(dwdomain) > 0 Then
            If iCookie(i).Cdomain = dwdomain Then
                If Len(dwpath) > 0 Then
                    If iCookie(i).Cpath = dwpath Then GetValue = iCookie(i).Cvalue
                Else
                    GetValue = iCookie(i).Cvalue
                    Exit For
                End If
            End If
        Else
            If Len(dwpath) > 0 Then
                If iCookie(i).Cpath = dwpath Then GetValue = iCookie(i).Cvalue
            Else
                GetValue = iCookie(i).Cvalue
                Exit For
            End If
        End If
    End If
Next i
End Function

'添加一个COOKIE
Public Function AddItem(ckName As String, ckValue As String, ckDomain As String, Optional ckPath As String = "/", Optional ckSecure As Boolean = False) As Boolean
    Dim dwCookie As COOKIE_TYPE
    Dim i As Long
    Dim inx As Long
    dwCookie.Cname = Trim(ckName)
    dwCookie.Cvalue = Trim(ckValue)
    dwCookie.Cdomain = Trim(ckDomain)
    dwCookie.Cpath = Trim(ckPath)
    If Len(dwCookie.Cname) = 0 Then Exit Function
    If Len(dwCookie.Cdomain) = 0 Then Exit Function
    If Len(dwCookie.Cpath) = 0 Then dwCookie.Cpath = "/"
    If InStr(dwCookie.Cpath, "/") = 0 Then Exit Function
    inx = -1
    For i = 0 To UBound(iCookie)
        If iCookie(i).Cname = dwCookie.Cname Then
            If iCookie(i).Cpath = dwCookie.Cpath Then
                If iCookie(i).Cdomain = dwCookie.Cdomain Then
                    inx = i
                    Exit For
                End If
            End If
        End If
    Next i
    If inx = -1 Then
        inx = UBound(iCookie) + 1
        ReDim Preserve iCookie(inx)
    End If
    iCookie(inx) = dwCookie
    AddItem = True
End Function

'用于调试输出,查看所有的COOKIE
Public Function DebugCookie(Optional hasDomainAndPath As Boolean = True) As String
    Dim i As Long
    Dim s As String
    For i = 0 To UBound(iCookie)
        With iCookie(i)
            If Len(.Cname) > 0 Then
                If hasDomainAndPath Then s = s & .Cdomain & .Cpath & "; "
                s = s & .Cname & "=" & .Cvalue & vbCrLf
            End If
        End With
    Next i
    DebugCookie = s
End Function
'清除COOKIE
Public Sub Clear()
    ReDim iCookie(0)
End Sub

Private Sub Class_Initialize()
    Clear
End Sub

'--------------------------------------------------------------------
'COOKIE规则
'--------------------------------------------------------------------
'示例:
'请求的网站为:http://www.mts.com/translate?url=http://www.ibm.com/foo/index.html&language=French(注意其中包含两个URL,并且都访问了,在www.mts.com上将www.ibm.com/foo/index.html翻译成法文)
'收到的记录为:Set-Cookie：customer=huangxp; path=/foo; domain=.ibm.com; expires= Wednesday, 19-OCT-05 23:12:40 GMT; [secure]
'--------------------------------------------------------------------
'处理规则:
'浏览器将从Set-Cookie中解析出domain和path的值，分别是.ibm.com和/foo，
'并与请求的url：http://www.mts.com/translate?url=http://www.ibm.com/foo/index.html&language=French进行比较。
'请求url的domain是www.mts.com，path是/，与Set-Cookie中的属性不符，所以浏览器将忽略此Cookie
'--------------------------------------------------------------------
'Path
'--------------------------------------------------------------------
'path=/foo 控制哪些访问能够触发cookie 的发送。
'如果没有指定path，cookie 会在所有对此站点的HTTP 传送时发送。
'如果path=/directory，只有访问/directory 下面的网页时，cookie才被发送。
'在这个例子中，用户在访问目录/foo下的内容时，浏览器将发送此cookie。
'如果指定了path，但是path与当前访问的url不符，则此cookie将被忽略。  (重点)
'--------------------------------------------------------------------
'domain
'--------------------------------------------------------------------
'domain=.ibm.com 指定cookie被发送到哪台计算机上。
'正常情况下，cookie只被送回最初向用户发送cookie 的计算机。
'在这个例子中，cookie 会被发送到任何在.ibm.com域中的主机。
'如果domain 被设为空，domain 就被设置为和提供cookie 的Web 服务器相同。
'如果domain不为空，并且它的值又和提供cookie的Web服务器域名不符，这个Cookie将被忽略。
'--------------------------------------------------------------------
'EXPIRES(由于是即时访问,本类对于这个属性不做处理)
'--------------------------------------------------------------------
'expires= Wednesday, 19-OCT-05 23:12:40 GMT 指定cookie 失效的时间。
'如果没有指定失效时间，这个cookie 就不会被写入计算机的硬盘上，并且只持续到这次会话结束。
'--------------------------------------------------------------------
'secure
'--------------------------------------------------------------------
'secure 如果secure 这个词被作为Set-Cookie 头的一部分，那么cookie 只能通过安全通道传输（目前即SSL通道）。
'否则，浏览器将忽略此Cookie。
'--------------------------------------------------------------------
'总结
'--------------------------------------------------------------------
'基于上面Cookie规范的介绍和例证，我们能看出:
'浏览器在接受某一个站点的Cookie的时候，需要检查Cookie的参数domain、path、secure，
'看是否与当前的站点和URL相符，如果不符的话，就会忽略。
'另一方面。浏览器在上传Cookie的时候，也会根据当前所访问站点的属性，上传相关的Cookie，而其他的Cookie则不予上传。
'--------------------------------------------------------------------