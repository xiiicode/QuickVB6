VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HtmlDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------
'https://github.dev/xiiicode/QuickVB6
'------------------------------
Public Document As Object
Dim isLoad As Boolean
Public Function LoadHtml(html As String) As Long
    isLoad = False
    Set Document = Nothing
    Set Document = CreateObject("HTMLFILE")
    Document.designMode = "on"
    Document.Write html
    Document.Close
    Dim i As Long
    i = Document.All.length
    If i > 0 Then
        isLoad = True
        LoadHtml = i
    End If
End Function

Private Function get_Img_src(Optional class_name As String = "") As String()
    On Error Resume Next
    Dim s() As String
    Dim n As Long
    Dim e
    ReDim s(0)
    For Each e In Document.getElementsByTagName("img")
        If Len(class_name) > 0 Then
            If e.className = class_name Then
                s(n) = e.src
                n = n + 1
            End If
        Else
            ReDim Preserve s(n)
            s(n) = e.src
            n = n + 1
        End If
        DoEvents
    Next
    get_Img_src = s
End Function

Private Function get_a_src(Optional class_name As String = "") As String()
    On Error Resume Next
    Dim s() As String
    Dim n As Long
    Dim e
    ReDim s(0)
    For Each e In Document.getElementsByTagName("a")
        If Len(class_name) > 0 Then
            If e.className = class_name Then
                s(n) = e.href
                n = n + 1
            End If
        Else
            ReDim Preserve s(n)
            s(n) = e.href
            n = n + 1
        End If
        DoEvents
    Next
    get_a_src = s
'    Dim s() As String
'    Dim n As Long
'    ReDim s(0)
'    For Each a In doc.getElementsByTagName("a")
'        If a.className = "someone" Then
'            s = s & a.innerText & "--" & a.href & vbCrLf
'        End If
'    Next
End Function
'For Each s In html.Document.getElementsByTagName("a")
'    If s.classname = "postTitle2" Then
'        Debug.Print s.innertext & vbCrLf & s.href & vbCrLf
'    End If
'Next

'For Each s In html.Document.links
'    Debug.Print CStr(s)
'Next


'Dim s
'Dim p
'Set s = html.Document.getElementByid("sina_keyword_ad_area2")
'Dim r As String
'For Each p In s.childNodes
'    r = r & p.innertext
'Next
'Debug.Print r

'//.getAttribute("name") 获取元素的name属性
'    For Each span In html.Document.getElementsByTagName("span")
'        If span.getAttribute("itemprop") = "ratingValue" Then
'            fs = span.innertext
'        ElseIf span.getAttribute("itemprop") = "ratingCount" Then
'            ct = span.innertext
'        End If
'    Next


'关于多层获取举例,比如要获取图片地址:
'<div id="mainpic" class="">
'    <a class="nbgnbg" href="http://movie.douban.com/subject/25778483/photos?type=R" title="点击看更多海报">
'        <img src="http://img4.douban.com/view/movie_poster_cover/spst/public/p2284651956.jpg" title="点击看更多海报" alt="消失的凶手" rel="v:image" />
'   </a>
'</div>
'-------------------------
'Dim em
'Set em = html.Document.getElementById("mainpic").getElementsByTagName("a")(0).getElementsByTagName("img")(0)
'Debug.Print em.src