VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "网页抓取程序(美工专用)"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   15720
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "下载资源"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开网址"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5400
      Width           =   14415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      ExtentX         =   27331
      ExtentY         =   9128
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "网址："
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Command2_Click()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
     '如果data文件夹不存在，则创建data文件夹,Public文件夹及其子文件夹,Tpl文件夹
    If fso.FolderExists(App.Path & "\data") = False Then
        fso.CreateFolder (App.Path & "\data")
        fso.CreateFolder (App.Path & "\data\Public")
        fso.CreateFolder (App.Path & "\data\Tpl")
        '创建微端文件夹中的文件夹
        fso.CreateFolder App.Path & "/data/Public/images"
        fso.CreateFolder App.Path & "/data/Public/css"
        fso.CreateFolder App.Path & "/data/Public/js"
    End If
    
    '让用户确认是否清空旧文件
    If MsgBox("需要清空旧文件吗？", vbYesNo) = vbYes Then
        '删除Public文件夹中的所有文件和文件夹
        Set datafolder = fso.GetFolder(App.Path & "\data\Public")
        For Each File In datafolder.Files
            fso.DeleteFile File.Path
        Next
        For Each Folder In datafolder.SubFolders
            fso.DeleteFolder Folder.Path
        Next
    
        '创建Public文件夹中的文件夹
        fso.CreateFolder datafolder.Path & "/images"
        fso.CreateFolder datafolder.Path & "/css"
        fso.CreateFolder datafolder.Path & "/js"
        
        '删除Tpl文件夹中的所有文件
        Set datafolder = fso.GetFolder(App.Path & "\data\Tpl")
        For Each File In datafolder.Files
            fso.DeleteFile File.Path
        Next
    End If
    
    '获得网页html文本和url
    weburl = WebBrowser1.LocationURL
    'HTML = WebBrowser1.Document.DocumentElement.outerhtml
    HTML = gethtml(weburl)
    
    
     '正则变量设置
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    
    '获得当前网页域名
    regex.Pattern = "(https?:\/\/.+?)\/"
    Set matchs = regex.Execute(weburl)
    domain = matchs(0).submatches(0)
    
    '获得当前网页路径
    regex.Pattern = "https?:\/\/.+\/"
    Set matchs = regex.Execute(weburl)
    webpath = matchs(0)
    
    '下载网页中的js,css和图片
    regex.Pattern = "[^""'\(\)\[\]\\]+?\.(gif|png|jpg|jpeg|js|css)"
    Dim cssarr() As String
    ReDim cssarr(1)
    Set matchs = regex.Execute(HTML)
    For i = 0 To matchs.Count - 1
        filepath = matchs(i)
        filepath = Replace(filepath, "\/", "/")
        If Left(filepath, 1) = "/" Then
            If Left(filepath, 2) = "//" Then
                If Left(weburl, 5) = "https" Then
                    filepath = "https:" & filepath
                Else
                    filepath = "http:" & filepath
                End If
            Else
                filepath = domain & filepath
            End If
        ElseIf Left(filepath, 7) = "http://" Or Left(filepath, 8) = "https://" Then
            
        Else
            filepath = webpath & filepath
        End If
        
        downloadfile (filepath)
        If Right(filepath, 4) = ".css" Then
            cssarr(UBound(cssarr) - 1) = filepath
            ReDim Preserve cssarr(UBound(cssarr) + 1)
        End If
    Next
    
    '下载css文件中的图片
    For i = 0 To UBound(cssarr) - 2
        cssurl = cssarr(i)
        '获得网络css文件所在域名
        regex.Pattern = "(https?:\/\/.+?)\/"
        Set matchs = regex.Execute(cssurl)
        cssdomain = matchs(0).submatches(0)
        '获得网络css文件所在路径
        regex.Pattern = "https?:\/\/.+\/"
        Set matchs = regex.Execute(cssurl)
        csswebpath = matchs(0)
        '获得css文件名
        regex.Pattern = "[^/\\]+?\.css"
        Set matchs = regex.Execute(cssurl)
        cssfilename = matchs(0)
        
        '获得css文件内容
        Set F = fso.OpenTextFile(App.Path & "/data/Public/css/" & cssfilename, 1)
        csstext = F.ReadAll
        
        '下载css文件中的图片
        regex.Pattern = "[^""=\(\)]+?\.(gif|png|jpg|jpeg)"
        Set matchs = regex.Execute(csstext)
        For j = 0 To matchs.Count - 1
            filepath = matchs(j)
            If Left(filepath, 1) = "/" Then
                filepath = cssdomain & filepath
            ElseIf Left(filepath, 7) = "http://" Then
                
            Else
                filepath = csswebpath & filepath
            End If
            downloadfile (filepath)
        Next
    Next
    
    '修改HTML文本
    regex.Pattern = "[^""'\(\)\[\]\\]+?\.(gif|png|jpg|jpeg|js|css)"
    Set matchs = regex.Execute(HTML)
    regex.Pattern = "[^/\\]*\.(gif|png|jpg|jpeg|js|css)"
    For i = 0 To matchs.Count - 1
        filepath = matchs(i)
        Set submatchs = regex.Execute(matchs(i))
        If submatchs(0).submatches(0) = "png" Or submatchs(0).submatches(0) = "gif" Or submatchs(0).submatches(0) = "jpg" Or submatchs(0).submatches(0) = "jpeg" Then
            HTML = Replace(HTML, filepath, "../Public/images/" & submatchs(0))
        ElseIf submatchs(0).submatches(0) = "css" Then
            HTML = Replace(HTML, filepath, "../Public/css/" & submatchs(0))
        ElseIf submatchs(0).submatches(0) = "js" Then
            HTML = Replace(HTML, filepath, "../Public/js/" & submatchs(0))
        End If
    Next
    
    '保存网页
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
    .Open
    .Charset = "utf-8"
    .Position = objStream.Size
    .WriteText = HTML
    .SaveToFile App.Path & "/data/Tpl/" & DateDiff("s", "1970-01-01 00:00:00", Now) & ".html", 2
    .Close
    End With
    Set objStream = Nothing
    
    MsgBox "成功了"
End Sub

'下载文件
Private Function downloadfile(url)
    On Error Resume Next
    
    '获得文件后缀
    filetype = LCase(Right(url, Len(url) - InStrRev(url, ".")))
    
    '获得文件下载路径
    If filetype = "jpg" Or filetype = "png" Or filetype = "gif" Then
        filepath = App.Path & "/data/Public/images"
    ElseIf filetype = "css" Then
        filepath = App.Path & "/data/Public/css"
    ElseIf filetype = "js" Then
        filepath = App.Path & "/data/Public/js"
    End If
    
    '获得文件下载地址
    If InStrRev(url, "/") > InStrRev(url, "\") Then
        filepath = filepath + "\" + Right(url, Len(url) - InStrRev(url, "/"))
    Else
        filepath = filepath + "\" + Right(url, Len(url) - InStrRev(url, "\"))
    End If
    
    '对文件进行下载
    Set h = CreateObject("MSXML2.XMLHTTP")
    h.Open "GET", url, False
    h.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.82 Safari/537.36"
    h.send
    While h.ReadyState <> 4
        h.waitForResponse 1000
    Wend
    If Len(h.Responsebody) > 0 Then
        Set s = CreateObject("ADODB.Stream")
        s.Type = 1
        s.Open
        s.Write h.Responsebody
        s.SaveToFile filepath, 2
        s.Close
    End If
    Set h = Nothing
    
    On Error GoTo 0
End Function

Function gethtml(url)
    Set h = CreateObject("MSXML2.XMLHTTP")
    h.Open "GET", url, False
    h.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.82 Safari/537.36"
    h.send
    gethtml = h.Responsetext
End Function
