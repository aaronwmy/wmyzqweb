VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "��ҳץȡ����(����ר��)"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   15720
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "������Դ"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ַ"
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
      Caption         =   "��ַ��"
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
    
     '���data�ļ��в����ڣ��򴴽�data�ļ���,Public�ļ��м������ļ���,Tpl�ļ���
    If fso.FolderExists(App.Path & "\data") = False Then
        fso.CreateFolder (App.Path & "\data")
        fso.CreateFolder (App.Path & "\data\Public")
        fso.CreateFolder (App.Path & "\data\Tpl")
        '����΢���ļ����е��ļ���
        fso.CreateFolder App.Path & "/data/Public/images"
        fso.CreateFolder App.Path & "/data/Public/css"
        fso.CreateFolder App.Path & "/data/Public/js"
    End If
    
    '���û�ȷ���Ƿ���վ��ļ�
    If MsgBox("��Ҫ��վ��ļ���", vbYesNo) = vbYes Then
        'ɾ��Public�ļ����е������ļ����ļ���
        Set datafolder = fso.GetFolder(App.Path & "\data\Public")
        For Each File In datafolder.Files
            fso.DeleteFile File.Path
        Next
        For Each Folder In datafolder.SubFolders
            fso.DeleteFolder Folder.Path
        Next
    
        '����Public�ļ����е��ļ���
        fso.CreateFolder datafolder.Path & "/images"
        fso.CreateFolder datafolder.Path & "/css"
        fso.CreateFolder datafolder.Path & "/js"
        
        'ɾ��Tpl�ļ����е������ļ�
        Set datafolder = fso.GetFolder(App.Path & "\data\Tpl")
        For Each File In datafolder.Files
            fso.DeleteFile File.Path
        Next
    End If
    
    '�����ҳhtml�ı���url
    weburl = WebBrowser1.LocationURL
    'HTML = WebBrowser1.Document.DocumentElement.outerhtml
    HTML = gethtml(weburl)
    
    
     '�����������
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    
    '��õ�ǰ��ҳ����
    regex.Pattern = "(https?:\/\/.+?)\/"
    Set matchs = regex.Execute(weburl)
    domain = matchs(0).submatches(0)
    
    '��õ�ǰ��ҳ·��
    regex.Pattern = "https?:\/\/.+\/"
    Set matchs = regex.Execute(weburl)
    webpath = matchs(0)
    
    '������ҳ�е�js,css��ͼƬ
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
    
    '����css�ļ��е�ͼƬ
    For i = 0 To UBound(cssarr) - 2
        cssurl = cssarr(i)
        '�������css�ļ���������
        regex.Pattern = "(https?:\/\/.+?)\/"
        Set matchs = regex.Execute(cssurl)
        cssdomain = matchs(0).submatches(0)
        '�������css�ļ�����·��
        regex.Pattern = "https?:\/\/.+\/"
        Set matchs = regex.Execute(cssurl)
        csswebpath = matchs(0)
        '���css�ļ���
        regex.Pattern = "[^/\\]+?\.css"
        Set matchs = regex.Execute(cssurl)
        cssfilename = matchs(0)
        
        '���css�ļ�����
        Set F = fso.OpenTextFile(App.Path & "/data/Public/css/" & cssfilename, 1)
        csstext = F.ReadAll
        
        '����css�ļ��е�ͼƬ
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
    
    '�޸�HTML�ı�
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
    
    '������ҳ
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
    
    MsgBox "�ɹ���"
End Sub

'�����ļ�
Private Function downloadfile(url)
    On Error Resume Next
    
    '����ļ���׺
    filetype = LCase(Right(url, Len(url) - InStrRev(url, ".")))
    
    '����ļ�����·��
    If filetype = "jpg" Or filetype = "png" Or filetype = "gif" Then
        filepath = App.Path & "/data/Public/images"
    ElseIf filetype = "css" Then
        filepath = App.Path & "/data/Public/css"
    ElseIf filetype = "js" Then
        filepath = App.Path & "/data/Public/js"
    End If
    
    '����ļ����ص�ַ
    If InStrRev(url, "/") > InStrRev(url, "\") Then
        filepath = filepath + "\" + Right(url, Len(url) - InStrRev(url, "/"))
    Else
        filepath = filepath + "\" + Right(url, Len(url) - InStrRev(url, "\"))
    End If
    
    '���ļ���������
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
