VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "版权所有 XIAOKONGS 2017"
   ClientHeight    =   4005
   ClientLeft      =   7155
   ClientTop       =   5820
   ClientWidth     =   3945
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3945
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   0
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   22
      Top             =   0
      Width           =   3975
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "疯狂加载中。。。"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2400
         TabIndex        =   23
         Top             =   2280
         Width           =   1440
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   5775
      Left            =   8640
      TabIndex        =   21
      Top             =   960
      Width           =   5415
      ExtentX         =   9551
      ExtentY         =   10186
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
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "呼叫"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "挂机"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         ItemData        =   "Form1.frx":68A6
         Left            =   120
         List            =   "Form1.frx":68B0
         TabIndex        =   9
         Text            =   ">>请选择您要拨打的号码"
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "紧急呼叫"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   735
         Width           =   720
      End
   End
   Begin VB.CommandButton cm4 
      Caption         =   "蔡万禹"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Ext 
      Caption         =   "确认退出"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cm2 
      Caption         =   "蔡万禹"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   4680
      TabIndex        =   2
      Top             =   8160
      Width           =   7575
   End
   Begin VB.CommandButton cm1 
      Caption         =   "XIAOKONGS"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7575
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      ExtentX         =   5741
      ExtentY         =   13361
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
   Begin VB.CommandButton cm3 
      Caption         =   "蔡万禹"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   480
      TabIndex        =   16
      Top             =   360
      Width           =   3015
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上海游弛网络技术有限公司"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   390
         TabIndex        =   18
         Top             =   600
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "感谢使用"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "⊙ 恭喜您 转移成功啦 ~"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   24
      Top             =   840
      Width           =   3045
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始进行参数设置 "
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1200
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label10 
      Caption         =   "紧急呼叫"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "参数设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "常用联系人"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "呼叫记录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   3600
      TabIndex        =   7
      Top             =   2040
      Width           =   210
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents m_Doc As MSHTML.HTMLDocument
Attribute m_Doc.VB_VarHelpID = -1
Dim Location As Long
Dim Address As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cm2_Click()

On Error Resume Next

Call CallSB(cm2.Tag)
DoEvents
Me.Caption = "·> 已激活 " & WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value

Label1.Visible = False
Label2.Visible = False
cm1.Visible = False
cm2.Visible = False
cm3.Visible = False
cm4.Visible = False

Label3.Visible = True
Ext.Visible = True

Frame2.Visible = False

End Sub

Private Sub cm1_Click()

On Error Resume Next

'呼叫转移设置到SB号码
Call CallSB(cm1.Tag)
DoEvents
Me.Caption = "·> 已激活 " & WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value

Label1.Visible = False
Label2.Visible = False
cm1.Visible = False
cm2.Visible = False
cm3.Visible = False
cm4.Visible = False

Label3.Visible = True
Ext.Visible = True

Frame2.Visible = False

End Sub

Private Sub cm3_Click()

On Error Resume Next

'Me.Caption = "·> 已激活 " & i

If Left(Me.Caption, 1) = "·" Then


Call CallSB(cm3.Tag)
DoEvents
Me.Caption = "·> 已激活 " & WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value

Label1.Visible = False
Label2.Visible = False
cm1.Visible = False
cm2.Visible = False
cm3.Visible = False
cm4.Visible = False

Label3.Visible = True
Ext.Visible = True
Frame2.Visible = False

    Else
    
'    MsgBox " 正在读取VOIP数据请稍后 ！！！", vbOKOnly, "XIAOKONGS實驗室"

End If

End Sub

Private Sub cm4_Click()

On Error Resume Next

Call CallSB(cm4.Tag)
DoEvents
Me.Caption = "·> 已激活 " & WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value

Label1.Visible = False
Label2.Visible = False
cm1.Visible = False
cm2.Visible = False
cm3.Visible = False
cm4.Visible = False

Label3.Visible = True
Ext.Visible = True
Frame2.Visible = False
End Sub

Private Sub Combo1_Change()
'Me.WebBrowser1.Navigate2 "http://" & e & "/webdial.htm"      '"http://10.21.22.15/webdial.htm"
'If Len(Combo1.Text) < 4 Then
'Label5.Caption = ">>内部电话"
'Else
'End If
End Sub

Private Sub Combo1_Click()

On Error Resume Next

   Dim StrArray() As String
   Dim ret As Long
   Dim logo As String
   Dim logo1 As String
   
   StrArray = Split(Combo1.Text, " ")
   
   ret = UBound(StrArray) + 1
   

    
    DelChina_Find_SH
    
    DoEvents
    
    If Len(Trim(StrArray(0))) = 11 Then
        logo1 = "手机"
    Else
    logo1 = "座机"
    
    End If
    

'   If Combo1.Tag <> "90" Then
'      logo = " 上海"
'        Else
'      logo = ""
'   End If   'Dim Address As String
    Address = " " & Trim(Address)
   
    If Len(Trim(StrArray(0))) = 4 Then logo1 = "内部号码": Combo1.Tag = "": Address = ""  '内部号码不需要加90
   
    If ret > 0 Then
     Label5.Caption = ">>" & Trim(StrArray(1)) & "" & Address & logo1
    End If
    
End Sub

Private Sub Command1_Click()

On Error Resume Next

WebBrowser1.Navigate "http://" & E & "/webdial.htm"
DoEvents
Call StopCaller

End Sub

Private Sub Command2_Click()

On Error Resume Next

' WebBrowser1.Navigate "www.baidu.com"
'    Do While WebBrowser1.Busy '等待加载完成.
'    DoEvents
'    Loop
'    MsgBox "加载完成!", vbOKOnly, "!"
'    End Sub

'验证手机号码
If Combo1.Tag = "" Then
'Label5.Caption = ">>计算号码归属地"
DelChina_Find_SH
DoEvents
End If

If Len(Combo1.Text) = 4 Then Combo1.Tag = ""  '内部号码不需要加90

WebBrowser1.Navigate "http://" & E & "/webdial.htm"

'
Do While WebBrowser1.Busy '等待加载完成.
    DoEvents
    Label5.Caption = ">>拨号中。。。"
    Command2.Enabled = False
Loop

MsgBox "VOIP拨号已完成 ！！！", vbOKOnly, "XIAOKONGS實驗室"

Label5.Caption = "紧急呼叫"

WebBrowser1.Document.All("WEB_DialNumber").value = Combo1.Tag & Combo1.Text
'WebBrowser1.Document.getElementById("WEB_DialNumber").Value =
Calling Combo1.Tag & Combo1.Text


End Sub

Private Sub Ext_Click()
End
End Sub
Private Sub Form_Activate()
NameInit
str = "http://" & E & "/sipset.htm"
Load Caller
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Label1_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Label12_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Label4_Click()

On Error Resume Next

If Height = BigggHeight Then
Height = SmallHeight
Label4.Left = 3600
Label4.Caption = "+"
Else
Height = BigggHeight
Label4.Caption = "⊙"
Label4.Left = 3570
End If

End Sub

Private Sub Label5_Click()

On Error Resume Next

If Label5.Caption = "紧急呼叫" Then
Combo1.Text = Urgent
Call USA
End If

End Sub

Private Sub Label6_Click()
On Error Resume Next
Shell App.Path & "\" & "VOIP 呼出记录.exe"
End Sub

Private Sub Label7_Click()
Me.Hide
Caller.Show
End Sub

Private Sub Label8_Click()
Me.Hide
Form3.Show
End Sub
Private Sub Label9_Click()
On Error Resume Next

If Label9.Caption = "拨打电话" Then
Label4_Click
Exit Sub
End If
Me.Hide
Form3.Show

End Sub
Private Function m_Doc_onclick() As Boolean

    Dim elem As IHTMLElement
    
    Set elem = m_Doc.parentWindow.event.srcElement
    Debug.Print "m_Doc_onclick", "当前触发事件的元素：", elem.tagName, elem.sourceIndex, elem.id
    m_Doc_onclick = True
    
End Function


Private Sub Picture1_Click()
Picture1.Visible = False
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Set m_Doc = Me.WebBrowser1.Document
End Sub

'自动登陆
Private Sub StartLogonA()

On Error Resume Next

'Debug.Print "开始尝试自动登陆……"

WebBrowser1.Document.All("username").value = "admin"

WebBrowser1.Document.All("password").value = "admin"

WebBrowser1.Document.All("goto").Click


End Sub

Public Sub CallSB(Number As String)

WebBrowser1.Document.All("SIP_EnableAlways_RW").Checked = True
WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value = Number
CCC
End Sub

Public Sub CallCWY()
WebBrowser1.Document.All("SIP_EnableAlways_RW").Checked = True
WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value = "9017315368616"
CCC
End Sub

Private Sub Form_Load()

If App.PrevInstance Then End

'MsgBox OS

Dim rtn
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)


SmallHeight = 3000
BigggHeight = 4450

Height = SmallHeight

'If OS <> "Win7" Then Label4_Click

Call OS

NameInit
str = "http://" & E & "/sipset.htm"
WebBrowser1.Navigate (str)   '连接VOIP



End Sub

Private Sub Form_Resize()

'WebBrowser1.Top = Me.Top - 200
'WebBrowser1.Left = Me.Left - 200
'WebBrowser1.Width = Me.Width - 200
'WebBrowser1.Height = Me.Height

End Sub

'Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'
'Dim szPost As String
'szPost = StrConv(PostData, vbUnicode)
'
'Debug.Print szPost  '输出POST数据
'
'End Sub

Public Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

On Error Resume Next

Dim i As String

If Not WebBrowser1.Busy Then '判断网页是否下载完毕

    Dim strTitle As String
    strTitle = WebBrowser1.Document.documentElement.innerHTML
    strTitle = Mid(strTitle, 1, InStr(1, LCase(strTitle), "</title>") - 1)
    strTitle = Mid(strTitle, InStr(1, LCase(strTitle), "<title>") + 7)

'    Me.Caption = strTitle

        If strTitle = "Logon" Then
        
           StartLogonA '网页标题=Logon？立即执行自动登陆
           
        End If
    
   i = WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value
   
   If i <> "" Then
   
   Me.Caption = "·> 已激活 " & i
   Sleep 666
   Picture1.Visible = False
   
   Else
   
'      Me.Caption = ">>呼叫转移业务已停止"
   
   End If
   
'
'   If pDisp = WebBrowser1.Object Then
'    MsgBox "网页完全加载完了！！！"
'   End If

End If


End Sub

'屏蔽网页错误
Private Sub WebBrowser1_DownloadBegin()
WebBrowser1.Silent = True
End Sub
Private Sub WebBrowser1_DownloadComplete()
WebBrowser1.Silent = True
End Sub
Private Sub CCC()

On Error Resume Next

Dim vDoc, vTag
Dim i As Integer
Set vDoc = WebBrowser1.Document
List1.Clear
For i = 0 To vDoc.All.length - 1
If UCase(vDoc.All(i).tagName) = "INPUT" Then
Set vTag = vDoc.All(i)
If vTag.Type = "submit" Then
List1.AddItem vTag.Name
Select Case vTag.Name
Case "DefaultSubmit"
vTag.Click
Exit Sub
End Select
End If
End If
Next i

End Sub
Public Function CleanNum()

On Error Resume Next

Me.Caption = "·呼叫转移已解除"
WebBrowser1.Document.All("SIP_EnableAlways_RW").Checked = False
WebBrowser1.Document.All("SIP_AlwaysFwdNum_RW").value = "呼叫转移未启用"
Call CCC
'MsgBox "           设置完成请重新打开本程序 ！！！", vbOKOnly, "XIAOKONGS實驗室"
'End

End Function

Private Function Calling(Number As String)

On Error Resume Next

'这是打电话的函数
Dim vDoc As Object
Dim vTag
Dim i As Integer

Set vDoc = WebBrowser1.Document
For i = 0 To vDoc.All.length - 1
If UCase(vDoc.All(i).tagName) = "INPUT" Then
Set vTag = vDoc.All(i)
If vTag.Type = "submit" Then
Select Case vTag.Name
Case "AutoDialSubmit"
vTag.Click   '拨通
'Label5.Caption = "拨号完成 ！！！"
Command2.Enabled = True
Exit Function
End Select
End If
End If
Next i

End Function
Private Function StopCaller()

On Error Resume Next

'这是挂断电话的函数
Dim vDoc, vTag
Dim i As Integer
Set vDoc = WebBrowser1.Document
For i = 0 To vDoc.All.length - 1
If UCase(vDoc.All(i).tagName) = "INPUT" Then
Set vTag = vDoc.All(i)
If vTag.Type = "submit" Then
Select Case vTag.Name
Case "HangupSubmit"
vTag.Click
Exit Function
End Select
End If
End If
Next i
End Function


Public Function DelChina_Find_SH()

  On Error Resume Next

  Dim S As String
  S = ""
  For i = 1 To Len(Combo1.Text)
   If Asc(Mid(Combo1.Text, i, 1)) > 0 Then S = S + Mid(Combo1.Text, i, 1)
  Next i
     Label5.Caption = ">>正在计算归属地"
     
  Combo1.Text = Trim(S)
  Combo1.Tag = GetLocation(Combo1.Text)
  Debug.Print Combo1.Tag
  
'  If Combo1.Tag = 9 Then MsgBox "这是上海手机"

End Function


Public Function getWebContent() As Long

    On Error Resume Next

    Dim doc As Object
    Dim i As Object
    Dim strHtml As String
 
    
    Set doc = WebBrowser2.Document
    For Each i In doc.All
'        strHtml = strHtml & Chr(13) & i.innerText
         
         '获取号码归属地
         If Len(i.innerText) = 8 Then
         
         Debug.Print i.innerText
         
         If Left(i.innerText, 4) = "归属城市" Then
            Address = Right(i.innerText, 3)

            If Address = "上海 " Then
               getWebContent = 9
               Exit Function
            End If
         
         End If
         
         End If
    Next
    
    getWebContent = 90
    
End Function

Public Function GetLocation(TEL As String) As Long

    On Error Resume Next

    WebBrowser2.Navigate "http://guishu.showji.com/search.htm?m=" & TEL

    DoEvents
    
    Do While WebBrowser2.Busy '等待加载完成.
    DoEvents
    Loop
    
    DoEvents
    
    GetLocation = Location

End Function

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If Not WebBrowser1.Busy Then '判断网页是否下载完毕
Location = getWebContent
End If
End Sub


Public Sub USA()

On Error Resume Next

WebBrowser1.Navigate "http://" & E & "/webdial.htm"

'
Do While WebBrowser1.Busy '等待加载完成.
    DoEvents
    Label5.Caption = ">>紧急拨号中。。。"
    Command2.Enabled = False
Loop

MsgBox "    紧急呼叫 ！！！", vbOKOnly, "XIAOKONGS實驗室"

Label5.Caption = "紧急呼叫"

DoEvents

WebBrowser1.Document.All("WEB_DialNumber").value = Combo1.Text

Calling Combo1.Text


End Sub
