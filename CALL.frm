VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Caller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   Icon            =   "CALL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   3840
   StartUpPosition =   2  '屏幕中心
   Begin Cvoip.Abutton Command8 
      Height          =   375
      Left            =   2760
      TabIndex        =   27
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   7
      CaptionAlignment=   0
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "关闭"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Cvoip.Abutton command9 
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   7
      CaptionAlignment=   0
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "拨号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   5520
      Width           =   855
      ExtentX         =   1508
      ExtentY         =   661
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
      Location        =   ""
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   7800
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sample"
      Height          =   615
      Left            =   6840
      TabIndex        =   18
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "新建联系人"
      Height          =   270
      Left            =   1560
      TabIndex        =   17
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1200
      TabIndex        =   16
      Text            =   "王西平 18936877668"
      Top             =   6600
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   3840
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清空表格"
      Height          =   540
      Left            =   6240
      TabIndex        =   12
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Text            =   "18936877668"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Text            =   "9306"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   4
      Text            =   "1"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "请输入："
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2895
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1320
         TabIndex        =   14
         Text            =   "A1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "机柜："
         Height          =   180
         Left            =   600
         TabIndex        =   13
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "序号："
         Height          =   180
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "型号标识："
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "运营商："
         Height          =   180
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   6480
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   6480
      ScaleHeight     =   3795
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "读取数据库"
      Height          =   405
      Left            =   7440
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "数据排序"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   25
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture3 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   26
      Top             =   0
      Width           =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "更新通讯录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   4635
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "编辑联系人"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4630
      Width           =   900
   End
End
Attribute VB_Name = "Caller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Dim staff(99) As stafftype
Dim n%, i%, kk%, jj%
Dim str As String


'新增数据
Private Sub AddData()
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
'        MsgBox "信息不可为空！", 0, "提示"
    Exit Sub
    End If
    If n >= 100 Then
        MsgBox ("输入人数超出限制100")
    ElseIf Not IsNumeric(Text1) Then
        MsgBox ("工号必须为数字")
        Text1 = ""
        Text1.SetFocus
    Else
         With staff(n)
         .num = n + 1
         .Name = Text2
         .IDC = Text3
         .Location = ""
        End With
'        Text1 = "": Text2 = "": Text3 = "": Text4 = ""
        n = n + 1
    End If
   
'Picture1.Cls
'Picture1.Print "     序号    型号标识    运营商    机柜"
'Picture1.Print "--------------------------------------------"
'
  
    
    For i = 0 To n - 1
    
        With staff(i)
'            Picture1.Print Tab(8); Trim(.num); Tab(16); Trim(.Name); Tab(27); Trim(.IDC); "      "; Trim(.Location)
'             Picture1.Print Tab(7); .num; Tab(15); .Name; Tab(27); .IDC; Tab(27); .Location;
          str = .num & "." & Trim(.Name) & "." & Trim(.IDC)

        End With
    Next i
             
             List1.AddItem str
              Open "D:\mydata.txt" For Append As #1
                Print #1, str
              Close #1
              
    ShowBody
    
End Sub

Private Sub Command1_Click()
AddData
End Sub

Private Sub Command2_Click()
    i = 1
    If i = 1 Then
       
        Open "D:\mydata.txt" For Output As #1
         Close #1
         Picture1.Cls
         List1.Clear
         n = 0
    Else
        Exit Sub
    End If
End Sub

Private Sub Command3_Click()
Picture1.Cls
End Sub

'读取数据
Private Sub Command4_Click()



End Sub

Private Sub Command5_Click()

On Error Resume Next

GetToTEXT Text5.Text   '获取数据
AddData   '写入数据
ShowBody  '显示数据

End Sub

'Download by http://www.NewXing.com
Private Sub Command7_Click()

On Error Resume Next
Dim num, Name, loca, salary As String
Dim backup(99) As stafftype
Picture1.Cls
Picture1.Print "     序号    型号标识    运营商    机柜"
Picture1.Print "--------------------------------------------"


Open "D:\mydata.txt" For Input As #1

Do While Not EOF(1)
    Input #1, num, Name, salary
    With backup(n)
         .num = num
         .Name = Name
         .IDC = salary
         .Location = loca
    End With
    n = n + 1
Loop
    For i = 0 To n - 1
        With backup(i)
            Picture1.Print Tab(6); Trim(.num); Tab(15); Trim(.Name); Tab(25); Trim(.IDC); Tab(25); Trim(.Location)
'            Print .num; .name; .IDC
        End With
    Next i
    
Close #1

End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Command9_Click()
MsgBox "          正在拼命研发中！！！", vbOKOnly, "XIAOKONGS實驗室"
End Sub

Private Sub Form_Load()
Dim rtn
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
ShowBody
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #1
    Form1.Show
End Sub

'获取文本文件行数

Function GetFileNumber() As Integer

On Error Resume Next

Dim i As Integer
Dim iFileNum As Integer
Dim str As String
  iFileNum = FreeFile
Open "D:\mydata.txt" For Input As #iFileNum '这里要完整的路径和文件名
   Do While Not EOF(iFileNum)
    i = i + 1
    Line Input #iFileNum, str
     'if len(str)=0 then i=i-1'这一行你可以选用,其目的是忽略空行
    Loop
Close #iFileNum
'   MsgBox "该文件共 " & i & "行"
GetFileNumber = i
End Function

Public Function A()

End Function

Private Sub Label2_Click()
 On Error Resume Next
 ShellExecute Me.hwnd, vbNullString, "D:\mydata.txt", vbNullString, "", 1
End Sub

Private Sub Label3_Click()
ShowBody
MsgBox "更新成功~请重新打開本程序", vbOKOnly, "XIAOKONGS實驗室"
End
End Sub

Private Sub Label7_Click()
MsgBox GetLocation("13601609004")
End Sub

Private Sub List1_Click()

Dim connect As Integer
connect = List1.ListIndex

Text6.Tag = GetLocation(Combo1.List(connect))
Text6 = "Call：" & Combo1.List(connect)

End Sub

Private Sub Picture1_Click()
SavePicture Picture1.Image, "D:\PIC.jpeg"
End Sub

Public Function getCrlfCount(ByVal SearchString As String) As Long

    On Error Resume Next
    
    Dim ret As Long
    Dim StrArray() As String
    StrArray = Split(SearchString, 9306)
    ret = UBound(StrArray) + 1
    getCrlfCount = ret
    
End Function

Public Function GetToTEXT(ByVal SearchString As String) As Long

    On Error Resume Next
    
    Dim ret As Long
    Dim StrArray() As String
    StrArray = Split(SearchString, " ")
    ret = UBound(StrArray) + 1
    
    If ret > 0 Then
        Text2 = StrArray(0)
        Text3 = StrArray(1)
    End If
    
    GetData = ret
    
End Function

'显示数据
Public Function ShowBody()

On Error Resume Next

n = GetFileNumber
Me.Caption = "联系人：" & n & "位"
Dim num, Name, salary, loca As String
On Error Resume Next
Picture1.Cls
List1.Clear
Form1.Combo1.Clear
Combo1.Clear

Form1.Combo1.Text = ">>请选择您要拨打的号码"

'Picture1.Print "     序号      姓名       联系电话    "
'Picture1.Print "--------------------------------------------"


'List1.AddItem "     序号        姓名         "
'List1.AddItem "--------------------------------------------"

'Debug.Print getCrlfCount(str)

Dim BeautyString As String
Dim StrArray() As String
Dim ret As Long
Dim A, b, C As String

Open "D:\mydata.txt" For Input As #1
    Do While Not EOF(1)
    
        Input #1, BeautyString
        StrArray = Split(BeautyString, ".")
        ret = UBound(StrArray) + 1
            If ret > 0 Then
                A = Trim(StrArray(0))
                b = Trim(StrArray(1))
                C = Trim(StrArray(2))
                
                
'        MsgBox Len("     序号      姓名        联")
        BeautyString = "      " & A & "         " & b '& "      " & C

                
'                If A < 10 Then
'                   If Len(B) < 3 Then B = " " & B
'                   BeautyString = "      " & A & "       " & B & "      " & C
'                Else
'                    If Len(B) < 3 Then
'                    B = " " & B
'                    Else
'                    B = "" & B
'                    End If
'                   BeautyString = "     " & A & "       " & B & "      " & C
'                End If
                
            End If
            
        List1.AddItem BeautyString
        i = i + 1
        Combo1.AddItem C
        Form1.Combo1.AddItem C & " " & b

        
    Loop
Close #1

End Function


Private Function getWebContent() As Long

    On Error Resume Next

    Dim doc As Object
    Dim i As Object
    Dim strHtml As String
    
    Set doc = WebBrowser1.Document
    For Each i In doc.All
'        strHtml = strHtml & Chr(13) & i.innerText
         Debug.Print i.innerText
         If Trim(i.innerText) = "归属城市：上海 " Then
         getWebContent = 9
         Exit Function
         End If
    Next
    getWebContent = 90
    
End Function

Private Function GetLocation(TEL As String) As Long

    On Error Resume Next

    WebBrowser1.Navigate "http://guishu.showji.com/search.htm?m=" & TEL
'WebBrowser1.Navigate ("http://www.ip.cn/db.php?num=13601609004")   '连接VOIP
   DoEvents
    
    Do While WebBrowser1.Busy '等待加载完成.
    DoEvents
    Loop
    
    GetLocation = getWebContent

End Function

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'
''Browser.Navigate "http://hi.baidu.com"
''Do Until WebBrowser1.ReadyState = 4
''DoEvents
''Loop
'
'Do While WebBrowser1.Busy '等待加载完成.
'    DoEvents
'    Loop
'
'   Dim doc As Object
'    Dim i As Object
'    Dim strHtml As String
'
'    Set doc = WebBrowser1.Document
'    For Each i In doc.All
''        strHtml = strHtml & Chr(13) & i.innerText
''         Debug.Print i.innerText
'         If Trim(i.innerText) = "卡号归属地上海" Then
''         getWebContent = 9
'         Exit Sub
'         End If
'    Next
''    getWebContent = 90
''Location = getWebContent

End Sub

