VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置 - Cvoip Beta"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin Cvoip.Abutton Command3 
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   7
      ButtonStyleColors=   3
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
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   1560
      TabIndex        =   21
      Text            =   "9018936877668"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1560
      TabIndex        =   19
      Text            =   "127.0.0.1"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "机房办公电话IP地址"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   6720
      Width           =   4575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "注意：请正确填写WAN口地址，否则将无法操作您的电话"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   150
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   3675
      End
   End
   Begin VB.CheckBox Check5 
      Appearance      =   0  'Flat
      Caption         =   "Check1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   290
      TabIndex        =   14
      Top             =   2610
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   13
      Text            =   "将转移到某号码"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   12
      Text            =   "EEE"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "上海游弛网络技术有限公司 Cvoip 1.0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin Cvoip.Abutton Abutton1 
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         ButtonStyle     =   3
         CaptionAlignment=   0
         BorderColor     =   -2147483633
         BorderColorPressed=   -2147483628
         BorderColorHover=   -2147483627
         ForeColor       =   16576
         Caption         =   "帮助"
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
      Begin Cvoip.Abutton Command2 
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         CaptionAlignment=   0
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "搞定"
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
      Begin Cvoip.Abutton Command1 
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         CaptionAlignment=   0
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "立即解除呼叫转移"
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
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   160
         TabIndex        =   23
         Top             =   1050
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   160
         TabIndex        =   11
         Top             =   2010
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         TabIndex        =   10
         Text            =   "将转移到某号码"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   9
         Text            =   "DDD"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   160
         TabIndex        =   8
         Top             =   1530
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         TabIndex        =   7
         Text            =   "将转移到某号码"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   6
         Text            =   "CCC"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         TabIndex        =   5
         Text            =   "将转移到某号码"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   480
         TabIndex        =   4
         Text            =   "BBB"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   3
         Text            =   "AAA"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         TabIndex        =   2
         Text            =   "将转移到某号码"
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   160
         TabIndex        =   1
         Top             =   560
         Width           =   255
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "关于"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   4875
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "办公电话IP地址："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   3990
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "紧急呼叫号码："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   4380
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "【意见反馈】"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   4875
      Width           =   1260
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abutton1_Click()
MsgBox "          号码前必须添加“9”或者 “90”", vbOKOnly, "XIAOKONGS實驗室"
End Sub

Private Sub Check1_Click()

    If Check1.value Then
        Text1.Enabled = True
        Text2.Enabled = True

       Else
         Text1.Enabled = False
         Text2.Enabled = False
         
         Text1.Text = "AAA"

    End If

End Sub

Private Sub Check2_Click()

    If Check2.value Then
        Text3.Enabled = True
        Text4.Enabled = True

       Else
         Text3.Enabled = False
         Text4.Enabled = False
         
         Text3.Text = "BBB"

    End If

End Sub

Private Sub Check3_Click()

    If Check3.value Then
        Text5.Enabled = True
        Text6.Enabled = True

       Else
         Text5.Enabled = False
         Text6.Enabled = False
         
         Text5.Text = "CCC"

    End If

End Sub

Private Sub Check4_Click()

    If Check4.value Then
        Text7.Enabled = True
        Text8.Enabled = True

       Else
         Text7.Enabled = False
         Text8.Enabled = False
         
         Text7.Text = "DDD"

    End If

End Sub

Private Sub Check5_Click()

MsgBox "          正在拼命研发中！！！", vbOKOnly, "XIAOKONGS實驗室"


'    If Check5.Value Then
'        Text9.Enabled = True
'        Text10.Enabled = True
'
'       Else
'         Text9.Enabled = False
'         Text10.Enabled = False
'
'         Text9.Text = "EEE"
'
'    End If

End Sub



Private Sub Command3_Click()
Unload Me
End Sub

''
''Private Sub Command1_Click()
''MsgBox "呼叫转移解除成功"
''End Sub
'
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub


Private Sub Command1_Click()

On Error Resume Next

Form1.CleanNum
MsgBox "           呼叫转移已解除 ！！！", vbOKOnly, "XIAOKONGS實驗室"
Unload Me

End Sub
 
Private Sub Command2_Click()

SaveSetting "MyApp", "setup", "Text1", Text1.Text
SaveSetting "MyApp", "setup", "Text2", Text2.Text
SaveSetting "MyApp", "setup", "Text3", Text3.Text
SaveSetting "MyApp", "setup", "Text4", Text4.Text
SaveSetting "MyApp", "setup", "Text5", Text5.Text

SaveSetting "MyApp", "setup", "Text6", Text6.Text
SaveSetting "MyApp", "setup", "Text7", Text7.Text
SaveSetting "MyApp", "setup", "Text8", Text8.Text
SaveSetting "MyApp", "setup", "Text9", Text9.Text

SaveSetting "MyApp", "setup", "Text11", Text11.Text
SaveSetting "MyApp", "setup", "Text12", Text12.Text

MsgBox "           设置完成请重新打开本程序 ！！！", vbOKOnly, "XIAOKONGS實驗室"
End

End Sub

Private Sub Form_Load()

Dim rtn
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)

'Label1.Caption = GetSetting("MyApp", "setup", "Label1", Label1.Caption)
Text1.Text = GetSetting("MyApp", "setup", "Text1", Text1.Text)
Text2.Text = GetSetting("MyApp", "setup", "Text2", Text2.Text)
Text3.Text = GetSetting("MyApp", "setup", "Text3", Text3.Text)
Text4.Text = GetSetting("MyApp", "setup", "Text4", Text4.Text)
Text5.Text = GetSetting("MyApp", "setup", "Text5", Text5.Text)

Text6.Text = GetSetting("MyApp", "setup", "Text6", Text6.Text)
Text7.Text = GetSetting("MyApp", "setup", "Text7", Text7.Text)
Text8.Text = GetSetting("MyApp", "setup", "Text8", Text8.Text)
Text9.Text = GetSetting("MyApp", "setup", "Text9", Text9.Text)

Text11.Text = GetSetting("MyApp", "setup", "Text11", Text11.Text)
Text12.Text = GetSetting("MyApp", "setup", "Text12", Text12.Text)

If Text1.Text <> "AAA" Then Check1.value = 1
If Text3.Text <> "BBB" Then Check2.value = 1
If Text5.Text <> "CCC" Then Check3.value = 1
If Text7.Text <> "DDD" Then Check4.value = 1
If Text9.Text <> "EEE" Then Check5.value = 1




End Sub


Private Sub Label5_Click()
Me.Hide
StatusRegister.Show
End Sub
