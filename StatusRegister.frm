VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form StatusRegister 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      Picture         =   "StatusRegister.frx":0000
      ScaleHeight     =   4425
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin SystemANTI.Abutton Abutton1 
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "确定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ListView PsListView 
         Height          =   4215
         Left            =   240
         TabIndex        =   7
         Top             =   4440
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7435
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "警告 XIAOKONGS 實驗室 拥有该程序的专有权"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   600
         TabIndex        =   12
         Top             =   4105
         Width           =   3690
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "经过 XIAOKONGS 授权方可正常使用。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   2970
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   " 请联系 XIAOKONGS 为您注册(免费)。"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   3060
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   " 如果您想在个人计算机上使用本程序"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   2970
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1200
         Picture         =   "StatusRegister.frx":657E
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "本程序设计用于保障 XIAOKONGS 秘密计算机基础设施"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   4230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "232323232"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   5
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "121212121"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   3000
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "注册到 :"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   3
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "（许可：XIAOKONGS 實驗室）"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   3660
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[使用者注册授权]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   2385
      End
   End
End
Attribute VB_Name = "StatusRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Abutton1_Click()
On Error Resume Next
If Abutton1.Caption = "注册" Then Call RegRunExeCRC
Unload Me
End Sub

Private Sub Form_Load()

'MsgBox FileLen(App.path & "\" & App.EXENAME & ".exe")
'MsgBox GetCheckCRC(App.path & "\" & App.EXENAME & ".exe")

Dim rtn
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)

GetUserCom

If valMyExeCRC = "0" Then
Label2.Caption = "未注册"
Abutton1.Caption = "注册"
Label5.Caption = "（许可：未注册 用户）"
Else
Label2.Caption = "注册到 :" & valMyExeCRC
Label5.Caption = "（许可：XIAOKONGS 實驗室）"
End If

End Sub

'重要
Public Sub RegRunExeCRC()

On Error Resume Next

Dim ListWith As Integer
Dim Proc As PROCESSENTRY32
Dim snap As Long
Dim theloop As Long
Dim SplitSpace() As String
Dim ListLineNum As Integer

    With PsListView
        ListWith = (.Width - 60) \ 24
        .View = lvwReport
        With .ColumnHeaders
'            .Add , , "进程名称", 9 * ListWith
'            .Add , , "进程PID", 4 * ListWith
'            .Add , , "CRC效验", 6 * ListWith     'SubItems(2) 校验码
'            .Add , , "进程路径", 30 * ListWith   'SubItems(3)名称
        End With
    End With

    EnablePrivilege
    ''''''''''''''''''''''''''''''''''''首次运行更新列表1次
    snap = CreateToolhelpSnapshot(15, 0)
    Proc.dwsize = Len(Proc)
    theloop = ProcessFirst(snap, Proc)
    With PsListView
            Do While theloop <> 0
                    If Proc.th32ProcessID <> GetCurrentProcessId Then
                        If HaveItem(PsListView, Proc.th32ProcessID) = 0 Then
                            SplitSpace = Split(Proc.szExeFile, "")

                            '''''''''''''''''''
                            With .ListItems.Add(, "ID:" & CStr(Proc.th32ProcessID), SplitSpace(LBound(SplitSpace)))
                                .SubItems(1) = Proc.th32ProcessID
'                                .SubItems(2) = Proc.th32ParentProcessID   '父进程

                                If GetProcessImageFileName(Proc.th32ProcessID) <> "NULL" Then .SubItems(2) = GetCheckCRC(GetProcessImageFileName(Proc.th32ProcessID))

                                 .SubItems(3) = GetProcessImageFileName(Proc.th32ProcessID)

                                If valMyExeCRC = "0" Then BakSafeEXE GetFileName(.SubItems(3)), .SubItems(2)

                            End With
                        End If
                    End If
                theloop = ProcessNext(snap, Proc)
            Loop
'        PsNum.Caption = .ListItems.count


    End With
    CloseHandle snap
    CheckProcess PsListView
    '''''''''''''''''''''''''''''''''''''''''''''''
''    GetPsList.Enabled = True
'    GetPsList.Interval = Val(ChangeTimer.Text)

    Me.Height = 5055
    
    BakSafeEXE App.EXENAME, GetCheckCRC(App.path & "\" & App.EXENAME & ".exe")
    Unload Me
    
End Sub

