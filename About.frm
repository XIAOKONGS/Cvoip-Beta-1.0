VERSION 5.00
Begin VB.Form StatusRegister 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      Picture         =   "About.frx":0000
      ScaleHeight     =   4425
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin Cvoip.Abutton Abutton2 
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   3360
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
         Caption         =   "����"
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
      Begin Cvoip.Abutton Abutton1 
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   7
         CaptionAlignment=   0
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "����"
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cvoip 1.0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3600
         TabIndex        =   16
         Top             =   880
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ӧ�þ����� ������"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   1320
         Width           =   2070
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>XiaoKongs.org"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         MouseIcon       =   "About.frx":657E
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   3960
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ": XIAOKONGS �����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   960
         TabIndex        =   12
         Top             =   3480
         Width           =   2190
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ": XIAOKONGS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   960
         TabIndex        =   9
         Top             =   2760
         Width           =   2190
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "��˾"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ": �Ϻ��γ����缼�����޹�˾"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XIAOKONGS ���ṩ�κ������ϵĵ���."
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   5
         Top             =   1800
         Width           =   2970
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   " ����ϵ XIAOKONGS ��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   2340
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   " �������ʹ���з����κ�������߸��õ��뷨"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   3
         Top             =   2160
         Width           =   3690
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1200
         Picture         =   "About.frx":66D0
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������������Զ�̹���VOIP�칫�绰����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   480
         TabIndex        =   2
         Top             =   1560
         Width           =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[����- �γ�����VOIP�绰ϵͳ]"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4185
      End
   End
End
Attribute VB_Name = "StatusRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Abutton1_Click()
Unload Me
End Sub

Private Sub Abutton2_Click()
Shell "explorer.exe http://xiaokongs.org/x/SendXiaoKongs.asp"
End Sub

Private Sub Form_Load()

'MsgBox FileLen(App.path & "\" & App.EXENAME & ".exe")
'MsgBox GetCheckCRC(App.path & "\" & App.EXENAME & ".exe")

Dim rtn
rtn = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)

'GetUserCom
'
'If valMyExeCRC = "0" Then
'Label2.Caption = "δע��"
'Abutton1.Caption = "ע��"
'Label5.Caption = "����ɣ�δע�� �û���"
'Else
'Label2.Caption = "ע�ᵽ :" & valMyExeCRC
'Label5.Caption = "����ɣ�XIAOKONGS ����ң�"
'End If

End Sub

'��Ҫ
'Public Sub RegRunExeCRC()
'
'On Error Resume Next
'
'Dim ListWith As Integer
'Dim Proc As PROCESSENTRY32
'Dim snap As Long
'Dim theloop As Long
'Dim SplitSpace() As String
'Dim ListLineNum As Integer
'
'    With PsListView
'        ListWith = (.Width - 60) \ 24
'        .View = lvwReport
'        With .ColumnHeaders
''            .Add , , "��������", 9 * ListWith
''            .Add , , "����PID", 4 * ListWith
''            .Add , , "CRCЧ��", 6 * ListWith     'SubItems(2) У����
''            .Add , , "����·��", 30 * ListWith   'SubItems(3)����
'        End With
'    End With
'
'    EnablePrivilege
'    ''''''''''''''''''''''''''''''''''''�״����и����б�1��
'    snap = CreateToolhelpSnapshot(15, 0)
'    Proc.dwSize = Len(Proc)
'    theloop = ProcessFirst(snap, Proc)
'    With PsListView
'            Do While theloop <> 0
'                    If Proc.th32ProcessID <> GetCurrentProcessId Then
'                        If HaveItem(PsListView, Proc.th32ProcessID) = 0 Then
'                            SplitSpace = Split(Proc.szExeFile, "")
'
'                            '''''''''''''''''''
'                            With .ListItems.Add(, "ID:" & CStr(Proc.th32ProcessID), SplitSpace(LBound(SplitSpace)))
'                                .SubItems(1) = Proc.th32ProcessID
''                                .SubItems(2) = Proc.th32ParentProcessID   '������
'
'                                If GetProcessImageFileName(Proc.th32ProcessID) <> "NULL" Then .SubItems(2) = GetCheckCRC(GetProcessImageFileName(Proc.th32ProcessID))
'
'                                 .SubItems(3) = GetProcessImageFileName(Proc.th32ProcessID)
'
'                                If valMyExeCRC = "0" Then BakSafeEXE GetFileName(.SubItems(3)), .SubItems(2)
'
'                            End With
'                        End If
'                    End If
'                theloop = ProcessNext(snap, Proc)
'            Loop
''        PsNum.Caption = .ListItems.count
'
'
'    End With
'    CloseHandle snap
'    CheckProcess PsListView
'    '''''''''''''''''''''''''''''''''''''''''''''''
'''    GetPsList.Enabled = True
''    GetPsList.Interval = Val(ChangeTimer.Text)
'
'    Me.Height = 5055
'
'    BakSafeEXE App.EXEName, GetCheckCRC(App.Path & "\" & App.EXEName & ".exe")
'    Unload Me
    
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub

Private Sub Label19_Click()
Shell "explorer.exe http://xiaokongs.org"
End Sub
