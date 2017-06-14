Attribute VB_Name = "Module1"
    Public E As String
    Public str As String
    
    Type stafftype
        num As Integer
        Name As String * 10
        IDC As String * 11
        Location As String * 6
    End Type
    
    Public SmallHeight As String
    Public BigggHeight As String
    Public Urgent As String   '紧急呼叫号码
    
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function OS() As String

'MsgBox SystemVer
 Dim objWMIService, colItems, objItem, strOSversion As String
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objItem In colItems
        strOSversion = objItem.version
    Next
    Select Case Left(strOSversion, 3)
        Case "5.2": strOSversion = "Windows Server 2003": SmallHeight = 2850: BigggHeight = 4390: Form1.Height = SmallHeight: Form1.Combo1.BackColor = &H8000000E: Form1.Label4.Top = 2030: Form1.Label10.Top = 2150
        Case "5.0": strOSversion = "Windows 2000"
        Case "5.1": strOSversion = "Windows XP": Form1.BorderStyle = 3: SmallHeight = 2800: BigggHeight = 4300: Form1.Height = SmallHeight: Form1.Combo1.BackColor = &H8000000E: Form1.Label4.Top = 2100
        Case "6.0": strOSversion = "windows vista"
        Case "6.1": strOSversion = "Win7"
        Case "6.2": strOSversion = "Win8"
        Case "6.3": strOSversion = "Win8.1"
        Case "10.": strOSversion = "Win10"
        Case Else: strOSversion = "未知系统版本"
    End Select
'    MsgBox "你的操作系统是：" & strOSversion
    OS = strOSversion
    
End Function


Public Function NameInit()

    Dim A As String
    Dim b As String
    Dim C As String
    Dim D As String

    
    Dim i As Integer
    Dim ip As Integer
    
    If Dir("D:" & "\myData.txt") = "" Then
    Open "d:\" & "myData.txt" For Output As #1
    Print #1, ""
    Close #1
    End If

    A = GetSetting("MyApp", "setup", "Text1", A)
    b = GetSetting("MyApp", "setup", "Text3", b)
    C = GetSetting("MyApp", "setup", "Text5", C)
    D = GetSetting("MyApp", "setup", "Text7", D)
    E = GetSetting("MyApp", "setup", "Text11", E)
    
    '读取紧急呼叫号码
    Urgent = GetSetting("MyApp", "setup", "Text12", Urgent)
    



    Dim num1 As String
    Dim num2 As String
    Dim num3 As String
    Dim num4 As String

    num1 = GetSetting("MyApp", "setup", "Text2", num1)
    num2 = GetSetting("MyApp", "setup", "Text4", num2)
    num3 = GetSetting("MyApp", "setup", "Text6", num3)
    num4 = GetSetting("MyApp", "setup", "Text8", num4)
    

    If A <> "AAA" And A <> "" Then
    Form1.cm1.Caption = A   '填入cm1姓名
    Form1.cm1.Tag = num1    '填入cm1号码
    Form1.cm1.Visible = True
      i = i + 1
    Else
    Form1.cm1.Visible = False
    End If
    
    If b <> "BBB" And A <> "" Then
    Form1.cm2.Caption = b
    Form1.cm2.Tag = num2   '填入cm2号码
    Form1.cm2.Visible = True
      i = i + 1
    Else
    Form1.cm2.Visible = False
    End If
    
    If C <> "CCC" And A <> "" Then
    Form1.cm3.Caption = C
    Form1.cm3.Tag = num3    '填入cm3号码
    Form1.cm3.Visible = True
      i = i + 1
    Else
    Form1.cm3.Visible = False
    End If
    
    If D <> "DDD" And A <> "" Then
    Form1.cm4.Caption = D
    Form1.cm4.Tag = num4    '填入cm4号码
    Form1.cm4.Visible = True
      i = i + 1
    Else
    Form1.cm4.Visible = False
    End If
    
        If E <> "127.0.0.1" And A <> "" Then
'    Form1.cm4.Caption = d
'    Form1.cm4.Tag = num4    '填入cm4号码
'    Form1.cm4.Visible = True
      ip = ip + 1
'       Form1.Label9.Caption = "拨打电话"
'        Form1.Label9.Visible = True
'         Form1.Label9.Left = 1600
'    Else
'    Form1.cm4.Visible = False
    End If
    
'        If A = "" And B = "" And C = "" And D = "" Then i = 0
    
'---------------------------------------邪恶分割线--------------------------------------------------------------
'    Debug.Print i

    
    If i = 0 And isLoaded("Form3") = False And Form1.Label9.Visible = False Then
'    Form1.Hide
'    MsgBox "             感谢使用请先设置参数 ！！！", , "XIAOKONGS 實驗室"
'    Form3.Show
     Form1.Label9.Visible = True
'     Form3.Text11 = "127.0.0.1"
     Form1.Label10.Visible = False
    End If
    
      If i = 1 Then
      
    Form1.cm1.Top = 1580
'    Form1.cm1.Left = 1440
'    Form1.cm2.Top = 1800
    End If
    
    If i = 2 Then
    Form1.cm1.Top = 1320
    Form1.cm1.Left = 1440
    Form1.cm2.Top = 1800
    End If
    
    If i = 3 Then
    Form1.cm1.Left = Form1.cm4.Left
    Form1.cm1.Top = Form1.cm4.Top
    Form1.cm2.Top = 1680
    End If
    
    If i = 4 Then
    Form1.cm1.Left = 1440
    Form1.cm1.Top = 1250
    End If
    
    

End Function

Public Function NumberInit()
 
    
End Function

Function isLoaded(strForm As String) As Boolean
'参数为窗体名
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = strForm Then
            isLoaded = True
            Exit Function
        End If
    Next
End Function
