VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "聊天窗口"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9195
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   5040
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "刷新"
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   5280
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   4560
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5835
      Left            =   9600
      TabIndex        =   5
      Top             =   0
      Width           =   8295
      ExtentX         =   14631
      ExtentY         =   10292
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
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "退出"
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "发送"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4560
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1560
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   960
      Picture         =   "Form2.frx":0557
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Height          =   180
      Left            =   1080
      TabIndex        =   8
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "、"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4200
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public Mingzi As String
Public suoyou As String
Public lishi As String
Public denglu As String
Public doudong As String
Public jiazaicishu As String
Public dd2 As String
Public szMyText, szMyText2 As String
Public lishi2 As String
Private oShadow As New aShadow
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Type WSADATA
        wversion As Integer
        wHighVersion As Integer
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpszVendorInfo As Long
    End Type
    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
    Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHostname As String) As Long
    Private Const WS_VERSION_REQD = &H101
    '-------------------------------------------------------------------------------------------------------------------------
    Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const MAX_PATH As Integer = 260
Const TH32CS_SNAPPROCESS As Long = 2&
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
    szExeFile As String * 1024
    End Type
Private Declare Function FINDWINDOW Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

    Public Function IsConnectedState() As Boolean
        Dim udtWSAD As WSADATA
        Call WSAStartup(WS_VERSION_REQD, udtWSAD)
        IsConnectedState = CBool(gethostbyname("www.baidu.com"))
        Call WSACleanup
    End Function

Private Sub Command1_Click()
If IsConnectedState Then '检查网络连接
WebBrowser1.Navigate "http://notepad.live/kart3"
'-------------------------------------------发送中心-------------------------------------------------
Dim fasongneirong As String
fasongneirong = Text1.Text
lishi2 = szMyText

Dim vDoc, VTag, mType As String, mTagName As String
Dim ia As Integer
    Set vDoc = WebBrowser1.Document
    For ia = 0 To vDoc.All.Length - 1
        Select Case UCase(vDoc.All(ia).tagName)
        Case "TEXTAREA"     '"TEXTAREA" 标签,文本框的填写
        Set VTag = vDoc.All(ia)
         VTag.Value = "名字" & Mingzi & ":" & fasongneirong & "结束" '将Text1中的内容填入
         Debug.Print ("发送内容：" & szMyText & "名字" & Mingzi & ":" & fasongneirong & "结束")
         End Select
Text1.Text = ""
Timer1.Enabled = True
Timer1.Interval = 1000
Command1.Enabled = False
Command3.Enabled = False
Next ia
Else
Label4.Caption = "网络未连接"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Command3_Click()
If IsConnectedState Then
WebBrowser1.Navigate "http://notepad.live/kart3"
Timer1.Enabled = True
Timer1.Interval = 1000
Command1.Enabled = False
Command3.Enabled = False
Else
Label4.Caption = "网络未连接"
End If
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
Fujincheng
GuanbiFindwindow ("Numega SmartCheck")
Debug.Print ("已关闭Numega SmartCheck")
WebBrowser1.Silent = True
'-------------------------------------------------名字显示----------------------------------------
Open App.Path & "\xxx" For Input As #2
Line Input #2, Mingzi
Close #2
If Mingzi = "" Then
Mingzi = "游客"
End If
'-------------------------------------------------检查网络连接------------------------------------
If IsConnectedState Then
WebBrowser1.Navigate "http://notepad.live/kart3"
Label4.Caption = "正在获取内容"
denglu = "no"
Command1.Enabled = False
Command3.Enabled = False
Timer1.Enabled = True
Timer1.Interval = 1000
Else
Label4.Caption = "网络未连接"
Command1.Enabled = False
Command3.Enabled = False
End If
'--------------------------------------------周边阴影---------------------------------------------
With oShadow
    If .Shadow(Me) Then
        .Depth = 7 '阴影宽度
        .Color = RGB(0, 0, 0) '阴影颜色
        .Transparency = 50 '阴影色深
    End If
 End With
 '---------------------------------------------文件创建读取-----------------------------------------
 If Dir(App.Path & "\history") = "" Then
 Open App.Path & "\history" For Output As #3
 Print #3, ""
 Close #3
Open App.Path & "\history" For Binary As #1
  lishi = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
 Label3.Caption = lishi
 Else
 Open App.Path & "\history" For Binary As #1
  lishi = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
Label3.Caption = lishi
 End If
 Timer2.Enabled = False
Timer2.Interval = 50 '抖动频率
End Sub

Private Sub Image1_Click()
cd1.ShowColor
Label1.BackColor = cd1.Color
Label2.BackColor = cd1.Color
End Sub

Private Sub Image2_Click()
doudong = 0
dd2 = 15
Timer2.Enabled = True
Text1.Text = Mingzi & "发送了一个窗口抖动"
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
jiazaicishu = jiazaicishu + "1"
Debug.Print ("加载次数：" & jiazaicishu)
'--------------------------------------------接收中心------------------------------------------
If WebBrowser1.Busy Then
Debug.Print ("网页未加载完成")
        Exit Sub
    Else
    Debug.Print ("网页加载完成")
Timer1.Enabled = False
WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
Dim vDoc, x, VTag
Set vDoc = WebBrowser1.Document
For x = 0 To vDoc.All.Length - 1 '检测所有标签
If UCase(vDoc.All(x).tagName) = "INPUT" Then '找到input标签
Set VTag = vDoc.All(x)
If VTag.Value = "提交" Then VTag.Click '点击提交了，一切都OK了
End If
Next x
denglu = "yes"
Timer3.Enabled = True
Timer3.Interval = 1500
Debug.Print ("登录状态：" & "yes")
End If
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then  '如果，是回车键按下
If Command1.Enabled = True Then
Call Command1_Click
Else
End If
End If
End Sub

Private Sub Form_DblClick()
Unload Me
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Timer2_Timer()
dd2 = dd2 - 1
If dd2 = "0" Then
Timer2.Enabled = False
Else
If doudong = 0 Then
Form2.Top = Form2.Top + 80
doudong = doudong + 1
ElseIf doudong = 1 Then
Form2.Left = Form2.Left + 80
doudong = doudong + 1
ElseIf doudong = 2 Then
Form2.Top = Form2.Top - 80
doudong = doudong + 1
ElseIf doudong = 3 Then
Form2.Left = Form2.Left - 80
doudong = 0
End If
End If
End Sub

Private Sub Timer3_Timer()
If WebBrowser1.Busy Then
Debug.Print ("网页未加载完成")
        Exit Sub
    Else
    Timer3.Enabled = False
Debug.Print ("开始提取文字")
Dim szText As String
Dim szFindStrBegin As String
Dim szFindStrEnd As String
Dim nBegin As Long
Dim nEnd As Long
Dim nLength  As Long
szFindStrBegin = "名字" '定义要查找的字符串开头
szFindStrEnd = "结束" '定义要查找的字符串结尾

szText = WebBrowser1.Document.body.innerText '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText

nBegin = InStr(szText, szFindStrBegin) '找开头字符串
If nBegin > 0 Then '必须有能找到开头了才继续
    nEnd = InStr(nBegin, szText, szFindStrEnd) '找结尾字符串
    If nEnd > nBegin Then '结尾必须比开头的位置大
    
        '包含查找的字符串模式，注释掉下面的2行
        nLength = nEnd - nBegin + Len(szFindStrEnd)   '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
        
        '不包含查找的字符串模式
        nLength = nEnd - nBegin - Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
        nBegin = nBegin + Len(szFindStrBegin) '如果不包含查找的字符串，用这2行
        
        szMyText = Mid(szText, nBegin, nLength)   '取出“before then.”到 "test" 中间的东西
Debug.Print ("截取内容：" & szMyText)
    End If
End If
Label4.Caption = ""
Open App.Path & "\history" For Output As #3
Print #3, lishi
Print #3, szMyText
Close #3
Dim Sm As String
Dim hang As String
hang = 0
Open App.Path & "\history" For Input As #1
Do While Not EOF(1)
Line Input #1, Sm
hang = hang + 1
Loop
Debug.Print ("历史文件行数：" & hang)
Close #1
Open App.Path & "\history" For Binary As #1
Dim S As String
  lishi = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
  If hang > 19 Then
  If Dir(App.Path & "\history2") = "" Then
Name App.Path & "\history" As App.Path & "\history" & "2"
Else
Kill App.Path & "\history2"
End If
Open App.Path & "\history" For Output As #4
Print #4, ""
Close #4
Debug.Print ("转行内容：" & lishi2)
lishi = lishi2 & Chr(13)
  Else
  End If
Label3.Caption = lishi
Text1.SetFocus
End If
Command1.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Fujincheng()
'这个过程是检测父进程的父进程是否是EXPLORE的父进程
Dim Process As PROCESSENTRY32
Dim hSnapShot As Long
Dim XNN As Long
Dim flag As Boolean
Dim mName As String
Dim i As Integer
Dim pid As Long, explorer As Long '注意这2个变量就用来存放2个ID
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&) '建立进程快照 '搜索explorer.exe进程,并获得其ID   If hSnapShot Then
    Process.dwSize = 1060
    If (Process32First(hSnapShot, Process)) Then '遍历第一个进程,获得PROCESSENTRY32结构
      Do
        i = InStr(1, Process.szExeFile, Chr(0))       '获得映像名称         mName = LCase(Left(Process.szExeFile, i - 1)) '并转换成小写
        If mName = "explorer.exe" Then      '是不是explorer.exe         explorer = Process.th32ProcessID
        ElseIf mName = LCase(App.EXEName & ".exe") Then  '是不是自己              pid = Process.th32ParentProcessID   '获得父进程ID         Else
             flag = False
             End If
      Loop Until (Process32Next(hSnapShot, Process) < 1) '遍历所有进程直到返回值为False     End If
    XNN = CloseHandle(hSnapShot)
    End If
Dim Openit As Long
Openit = OpenProcess(1&, -1&, pid)
If pid <> explorer Then
MsgBox "发现进程调试", , "警告": TerminateProcess Openit, 0
'如果发现父进程不对，就结束掉父进程，对使用HIDEOD中的Process32NEXT的OD无效
End
End If
End Sub

Private Sub GuanbiFindwindow(Mingzi As String)
Dim winHwnd As Long
winHwnd = FINDWINDOW(vbNullString, Mingzi) '获得窗口句柄
If winHwnd <> 0 Then AppActivate Mingzi '激活窗体为活动
SendKeys "%{f4}", True 'ALT＋F4 结束掉
SendKeys "%y", True
End Sub

