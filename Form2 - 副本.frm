VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "���촰��"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9195
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   5040
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "ˢ��"
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
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "�˳�"
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "����"
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
      Picture         =   "Form2 - ����.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   960
      Picture         =   "Form2 - ����.frx":0557
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
      Caption         =   "��"
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public mingzi As String
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

    Public Function IsConnectedState() As Boolean
        Dim udtWSAD As WSADATA
        Call WSAStartup(WS_VERSION_REQD, udtWSAD)
        IsConnectedState = CBool(gethostbyname("www.baidu.com"))
        Call WSACleanup
    End Function

Private Sub Command1_Click()
If IsConnectedState Then '�����������
WebBrowser1.Navigate "http://notepad.live/kart3"
'-------------------------------------------��������-------------------------------------------------
Dim fasongneirong As String
fasongneirong = Text1.Text
lishi2 = szMyText

Dim vDoc, VTag, mType As String, mTagName As String
Dim ia As Integer
    Set vDoc = WebBrowser1.Document
    For ia = 0 To vDoc.All.Length - 1
        Select Case UCase(vDoc.All(ia).tagName)
        Case "TEXTAREA"     '"TEXTAREA" ��ǩ,�ı������д
        Set VTag = vDoc.All(ia)
         VTag.Value = "����" & mingzi & ":" & fasongneirong & "����" '��Text1�е���������
         Debug.Print ("�������ݣ�" & szMyText & "����" & mingzi & ":" & fasongneirong & "����")
         End Select
Text1.Text = ""
Timer1.Enabled = True
Timer1.Interval = 1000
Command1.Enabled = False
Command3.Enabled = False
Next ia
Else
Label4.Caption = "����δ����"
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
Label4.Caption = "����δ����"
End If
End Sub

Private Sub Form_Load()
WebBrowser1.Silent = True
'-------------------------------------------------������ʾ----------------------------------------
Open App.Path & "\xxx" For Input As #2
Line Input #2, mingzi
Close #2
If mingzi = "" Then
mingzi = "�ο�"
End If
'-------------------------------------------------�����������------------------------------------
If IsConnectedState Then
WebBrowser1.Navigate "http://notepad.live/kart3"
Label4.Caption = "���ڻ�ȡ����"
denglu = "no"
Command1.Enabled = False
Command3.Enabled = False
Timer1.Enabled = True
Timer1.Interval = 1000
Else
Label4.Caption = "����δ����"
Command1.Enabled = False
Command3.Enabled = False
End If
'--------------------------------------------�ܱ���Ӱ---------------------------------------------
With oShadow
    If .Shadow(Me) Then
        .Depth = 7 '��Ӱ���
        .Color = RGB(0, 0, 0) '��Ӱ��ɫ
        .Transparency = 50 '��Ӱɫ��
    End If
 End With
 '---------------------------------------------�ļ�������ȡ-----------------------------------------
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
Timer2.Interval = 50 '����Ƶ��
End Sub


Private Sub Image2_Click()
doudong = 0
dd2 = 15
Timer2.Enabled = True
Text1.Text = mingzi & "������һ�����ڶ���"
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
jiazaicishu = jiazaicishu + "1"
Debug.Print ("���ش�����" & jiazaicishu)
'--------------------------------------------��������------------------------------------------
If WebBrowser1.Busy Then
Debug.Print ("��ҳδ�������")
        Exit Sub
    Else
    Debug.Print ("��ҳ�������")
Timer1.Enabled = False
WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
Dim vDoc, x, VTag
Set vDoc = WebBrowser1.Document
For x = 0 To vDoc.All.Length - 1 '������б�ǩ
If UCase(vDoc.All(x).tagName) = "INPUT" Then '�ҵ�input��ǩ
Set VTag = vDoc.All(x)
If VTag.Value = "�ύ" Then VTag.Click '����ύ�ˣ�һ�ж�OK��
End If
Next x
denglu = "yes"
Timer3.Enabled = True
Timer3.Interval = 1500
Debug.Print ("��¼״̬��" & "yes")
End If
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then  '������ǻس�������
Call Command1_Click
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
Debug.Print ("��ҳδ�������")
        Exit Sub
    Else
    Timer3.Enabled = False
Debug.Print ("��ʼ��ȡ����")
Dim szText As String
Dim szFindStrBegin As String
Dim szFindStrEnd As String
Dim nBegin As Long
Dim nEnd As Long
Dim nLength  As Long
szFindStrBegin = "����" '����Ҫ���ҵ��ַ�����ͷ
szFindStrEnd = "����" '����Ҫ���ҵ��ַ�����β

szText = WebBrowser1.Document.body.innerText '�õ��������֣���ʱ��ģ�壬ʵ��ʹ���л���ȥWebBrowser1.Document.body.innerText

nBegin = InStr(szText, szFindStrBegin) '�ҿ�ͷ�ַ���
If nBegin > 0 Then '���������ҵ���ͷ�˲ż���
    nEnd = InStr(nBegin, szText, szFindStrEnd) '�ҽ�β�ַ���
    If nEnd > nBegin Then '��β����ȿ�ͷ��λ�ô�
    
        '�������ҵ��ַ���ģʽ��ע�͵������2��
        nLength = nEnd - nBegin + Len(szFindStrEnd)   '������Ҫ��ȡ���ַ�������,���Ҫ�������ҵ��ַ�������1�У�ע������2��
        
        '���������ҵ��ַ���ģʽ
        nLength = nEnd - nBegin - Len(szFindStrBegin) '������������ҵ��ַ���������2��
        nBegin = nBegin + Len(szFindStrBegin) '������������ҵ��ַ���������2��
        
        szMyText = Mid(szText, nBegin, nLength)   'ȡ����before then.���� "test" �м�Ķ���
Debug.Print ("��ȡ���ݣ�" & szMyText)
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
Debug.Print ("��ʷ�ļ�������" & hang)
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
Debug.Print ("ת�����ݣ�" & lishi2)
lishi = lishi2 & Chr(13)
  Else
  End If
Label3.Caption = lishi
Text1.SetFocus
End If
Command1.Enabled = True
Command3.Enabled = True
End Sub

