VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "��¼����"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5295
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "       ��¼       "
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   2160
   End
   Begin VB.Label Label3 
      Caption         =   "�ʺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C000&
      X1              =   0
      X2              =   5280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C000&
      X1              =   5280
      X2              =   5280
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      X1              =   0
      X2              =   5280
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oShadow As New aShadow
Public jianrong As String
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "�ó���������"
End
End If

If Dir(App.Path & "\xxx") = "" Then
Open App.Path & "\xxx" For Output As #11
Print #11, "�ο�"
Close #11
End If
If Dir(App.Path & "\idcard") = "" Then
Open App.Path & "\idcard" For Output As #11
Print #11, ""
Close #11
End If
With oShadow
    If .Shadow(Me) Then
        .Depth = 20 '��Ӱ���
        .Color = RGB(0, 0, 0) '��Ӱ��ɫ
        .Transparency = 50 '��Ӱɫ��
    End If
 End With
 If Dir("c:\windows\system32\Comdlg32.ocx") = "" Then
MsgBox "ȱ������������ü���ģʽ"
jianrong = "true"
Else
jianrong = "false"
End If
preWinproc = GetWindowLong(Text2.hwnd, GWL_WNDPROC) '��¼ԭ���Ĵ��ڳ���ĵ�ַ
SetWindowLong Text2.hwnd, GWL_WNDPROC, AddressOf windproc '�����������Ϣ����ԭ��ַ
End Sub

Private Sub Label3_dblClick()
If Dir("c:\windows\system32\Comdlg32.ocx") = "" Then
MsgBox "ȱ������������ü���ģʽ"
jianrong = "true"
Else
MsgBox "�ѹرռ���ģʽ"
jianrong = "false"
End If
End Sub

Private Sub Label5_Click()
Dim idcard, shuru, mingzi As String
mingzi = Text1.Text
Debug.Print ("���֣�" & mingzi)
Open App.Path & "\xxx" For Output As #2
Print #2, mingzi
Close #2
Open App.Path & "\idcard" For Input As #1
Line Input #1, idcard
Debug.Print ("���룺" & idcard)
Close #1
shuru = Text2.Text
If idcard = "" Then
Open App.Path & "\idcard" For Output As #11
Print #11, shuru
Close #11
Debug.Print ("ע�᣺" & shuru)
If jianrong = "true" Then
Form3.Show
Unload Me
Else
Form2.Show
Unload Me
End If
ElseIf shuru = idcard Then
Form2.Show
Form1.Hide
ElseIf Text1.Text = "" Then
Form2.Show
Form1.Hide
Else
MsgBox "�������"
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

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong Text2.hwnd, GWL_WNDPROC, preWinproc 'ȡ����Ϣ�Ľ�ȡ,ʹ֮����ԭ����windows����
End Sub
