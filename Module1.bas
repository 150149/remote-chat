Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public preWinproc As Long '记录原来的windows程序地址
Public Const WM_GETTEXT = &HD
Public Const WM_CLOSE = &H10
Public Const GWL_WNDPROC = (-4)


Public Function windproc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim h As Long '定义存放用来查看密码的软件的句柄的变量
h = GetForegroundWindow '得到用来查看密码的软件的句柄,因为这类软件窗口一般都是置于最顶端的
If msg = WM_GETTEXT Then
  MsgBox "偷看别人的密码是一种很不友好的行为", vbExclamation + vbOKOnly, "严重警告"
  Call SendMessage(Form1.hwnd, WM_CLOSE, 0, 0)
  Call SendMessage(h, WM_CLOSE, 0, 0)
End If

windproc = CallWindowProc(preWinproc, hwnd, msg, wParam, lParam) '接收下一条消息
End Function
