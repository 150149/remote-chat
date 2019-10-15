Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public preWinproc As Long '��¼ԭ����windows�����ַ
Public Const WM_GETTEXT = &HD
Public Const WM_CLOSE = &H10
Public Const GWL_WNDPROC = (-4)


Public Function windproc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim h As Long '�����������鿴���������ľ���ı���
h = GetForegroundWindow '�õ������鿴���������ľ��,��Ϊ�����������һ�㶼��������˵�
If msg = WM_GETTEXT Then
  MsgBox "͵�����˵�������һ�ֺܲ��Ѻõ���Ϊ", vbExclamation + vbOKOnly, "���ؾ���"
  Call SendMessage(Form1.hwnd, WM_CLOSE, 0, 0)
  Call SendMessage(h, WM_CLOSE, 0, 0)
End If

windproc = CallWindowProc(preWinproc, hwnd, msg, wParam, lParam) '������һ����Ϣ
End Function
