Attribute VB_Name = "modQCclass"
Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'QC�ڵ�ó�濡�� ������������ ȭ�� ����ִ°� �ϸ鼭 �߰�����
'2003/10/13 Append By Legends

'���������̳� ���ϱ�
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'�ֻ��� �����̳� ���ϱ�
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
'�� �׿������µ� ���
Public Const WM_CLOSE = &H10
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function AllowSetForegroundWindow Lib "user32.dll" (ByVal dwProcessId As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Global gIsDeveloper As Boolean
Global gBuildingCd As String
Global gEmpId As String

Public Function IsLastForm() As Boolean

    Dim i As Long
    Dim tmpFrm As Form
    
    i = 0
    IsLastForm = False
    
    For Each tmpFrm In Forms
        i = i + 1
    Next
    If i = 0 Then IsLastForm = True

End Function

Public Sub LoadForm(ByRef pCalled As Form, ByRef pCall As Form)
'Coding By legends 2003/11/17
'Setparent�� ���� ����쿡�� ����ϴ� �Լ�
'Called �Ҹ� ���ϴ� ��
'Call �θ��� ��

'    Dim frm As Form
'    Dim frmExist As Boolean
    
    pCalled.ParentHwnd = GetAncestor(pCall.hwnd, 1)
    
'    frmExist = False
'    For Each frm In Forms
'        If frm.Name = pChild.Name Then
'
'            frmExist = True
'        End If
'    Next
    
'    DoEvents
'    If frmExist = False Then
        Call SetParent(pCalled.hwnd, pCalled.ParentHwnd)
        pCalled.WindowState = 2
        pCalled.Show
'    End If
    pCalled.ZOrder 0
End Sub

Public Sub UnloadForm(ByRef pCalled As Form)
'Coding By legends 2003/11/17
'���� �������� ��� �Ҹ� ���� ���� ���� �ݱ� ���ؼ�...
'���� �Ҹ����� ��쿡�� ����ؾ� �Ѵ�. �׷��� ������ Hwnd�� ���ϱ� ���ؼ� �ǹ̾��� Form Load�� �õ���

    If pCalled.ParentHwnd <> 0 Then
        Call SendMessage(pCalled.ParentHwnd, WM_CLOSE, 0&, 0&)
    End If
End Sub
