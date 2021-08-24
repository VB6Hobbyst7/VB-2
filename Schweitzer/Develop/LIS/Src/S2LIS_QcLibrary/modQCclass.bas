Attribute VB_Name = "modQCclass"
Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'QC자동처방에서 내부정도관리 화면 띄워주는거 하면서 추가했음
'2003/10/13 Append By Legends

'엄마컨테이너 구하기
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'최상위 컨테이너 구하기
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
'폼 죽여버리는데 사용
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
'Setparent로 폼을 띄울경우에만 사용하는 함수
'Called 불림 당하는 폼
'Call 부르는 폼

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
'폼의 마지막인 경우 불린 폼의 엄마 폼을 닫기 위해서...
'폼이 불림당할 경우에만 사용해야 한다. 그렇지 않으면 Hwnd를 구하기 위해서 의미없는 Form Load가 시도됨

    If pCalled.ParentHwnd <> 0 Then
        Call SendMessage(pCalled.ParentHwnd, WM_CLOSE, 0&, 0&)
    End If
End Sub
