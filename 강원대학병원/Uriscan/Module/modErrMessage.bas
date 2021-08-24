Attribute VB_Name = "modErrMessage"
Option Explicit

'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private mDataBaseErrNo As Double
Private mDataBaseErrText As String

Private Sub AlwaysOnTop(F As Form, OnTop As Boolean)
    'hwndInsertAfter values
    Const HWND_TOP = 0
    Const HWND_BOTTOM = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    'wFlags values
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOREDRAW = &H8
    Const SWP_NOACTIVATE = &H10
    Const SWP_FRAMECHANGED = &H20           'The frame changed: send WM_NCCALCSIZE
    Const SWP_SHOWWINDOW = &H40
    Const SWP_HIDEWINDOW = &H80
    Const SWP_NOCOPYBITS = &H100
    Const SWP_NOOWNERZORDER = &H200    'Don't do owner Z ordering
    Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

    If OnTop = True Then
        'Turn on the TopMost attribute.
        SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    ElseIf OnTop = False Then
        'Turn off the TopMost attribute.
        SetWindowPos F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

Private Function GetWeek() As String
    Select Case vbUseSystemDayOfWeek
        Case 0
            GetWeek = "월요일"
        Case 1
            GetWeek = "화요일"
        Case 2
            GetWeek = "수요일"
        Case 3
            GetWeek = "목요일"
        Case 4
            GetWeek = "금요일"
        Case 5
            GetWeek = "토요일"
        Case 6
            GetWeek = "일요일"
    End Select
End Function

Public Sub ErrMsgProc(ByVal mMsg As String, Optional ByVal strMessage As String)
    
    Dim TempErrorText   As String
    Dim frmError        As Form
    
    TempErrorText = "        Procedure : " & mMsg & vbCrLf & vbCrLf & _
                    "     Error Source : " & Err.Source & vbCrLf & vbCrLf & _
                    "     Error Number : " & Err.Number & vbCrLf & vbCrLf & _
                    "Error Description : " & Err.Description & vbCrLf & vbCrLf & _
                    "             Date : " & Format(Now, "yyyy" & "년 " & "m" & "월 " & "d" & "일 " & GetWeek) & vbCrLf & vbCrLf & _
                    "             Time : " & Time
                                
    Set frmError = New frmErrMessage
    
    frmError.Text_View = TempErrorText
    frmError.Show vbModal

    Set frmError = Nothing
    
End Sub

'sql server error setting
Public Sub DBErrorSet(ByVal AdoCn As ADODB.Connection, Optional ByVal strSql As String = "", Optional mMsg As String = "")
    Dim errLoop As ADODB.Error
    Dim aryError(5) As String
    Dim i As Integer
    Dim aryLen(5) As Double
    Dim dblLen As Double
    
    Call DBErrClear
    For Each errLoop In AdoCn.Errors
        mDataBaseErrNo = errLoop.NativeError
        mDataBaseErrText = "    Procedure : " & mMsg & vbCrLf & vbCrLf & _
                           " Error Number : " & errLoop.Number & vbCrLf & vbCrLf & _
                           "  Description : " & errLoop.Description & vbCrLf & vbCrLf & _
                           "       Source : " & errLoop.Source & vbCrLf & vbCrLf & _
                           "    SQL State : " & errLoop.SQLState & vbCrLf & vbCrLf & _
                           " Native Error : " & errLoop.NativeError & vbCrLf & vbCrLf & _
                           "   SQL String : " & strSql & vbCrLf & vbCrLf & _
                           "         Date : " & Format(Now, "yyyy" & "년 " & "m" & "월 " & "d" & "일 " & GetWeek) & vbCrLf & vbCrLf & _
                           "         Time : " & Time
    Next
    
    frmErrMessage.Text_View = mDataBaseErrText
    frmErrMessage.Show vbModal, MainForm
End Sub

'db error clear
Public Sub DBErrClear()
    mDataBaseErrNo = 0
    mDataBaseErrText = ""
End Sub

'data base error: read only property
Public Function DataBaseErrNo() As Double
    DataBaseErrNo = mDataBaseErrNo
End Function

'data base error text: read only property
Public Function DataBaseErrText() As String
    DataBaseErrText = mDataBaseErrText
End Function

