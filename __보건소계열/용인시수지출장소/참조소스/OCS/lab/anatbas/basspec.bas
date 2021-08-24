Attribute VB_Name = "basSpec"
Option Explicit

'Public StrSql           As String
Public gStrUserID       As String
Public gStrUsername     As String
Public gStrPass         As String
Public gStrDept         As String
Public gStrRank         As String
Public gStrToisa        As String
Public gStrSlip         As String

Public gIntPassCnt      As Integer
Public gRefrash         As Integer
Public hWndReturn       As Long
Public gCallWin         As Integer
Public gSText           As String
Public gSOrgCall        As String
Public sOldGroupString  As String
Public sNewGroupString  As String

Public GstrSELECTOrderCode  As String
Public nIndex           As Integer

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long


Public Function Quot_Conv(ByVal sString As String) As Variant
    Dim sRecvStr
    Dim nStart      As Integer
    Dim sTemp       As String
    
    If Trim(Len(sString)) = "" Then Exit Function
    
    For nStart = 1 To Len(sString)
        sTemp = Mid(sString, nStart, 1)
        If Mid(sString, nStart, 1) = "'" Then
            sTemp = "''"
        ElseIf Mid(sString, nStart, 1) = """" Then
            sTemp = """"
        End If
        sRecvStr = sRecvStr & sTemp
    Next
    
    Quot_Conv = sRecvStr
    
End Function

Public Function SetComboBox(ByVal sCombo As Object, ByVal sCompString As String, Optional nLtCnt As Integer = 0) As Integer
    
    
    If Trim(sCompString) = "" Then
        sCombo.ListIndex = -1
        Exit Function
    End If
    
    SetComboBox = False
    
    If Val(nLtCnt) > 0 Then
        GoSub String_LeftCut_Sub
    Else
        GoSub String_Normal_Sub
    End If
    Exit Function
    
String_Normal_Sub:
    For i = 0 To sCombo.ListCount - 1
        If Trim(sCombo.List(i)) = Trim(sCompString) Then
            sCombo.ListIndex = i
            SetComboBox = True
            Exit For
        End If
    Next
    Return
    
String_LeftCut_Sub:
        
    nLtCnt = Len(Trim(sCompString))
    If nLtCnt = 0 Then
        sCombo.ListIndex = -1
        Exit Function
    End If
    
    For i = 0 To sCombo.ListCount - 1
        If Left(Trim(sCombo.List(i)), nLtCnt) = Trim(sCompString) Then
            sCombo.ListIndex = i
            SetComboBox = True
            Exit For
        End If
    Next
    Return
    
End Function

Public Function Spread_Set_Clear(ByVal sSpreadName As Object) As Integer
    
    sSpreadName.Row = 1
    sSpreadName.Col = 1
    sSpreadName.Row2 = sSpreadName.DataRowCnt
    sSpreadName.Col2 = sSpreadName.DataColCnt
    sSpreadName.BlockMode = True
    sSpreadName.Action = ActionClear
    sSpreadName.BlockMode = False
    
    sSpreadName.Row = 1
    sSpreadName.Col = 1
    sSpreadName.Action = ActionActiveCell
    
End Function
