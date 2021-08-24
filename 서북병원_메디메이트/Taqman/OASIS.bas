Attribute VB_Name = "modOASIS"
Option Explicit

Public ADR_HIS      As ADODB.Recordset
Public ADE_HIS      As ADODB.Error
Public ARC_HIS      As Double 'Fetch Record Count

Public Sub ErrSQL_HIS(ArgSQL As String)
    Beep
    Beep
    Beep
            
    For Each ADE_HIS In cn_Ser.Errors
        MsgBox "오류코드 - " & ADE_HIS.Number & vbCrLf & _
               "오류소스 - " & ADE_HIS.Source & vbCrLf & _
               "오류내용 - " & ADE_HIS.Description & vbCrLf & _
               "SQL 문장 - " & ArgSQL _
               , vbExclamation, "데이타작업중 오류가 발생했습니다."
    Next ADE_HIS
End Sub

Public Function ReadSQL_HIS(ArgSQL$, ArgARS As ADODB.Recordset) As Boolean
    Screen.MousePointer = 11
    ReadSQL_HIS = True
On Error GoTo ADO_ERR
    Set ArgARS = New ADODB.Recordset
    
    ArgARS.Open ArgSQL, cn_Ser, adOpenForwardOnly, adLockReadOnly

On Error GoTo 0
    If ArgARS.EOF Then
        ARC_HIS = 0
        ArgARS.Close
        Set ArgARS = Nothing
        Screen.MousePointer = 0
'''        ReadSQL_HIS = False
        Exit Function
    End If
    ARC_HIS = ArgARS.RecordCount
    ArgARS.MoveFirst
    Screen.MousePointer = 0
Exit Function

'/--------------------------------------------------------------------------------------------------------------------------------------------------------------/

ADO_ERR:
    ReadSQL_HIS = False
    ARC_HIS = 0
    Call ErrSQL_HIS(ArgSQL)
    Screen.MousePointer = 0
End Function


