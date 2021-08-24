Attribute VB_Name = "modBBSLibrary"
Option Explicit

Public Function GetCom003RecordSet(ByVal CDINDEX As String, Optional cdval1 As String = "", _
                                                       Optional expfg As Integer = 0) As Recordset
    '--------------------------------------------------------------------
    'expfg = 0 : 폐기된것 제외
    '      = 1 : 폐기된것 포함
    '--------------------------------------------------------------------
    Dim sSql As String
    
    sSql = "SELECT * FROM " & T_COM003 & " " & _
           "WHERE " & DBW("cdindex=", CDINDEX) & " "
           
    If cdval1 <> "" Then sSql = sSql & "AND " & DBW("cdval1=", cdval1) & " "
    If expfg = 0 Then sSql = sSql & "AND (field5 is null or field5='') "
    
    sSql = sSql & "ORDER BY cdval1 "

On Error GoTo OpenRecordSeq_error
    
    
    Set GetCom003RecordSet = New Recordset
    GetCom003RecordSet.Open sSql, dbconn
    
'    If GetCom003RecordSet.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetCom003RecordSet = Nothing
'    End If

    Exit Function
    
OpenRecordSeq_error:
    MsgBox Err.Description, vbCritical, "오류"
    Set GetCom003RecordSet = Nothing
End Function

Public Function OpenRecordSetDay(ByVal CDINDEX As String, Optional ByVal cdval1 As String = "") As Recordset
    Dim sSql As String
    
    If cdval1 = "" Then cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    sSql = "SELECT * FROM " & T_COM003 & " " & _
           "WHERE " & DBW("cdindex=", CDINDEX) & " " & _
           "AND cdval1=(" & _
                        "SELECT max(cdval1) " & _
                        "FROM " & T_COM003 & " " & _
                        "WHERE " & DBW("cdindex=", CDINDEX) & " " & _
                        "AND " & DBW("cdval1<=", cdval1) & " " & _
                        ") "

On Error GoTo OpenRecordSeq_error
    
    Set OpenRecordSetDay = New Recordset
    
    Call OpenRecordSetDay.Open(sSql, dbconn)
'    If OpenRecordSetDay.DBerror = True Then
'        dbconn.DisplayErrors
'        Set OpenRecordSetDay = Nothing
'    End If

    Exit Function
    
OpenRecordSeq_error:
    MsgBox Err.Description, vbCritical, "오류"
    Set OpenRecordSetDay = Nothing
End Function

