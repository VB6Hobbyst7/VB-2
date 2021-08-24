Attribute VB_Name = "modLabComments"
Option Explicit

Global objDoctor As New clsDoctor
Global objLabComments As New clsLabComments

Global gDoctId As String
Global gDoctNm As String

Global gPatId As String
Global gPatNm As String

Global gPtntId As String
Global gBedInDT As String
Global sBedInDT As String

Global Const iMsgTop1 = 75
Global Const iMsgTop2 = 195

Global fMain As MDIForm

Sub Main()

End Sub

Public Sub UnlockPtnt(ByVal pPtId As String, ByVal pBedDt As String)

    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = "select * from " & T_LAB501 & " where " & DBW("ptid = ", pPtId) & " and " & DBW("bedindt = ", pBedDt)
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Not Rs.EOF Then
        If Trim("" & Rs.Fields("DoneFg").Value) = "0" Then   '저장없이 선택만 했을경우...
            SqlStmt = " update  " & T_LAB501 & " set donefg = '', rptid = null " & _
                      " where " & DBW("ptid = ", pPtId) & " and " & DBW("bedindt = ", pBedDt)
            
            On Error GoTo Err_Trap
            
            DBConn.BeginTrans
            DBConn.Execute SqlStmt
            DBConn.CommitTrans
        End If
    End If
    
'    Rs.RsClose
    Set Rs = Nothing
    Exit Sub

Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Public Function CheckStatus(ByVal pPtId As String, ByVal pBedDt As String, pRptId As String) As Boolean

    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = "select * from " & T_LAB501 & " where " & DBW("ptid = ", pPtId) & " and " & DBW("bedindt = ", pBedDt)
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
'    MsgBox SqlStmt, vbExclamation
    
    CheckStatus = False
    
    If Not Rs.EOF Then
        If Trim("" & Rs.Fields("DoneFg").Value) = "" Then
            CheckStatus = True
            SqlStmt = " update  " & T_LAB501 & " set donefg = '0', " & DBW("rptid = ", pRptId) & _
                      " where   " & DBW("ptid = ", pPtId) & " and " & DBW("bedindt = ", pBedDt)
            
'            MsgBox SqlStmt, vbExclamation
            
            On Error GoTo Err_Trap
            
            DBConn.BeginTrans
            DBConn.Execute SqlStmt
            DBConn.CommitTrans
        Else
            If Trim("" & Rs.Fields("Rptid").Value) = objDoctor.DoctId Then CheckStatus = True
        End If
    Else
        CheckStatus = True
                    
        SqlStmt = " insert into  " & T_LAB501 & _
                  " (ptid, bedindt, disease, wardid, hosilid, deptcd, complain, history, reqdt, donefg, " & _
                  "  prtfg, rptdt, rpttm, rptid, mfydt, mfytm, mfyid, ptdiv, bedintm, bedoutdt, bedouttm, majdoct) " & _
                  " select a." & F_INPTID & ", " & F_BEDINDT2("a") & ", b.icd, a." & F_PTWARDID & ", " & _
                  "        a." & F_PTROOMID & ", a." & F_PTDEPTCD & ", null, null, null, '0', " & _
                  "        null, null, null, " & DBV("rptid", pRptId) & ", null, null, null, a." & F_PTDIV & ",    to_char(a.admtime,'hh24miss') " & ", " & _
                           F_BEDOUTDT2("a") & ", a." & F_BEDOUTTM & ", a." & F_MAJDOCT & _
                  " from   " & T_LAB106 & " b, " & T_HIS002 & " a " & _
                  " where  " & DBW("a." & F_PTID, pPtId, 2) & _
                  " and  " & F_BEDINDT2("a") & " = to_date('" & pBedDt & "','yyyymmdd') " & _
                  " and    " & DBJ("b.ptid    =* a." & F_PTID) & _
                  " and    " & DBJ("b.bedindt =* " & F_BEDINDT2("a")) & _
                  " and    " & DBJ("b.seq =* 1 ")
'                  " and    a.ip_set6 ='0' " & _

'        MsgBox SqlStmt, vbExclamation

        On Error GoTo Err_Trap
        
        DBConn.BeginTrans
        DBConn.Execute SqlStmt
        DBConn.CommitTrans
    End If
    
'    Rs.RsClose
    Set Rs = Nothing
    Exit Function

Err_Trap:
    DBConn.RollbackTrans
    CheckStatus = False
'    Call Error_Routine
    MsgBox Err.Description, vbExclamation
End Function
'
Public Function GetConnInfo() As String
    Dim strDB As String
    Dim strUID As String
    Dim strPWD As String
'DB

    strDB = GetSetting("Schweitzer2000 LIS", "Server", "DB", "")
    strUID = GetSetting("Schweitzer2000 LIS", "Server", "UID", "")
    strPWD = GetSetting("Schweitzer2000 LIS", "Server", "PWD", "")

    GetConnInfo = strDB & ";" & strUID & ";" & strPWD

End Function

