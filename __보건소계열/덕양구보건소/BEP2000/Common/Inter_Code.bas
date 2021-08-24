Attribute VB_Name = "Module8"
'
'   최종 수정일 : 1997.11.26  yk
'
Option Explicit

    Global Chk_LogID    As Integer
    
    Global UserID   As String
    Global UserNM   As String
    
    Global ListRow  As Integer

    '--- For Add_Server_Result(환자 신상자료 저장)
    Type SAMPLE_INFO
        LabID       As String       'Lab_ID
        CstIDNo     As String       '진찰권번호
        RtnCd       As String
        reclabno    As String
        age         As String
        LabTime     As String
        OrdID       As String
        RecChk      As String
        OrdStat     As String
    End Type
    Global Sam_Info As SAMPLE_INFO

'    Type PRE_RESULT
'        LabDate     As String
'        SlipCd      As String
'        LabSqNo     As String
'        OrdCd       As String
'        SubSqNo     As String
'        RstDate     As String
'        RstVal      As String
'        SysTime     As String
'        CstIDNo     As String
'    End Type
'    Global Pre_Res  As PRE_RESULT
    
    '--- For Insert_Server(경북대 Version)
    Type INSERT_SERVER_RESULT
        ordcd       As String
        SubNo       As String
        Result      As String
        Ref         As String
        RtnCd       As String
    End Type
    Global Insert_Server(1 To 90) As INSERT_SERVER_RESULT       '11/20 yk
    
    Global iResCnt  As Integer     '결과 등록 Count
    
    '--- For Update Check
    Global Chk_Exist    As Integer  '기존 결과 존재하는지 Check
    
    '--- For Common Form(frmQuery)
    Global sEq_Name As String       '장비명(예:AxSYM => Ax)
    
        
Public Function Get_RtnCd(sLabNo As String, sOrdCd As String, sSubNo As String, sOrderNo As String) As String

    Dim sStr    As String
    Dim sData() As String
    Dim iRet_Cd As Integer

    Get_RtnCd = ""
    
    sStr = " Select RSLIPCD + RORDCD + RSPCCD from LAB_DB..LAB030M " _
            & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
            & "   and NUMGBN  = '" & Mid(sLabNo, 9, 1) & "'" _
            & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
            & "   and SLIPCD = '" & Mid(sOrdCd, 1, 2) & "'" _
            & "   and ORDCD = '" & Mid(sOrdCd, 3, 3) & "'" _
            & "   and SPCCD = '" & Mid(sOrdCd, 6, 2) & "'" _
            & "   and SUBCD = '" & sSubNo & "'" _
            & "   and ORDDATE + DEPTCD + SEQNO = '" & sOrderNo & "'"
            
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 1, sStr, sData()
            
            Get_RtnCd = sData(1)
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
        
End Function
