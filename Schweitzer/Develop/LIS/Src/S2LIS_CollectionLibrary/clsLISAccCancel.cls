VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISAccCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objAccSql As New clsLISSqlAccession
Private TblNames(0 To 14)

Public Function DoCancelAccession(ByVal pPtid As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pStsCd As String, _
                                  ByVal pEmpId As String, ByVal pReason As String, _
                                  ByVal objordSheet As Object, ByVal pFlag As Long) As Boolean
    Dim i As Long
    Dim SqlStmt As String
    Dim Resp  As Boolean
    
    TblNames(0) = T_LAB203
    TblNames(1) = T_LAB205
    TblNames(2) = T_LAB302
    TblNames(3) = T_LAB303
    TblNames(4) = T_LAB304
    TblNames(5) = T_LAB305
    TblNames(6) = T_LAB308
    TblNames(7) = T_LAB351
    TblNames(8) = T_LAB353
    TblNames(9) = T_LAB354
    TblNames(10) = T_LAB404
    TblNames(11) = T_LAB405
    TblNames(12) = T_LAB407
    TblNames(13) = T_LAB360
    TblNames(14) = T_LAB361
   
    Dim sOrdDt As String, sOrdNo As String, sOrdSeq As String, sFlag As Long
    Dim objOrdDic As New clsDictionary
   
    On Error GoTo Err_Trap
    
    
    '������� Delete
    For i = 0 To 14
        If IsDataExists(TblNames(i), pWorkArea, pAccDt, pAccSeq) Then
            SqlStmt = objAccSql.SqlDelRstTable(TblNames(i), pWorkArea, pAccDt, pAccSeq)
             DBConn.Execute (SqlStmt)
        End If
    Next
    
'-------------------------------------------------------------------------------------------
'��ũ��Ʈ ���� ����
'Append By Legends 2003/10/15
'Ȥ�ó� ������ҽ� ������ �߻��ϸ� �Ʒ��� �ִ� ������ ����ũ ó���Ͻÿ�.
'������ Ŭ������ �޼ҵ�� ����ũ ó������ ���ÿ�. �ֳĸ� ȣȯ���� Ȧ��� ������ �Ǵϱ�....

    Dim TblNamesAdd(1 To 4) As String
    
    TblNamesAdd(1) = T_LAB301   '�Ϲݰ˻� ��ũ��Ʈ ����
    TblNamesAdd(2) = T_LAB401   '�̻��� ��ũ��Ʈ ���� �Ӹ�
    TblNamesAdd(3) = T_LAB402   '�̻��� ��ũ��Ʈ ���� ���׾Ƹ�
    TblNamesAdd(4) = T_LAB403   '�̻��� ��ũ��Ʈ ���� ����(�̰� �������� �ɼ��� ����)

    If IsDataExists(TblNamesAdd(1), pWorkArea, pAccDt, pAccSeq) Then
        SqlStmt = objAccSql.SqlDelRstTable(TblNamesAdd(1), pWorkArea, pAccDt, pAccSeq)
         DBConn.Execute (SqlStmt)
    End If
    
    If IsDataExists(TblNamesAdd(4), pWorkArea, pAccDt, pAccSeq) Then
        SqlStmt = objAccSql.SqlDelRstTable(TblNamesAdd(4), pWorkArea, pAccDt, pAccSeq)
          DBConn.Execute (SqlStmt)
    End If
    
'    Dim strWsCd As String
'    Dim strWsUnit As String

'    If IsDataExistsForMic(pWorkarea, pAccDt, pAccSeq, strWsCd, strWsUnit) Then
'        '401���� �����
'        SqlStmt = objAccSql.SqlDelRstTableForMic(TblNamesAdd(2), strWsCd, strWsUnit)
'        DBConn.Execute (SqlStmt)
'
'        '402�����
'        SqlStmt = objAccSql.SqlDelRstTableForMic(TblNamesAdd(3), pWorkarea, pAccDt, pAccSeq)
'        DBConn.Execute (SqlStmt)
'    End If
    
    '## �̻��� WorkSheet �������� - �̻��(2004-12-08)
    '   - �̻��� ������Ұ� �ƴϾ �׻� Ÿ�¹����� ����!
    Call DelMicWorksheet(pWorkArea, pAccDt, pAccSeq)
'-------------------------------------------------------------------------------------------

    '�������� Update : ó����·�.. status�� 'D'(���), 'ä�����·�.. status�� '1'(ä��)
    SqlStmt = objAccSql.SqlCancel201(pWorkArea, pAccDt, pAccSeq, pStsCd)
    Call DBConn.Execute(SqlStmt)
    
    '������� ���� Insert
    SqlStmt = objAccSql.SqlInsertReason_New(pWorkArea, pAccDt, pAccSeq, pEmpId, pReason) '-- �������/��ҽð� �߰� ���� 2016.04.14 �½�ȣ
    Call DBConn.Execute(SqlStmt)
        
    objOrdDic.Clear
    objOrdDic.FieldInialize "orddt,ordno", "updfg"
    
    With objordSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1: sOrdDt = Format(.Value, CS_DateDbFormat)  'ó����
            .Col = 2: sOrdNo = .Value   'ó���ȣ
            .Col = 7: sOrdSeq = .Value  'ó��Seq
            
            'ó��Body Update
            SqlStmt = objAccSql.SqlUpdateOrdB(pPtid, sOrdDt, sOrdNo, sOrdSeq, pFlag)
            Call DBConn.Execute(SqlStmt)
            
            'ó��Header Update
            If Not objOrdDic.Exists(sOrdDt & COL_DIV & sOrdNo) Then
                SqlStmt = objAccSql.SqlUpdateOrdH(pPtid, sOrdDt, sOrdNo, pFlag)
                Call DBConn.Execute(SqlStmt)
                
                objOrdDic.AddNew sOrdDt & COL_DIV & sOrdNo, "Y"
            End If
            
            '** �������� �߰� By M.G.Choi 2005.08.25
                '   ��� �� OCS������� ����ó�� (mdresult delete)
            SqlStmt = objAccSql.OCSResultDelete(pPtid, sOrdDt, sOrdNo, sOrdSeq, pFlag)
            
            Call DBConn.Execute(SqlStmt)
            
        Next
    End With
    DoCancelAccession = True
    
'    Set rs = Nothing
    
    Exit Function
    
Err_Trap:
    DoCancelAccession = False

End Function

'POCT �ϰ�ä����� �߰�
'�½�ȣ 2012-05-30
Public Function DoCancelAccession_POCT(ByVal pPtid As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, ByVal pStsCd As String, _
                                  ByVal pEmpId As String, ByVal pReason As String, _
                                  ByVal sOrdDt As String, ByVal sOrdNo As String, _
                                  ByVal sOrdSeq As String, ByVal pFlag As Long) As Boolean
    Dim i As Long
    Dim SqlStmt As String
    Dim Resp  As Boolean
    
    TblNames(0) = T_LAB203
    TblNames(1) = T_LAB205
    TblNames(2) = T_LAB302
    TblNames(3) = T_LAB303
    TblNames(4) = T_LAB304
    TblNames(5) = T_LAB305
    TblNames(6) = T_LAB308
    TblNames(7) = T_LAB351
    TblNames(8) = T_LAB353
    TblNames(9) = T_LAB354
    TblNames(10) = T_LAB404
    TblNames(11) = T_LAB405
    TblNames(12) = T_LAB407
    TblNames(13) = T_LAB360
    TblNames(14) = T_LAB361
   
'    Dim sOrdDt As String, sOrdNo As String, sOrdSeq As String, sFlag As Long
    Dim objOrdDic As New clsDictionary
   
    On Error GoTo Err_Trap
    
    
    '������� Delete
    For i = 0 To 14
        If IsDataExists(TblNames(i), pWorkArea, pAccDt, pAccSeq) Then
            SqlStmt = objAccSql.SqlDelRstTable(TblNames(i), pWorkArea, pAccDt, pAccSeq)
             DBConn.Execute (SqlStmt)
        End If
    Next
    
'-------------------------------------------------------------------------------------------
'��ũ��Ʈ ���� ����
'Append By Legends 2003/10/15
'Ȥ�ó� ������ҽ� ������ �߻��ϸ� �Ʒ��� �ִ� ������ ����ũ ó���Ͻÿ�.
'������ Ŭ������ �޼ҵ�� ����ũ ó������ ���ÿ�. �ֳĸ� ȣȯ���� Ȧ��� ������ �Ǵϱ�....

    Dim TblNamesAdd(1 To 4) As String
    
    TblNamesAdd(1) = T_LAB301   '�Ϲݰ˻� ��ũ��Ʈ ����
    TblNamesAdd(2) = T_LAB401   '�̻��� ��ũ��Ʈ ���� �Ӹ�
    TblNamesAdd(3) = T_LAB402   '�̻��� ��ũ��Ʈ ���� ���׾Ƹ�
    TblNamesAdd(4) = T_LAB403   '�̻��� ��ũ��Ʈ ���� ����(�̰� �������� �ɼ��� ����)

    If IsDataExists(TblNamesAdd(1), pWorkArea, pAccDt, pAccSeq) Then
        SqlStmt = objAccSql.SqlDelRstTable(TblNamesAdd(1), pWorkArea, pAccDt, pAccSeq)
         DBConn.Execute (SqlStmt)
    End If
    
    If IsDataExists(TblNamesAdd(4), pWorkArea, pAccDt, pAccSeq) Then
        SqlStmt = objAccSql.SqlDelRstTable(TblNamesAdd(4), pWorkArea, pAccDt, pAccSeq)
          DBConn.Execute (SqlStmt)
    End If
    
'    Dim strWsCd As String
'    Dim strWsUnit As String

'    If IsDataExistsForMic(pWorkarea, pAccDt, pAccSeq, strWsCd, strWsUnit) Then
'        '401���� �����
'        SqlStmt = objAccSql.SqlDelRstTableForMic(TblNamesAdd(2), strWsCd, strWsUnit)
'        DBConn.Execute (SqlStmt)
'
'        '402�����
'        SqlStmt = objAccSql.SqlDelRstTableForMic(TblNamesAdd(3), pWorkarea, pAccDt, pAccSeq)
'        DBConn.Execute (SqlStmt)
'    End If
    
    '## �̻��� WorkSheet �������� - �̻��(2004-12-08)
    '   - �̻��� ������Ұ� �ƴϾ �׻� Ÿ�¹����� ����!
    Call DelMicWorksheet(pWorkArea, pAccDt, pAccSeq)
'-------------------------------------------------------------------------------------------

    '�������� Update : ó����·�.. status�� 'D'(���), 'ä�����·�.. status�� '1'(ä��)
    SqlStmt = objAccSql.SqlCancel201(pWorkArea, pAccDt, pAccSeq, pStsCd)
    Call DBConn.Execute(SqlStmt)
    
    '������� ���� Insert
    SqlStmt = objAccSql.SqlInsertReason(pWorkArea, pAccDt, pAccSeq, pEmpId, pReason)
    Call DBConn.Execute(SqlStmt)
        
    objOrdDic.Clear
    objOrdDic.FieldInialize "orddt,ordno", "updfg"
               
    'ó��Body Update
    SqlStmt = objAccSql.SqlUpdateOrdB(pPtid, sOrdDt, sOrdNo, sOrdSeq, pFlag)
    Call DBConn.Execute(SqlStmt)
    
    'ó��Header Update
    If Not objOrdDic.Exists(sOrdDt & COL_DIV & sOrdNo) Then
        SqlStmt = objAccSql.SqlUpdateOrdH(pPtid, sOrdDt, sOrdNo, pFlag)
        Call DBConn.Execute(SqlStmt)
        
        objOrdDic.AddNew sOrdDt & COL_DIV & sOrdNo, "Y"
    End If
    
    '** �������� �߰� By M.G.Choi 2005.08.25
        '   ��� �� OCS������� ����ó�� (mdresult delete)
    SqlStmt = objAccSql.OCSResultDelete(pPtid, sOrdDt, sOrdNo, sOrdSeq, pFlag)
    
    Call DBConn.Execute(SqlStmt)
            
    DoCancelAccession_POCT = True
    
'    Set rs = Nothing
    
    Exit Function
    
Err_Trap:
    DoCancelAccession_POCT = False

End Function



Public Function DoCancelPart(ByVal pPtid As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pStsCd As String, _
                             ByVal pEmpId As String, ByVal pReason As String, _
                             ByVal objordSheet As Object, ByVal pFlag As Long) As Boolean
    Dim i As Long, j As Long
    Dim SqlStmt As String
    Dim Resp  As Boolean
    Dim lngTotCnt  As Long, lngInputCnt As Long
    Dim objRs As Recordset
    
    TblNames(0) = T_LAB205
    TblNames(1) = T_LAB305
    TblNames(2) = T_LAB302
    TblNames(3) = T_LAB351
    TblNames(4) = T_LAB404
    TblNames(5) = T_LAB405
    TblNames(6) = T_LAB407
    TblNames(7) = T_LAB360
    TblNames(8) = T_LAB361
   
    Dim sOrdDt As String, sOrdNo As String, sOrdSeq As String, sFlag As Long
    Dim objOrdDic As New clsDictionary
   
    On Error GoTo Err_Trap
    
    With objordSheet
        For j = 1 To .DataRowCnt
            .Row = j
            .Col = 8
            If .Value = 1 Then
                .Col = 1: sOrdDt = Format(.Value, CS_DateDbFormat)  'ó����
                .Col = 2: sOrdNo = .Value   'ó���ȣ
                .Col = 7: sOrdSeq = .Value  'ó��Seq
            
                '## �̻��� WorkSheet �������� - �̻��(2004-12-13)
                '   - �̻��� ������Ұ� �ƴϾ �׻� Ÿ�¹����� ����!
                If pStsCd = enStsCd.StsCd_LIS_Cancel Then
                    Call DelMicWorksheetPart(pWorkArea, pAccDt, pAccSeq, pPtid, sOrdDt, _
                         sOrdNo, sOrdSeq)
                End If
            
                '������� Delete
                For i = 0 To 8
                    If IsPartDataExists(TblNames(i), pWorkArea, pAccDt, pAccSeq, pPtid, sOrdDt, sOrdNo, sOrdSeq) Then
                        SqlStmt = objAccSql.SqlDelRstPart(TblNames(i), pWorkArea, pAccDt, pAccSeq, pPtid, sOrdDt, sOrdNo, sOrdSeq)
                        If SqlStmt <> "" Then DBConn.Execute (SqlStmt)
                    End If
                Next
            End If
        Next
    End With
    
'-------------------------------------------------------------------------------------------
'��ũ��Ʈ ���� ����
'Append By Legends 2003/10/15
'Ȥ�ó� ������ҽ� ������ �߻��ϸ� �Ʒ��� �ִ� ������ ����ũ ó���Ͻÿ�.
'������ Ŭ������ �޼ҵ�� ����ũ ó������ ���ÿ�. �ֳĸ� ȣȯ���� Ȧ��� ������ �Ǵϱ�....
    If pStsCd = enStsCd.StsCd_LIS_Cancel Then    'ó����·� ����� ��� �� ��������..
        Dim TblNamesAdd(1 To 4) As String

        TblNamesAdd(1) = T_LAB301   '�Ϲݰ˻� ��ũ��Ʈ ����
        TblNamesAdd(2) = T_LAB401   '�̻��� ��ũ��Ʈ ���� �Ӹ�
        TblNamesAdd(3) = T_LAB402   '�̻��� ��ũ��Ʈ ���� ���׾Ƹ�
        TblNamesAdd(4) = T_LAB403   '�̻��� ��ũ��Ʈ ���� ����(�̰� �������� �ɼ��� ����)

        If IsDataExists(TblNamesAdd(1), pWorkArea, pAccDt, pAccSeq) Then
            SqlStmt = objAccSql.SqlDelRstTable(TblNamesAdd(1), pWorkArea, pAccDt, pAccSeq)
             DBConn.Execute (SqlStmt)
        End If

        If IsDataExists(TblNamesAdd(4), pWorkArea, pAccDt, pAccSeq) Then
            SqlStmt = objAccSql.SqlDelRstTable(TblNamesAdd(4), pWorkArea, pAccDt, pAccSeq)
             DBConn.Execute (SqlStmt)
        End If

'        Dim strWsCd As String
'        Dim strWsUnit As String
'
'        If IsDataExistsForMic(pWorkArea, pAccDt, pAccSeq, strWsCd, strWsUnit) Then
'            '401���� �����
'                SqlStmt = objAccSql.SqlDelRstTableForMic(TblNamesAdd(2), strWsCd, strWsUnit)
'                 DBConn.Execute (SqlStmt)
'
'            '402�����
'                SqlStmt = objAccSql.SqlDelRstTableForMic(TblNamesAdd(3), pWorkArea, pAccDt, pAccSeq)
'                 DBConn.Execute (SqlStmt)
'        End If
    End If
'-------------------------------------------------------------------------------------------

    '�������� Update : ó����·�.. status�� 'D'(���), 'ä�����·�.. status�� '1'(ä��)
    lngTotCnt = 0
    Set objRs = New Recordset
    objRs.Open objAccSql.SqlGetTotCnt(pWorkArea, pAccDt, pAccSeq), DBConn
    
    While (Not objRs.EOF)
        lngTotCnt = lngTotCnt + Val(objRs.Fields("TotCnt").Value)
        objRs.MoveNext
    Wend
    Set objRs = Nothing
    
    lngInputCnt = 0
    Set objRs = New Recordset
    objRs.Open objAccSql.SqlGetInputCnt(pWorkArea, pAccDt, pAccSeq), DBConn
    
    While (Not objRs.EOF)
        lngInputCnt = lngInputCnt + Val(objRs.Fields("TotCnt").Value)
        objRs.MoveNext
    Wend
    Set objRs = Nothing
                                         
    If lngInputCnt = lngTotCnt Then '����� ��� �Է�
        SqlStmt = objAccSql.SqlUpdate201(pWorkArea, pAccDt, pAccSeq, enStsCd.StsCd_LIS_FinRst, lngTotCnt, lngInputCnt)
    ElseIf lngInputCnt > 0 Then '����� �ϳ��̻� �Է�
        SqlStmt = objAccSql.SqlUpdate201(pWorkArea, pAccDt, pAccSeq, enStsCd.StsCd_LIS_MidRst, lngTotCnt, lngInputCnt)
    Else    '��������� ����, �Է��� �ϳ��� �ȉ��� ���
        SqlStmt = objAccSql.SqlUpdate201(pWorkArea, pAccDt, pAccSeq, enStsCd.StsCd_LIS_Accession, lngTotCnt, lngInputCnt)
    End If
    Call DBConn.Execute(SqlStmt)
    
    '������� ���� Insert
    If IsDataExists(T_LAB304, pWorkArea, pAccDt, pAccSeq) Then
        SqlStmt = objAccSql.SqlUpdateReason(pWorkArea, pAccDt, pAccSeq, pEmpId, pReason)
    Else
        SqlStmt = objAccSql.SqlInsertReason_New(pWorkArea, pAccDt, pAccSeq, pEmpId, pReason)
    End If
    Call DBConn.Execute(SqlStmt)
        
    objOrdDic.Clear
    objOrdDic.FieldInialize "orddt,ordno", "updfg"
    
    With objordSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 8
            If .Value = 1 Then
                .Col = 1: sOrdDt = Format(.Value, CS_DateDbFormat)  'ó����
                .Col = 2: sOrdNo = .Value   'ó���ȣ
                .Col = 7: sOrdSeq = .Value  'ó��Seq
                
                'ó��Body Update
                SqlStmt = objAccSql.SqlUpdateOrdB(pPtid, sOrdDt, sOrdNo, sOrdSeq, pFlag)
                Call DBConn.Execute(SqlStmt)
                
                'ó��Header Update
                If Not objOrdDic.Exists(sOrdDt & COL_DIV & sOrdNo) Then
                    SqlStmt = objAccSql.SqlUpdateOrdH(pPtid, sOrdDt, sOrdNo, pFlag)
                    Call DBConn.Execute(SqlStmt)
                    
                    objOrdDic.AddNew sOrdDt & COL_DIV & sOrdNo, "Y"
                End If
                
                '** �������� �߰� By M.G.Choi 2005.08.25
                '   ��� �� OCS������� ����ó�� (mdresult delete)
                SqlStmt = objAccSql.OCSResultDelete(pPtid, sOrdDt, sOrdNo, sOrdSeq, pFlag)
                
                Call DBConn.Execute(SqlStmt)
                
            End If
        Next
    End With
    DoCancelPart = True
    Exit Function
    
Err_Trap:
    DoCancelPart = False

End Function

Public Sub LoadReasonTemplate(ByRef cboReason As Object)

    Dim tmpRs As Recordset
    
    cboReason.Clear
    Set tmpRs = New Recordset
    tmpRs.Open objAccSql.SQLGetCancelReason, DBConn
    
    While (Not tmpRs.EOF)
        cboReason.AddItem Trim("" & tmpRs.Fields("cdval1").Value) & Space(3) & ": " & Trim("" & tmpRs.Fields("text1").Value)
        tmpRs.MoveNext
    Wend
    Set tmpRs = Nothing

End Sub

Public Function CheckStatus(ByVal pWorkArea As String, ByVal pAccDt As String, _
                            ByVal pAccSeq As String, ByRef objData As clsDictionary) As String
    
    Dim tmpStr As String
    Dim tmpRs As Recordset
    Dim strKey As String, strData As String
    
    CheckStatus = "0"
    
    Set tmpRs = New Recordset
    tmpRs.Open objAccSql.SqlCheckStatus(pWorkArea, pAccDt, pAccSeq), DBConn
    
    If tmpRs.EOF Then GoTo NoData
    
    CheckStatus = Trim("" & tmpRs.Fields("StsCD").Value)
    
    Select Case CheckStatus
        Case enStsCd.StsCd_LIS_Collection:  CheckStatus = STS_LIS_HaveSpc
        Case enStsCd.StsCd_LIS_Accession:   CheckStatus = STS_LIS_Access
        Case enStsCd.StsCd_LIS_InProcess:   CheckStatus = STS_LIS_Worksheet
        Case enStsCd.StsCd_LIS_MidRst:      CheckStatus = STS_LIS_MidRst
        Case enStsCd.StsCd_LIS_FinRst:      CheckStatus = STS_LIS_Verify
        Case enStsCd.StsCd_LIS_Modify:      CheckStatus = STS_LIS_Modify
        Case enStsCd.StsCd_LIS_Cancel:      CheckStatus = STS_LIS_Cancel
    End Select
    
    strKey = medConcatString(COL_DIV, pWorkArea, pAccDt, pAccSeq)
    '"stscd,ptid,orddoct,deptcd,wardid,roomid,bedid,hosilid," & _
                                 "spccd,coldt,coltm,colid,rcvdt,rcvtm,rcvid,multifg,spcnm"
    strData = medConcatString(COL_DIV, "" & tmpRs.Fields("stscd").Value, "" & tmpRs.Fields("ptid").Value, _
                                       "" & tmpRs.Fields("orddoct").Value, "" & tmpRs.Fields("deptcd").Value, _
                                       "" & tmpRs.Fields("wardid").Value, "" & tmpRs.Fields("roomid").Value, _
                                       "" & tmpRs.Fields("bedid").Value, "" & tmpRs.Fields("hosilid").Value, _
                                       "" & tmpRs.Fields("spccd").Value, "" & tmpRs.Fields("coldt").Value, _
                                       "" & tmpRs.Fields("coltm").Value, "" & tmpRs.Fields("colid").Value, _
                                       "" & tmpRs.Fields("rcvdt").Value, "" & tmpRs.Fields("rcvtm").Value, _
                                       "" & tmpRs.Fields("rcvid").Value, "" & tmpRs.Fields("multifg").Value, _
                                       "" & tmpRs.Fields("spcnm").Value)

    objData.AddNew strKey, strData

NoData:
    Set tmpRs = Nothing
    
End Function


'% ������ȣ�� �������� �߻��� �˻系���� �˻��Ѵ�.
Public Function DisplayOrder(ByRef tblOrdSheet As Object, ByVal pWorkArea As String, _
                             ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pStsNm As String) As Boolean

    Dim tmpRs As Recordset
    Dim tmpSql As String
    Dim tmpBarcode As String
    Dim i As Long

      
    tmpSql = objAccSql.SqlAccOrder(pWorkArea, pAccDt, pAccSeq)
    Set tmpRs = New Recordset
    tmpRs.Open tmpSql, DBConn
      
    If tmpRs.EOF Then
        '������ü ���� ��ȸ...
        tmpSql = objAccSql.SqlAccOrder(pWorkArea, pAccDt, pAccSeq, True)
        Set tmpRs = Nothing
        Set tmpRs = New Recordset
        tmpRs.Open tmpSql, DBConn
        
        If tmpRs.EOF Then
            MsgBox "�ش� ����Ÿ�� �����ϴ� !!", vbOKOnly + vbExclamation, "Message"
            GoTo NoData
        End If
    End If
      
    With tblOrdSheet
        .ReDraw = False
        .MaxRows = 0
        If tmpRs.RecordCount < 16 Then
            .MaxRows = 16
         Else
            .MaxRows = tmpRs.RecordCount
        End If
        For i = 1 To tmpRs.RecordCount
            .Row = i
            .Col = 1: .Value = Format(Trim("" & tmpRs.Fields("OrdDt").Value), CS_DateMask)   'ó����
            .Col = 2: .Value = "" & tmpRs.Fields("OrdNo").Value                              'ó���ȣ
            .Col = 3: .Value = "" & tmpRs.Fields("TestCd").Value                             '�˻��ڵ�
            .Col = 4: .Value = "" & tmpRs.Fields("TestNm").Value                             '�˻��
                      .ForeColor = DCM_LightBlue
            .Col = 5: .Value = pStsNm                                                   '����
                      If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then .Value = STS_LIS_Verify
            .Col = 6: .Value = Format(Trim("" & tmpRs.Fields("VfyDt").Value), CS_DateMask)   '������
            .Col = 7: .Value = "" & tmpRs.Fields("OrdSeq").Value                             'ó��Seq
            tmpRs.MoveNext
        Next
        .RowHeight(-1) = 12.5
        .ReDraw = True
    End With

NoData:
   Set tmpRs = Nothing

End Function

Private Function IsDataExists(ByVal pTblNm As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean

    Dim tmpRs As Recordset
    
    Set tmpRs = New Recordset
    tmpRs.Open objAccSql.SqlDataExists(pTblNm, pWorkArea, pAccDt, pAccSeq), DBConn
    
    IsDataExists = Not tmpRs.EOF
    Set tmpRs = Nothing
End Function

Private Function IsPartDataExists(ByVal pTblNm As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String, _
                              ByVal pPtid As String, ByVal pOrdDt As String, _
                              ByVal pOrdNo As String, ByVal pOrdSeq As String) As Boolean

    Dim tmpRs As Recordset
    
    Set tmpRs = New Recordset
    tmpRs.Open objAccSql.SqlPartDataExists(pTblNm, pWorkArea, pAccDt, pAccSeq, pPtid, pOrdDt, pOrdNo, pOrdSeq), DBConn
    
    IsPartDataExists = Not tmpRs.EOF
    Set tmpRs = Nothing
End Function

Private Function IsDataExistsForMic(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, _
                                    ByRef pWsCd As String, ByRef pWsUnit As String) As Boolean
'�̻��� ��ũ��Ʈ ������ �ֳ� ���� ã�°�
'Append By legends 2003/10/15

    Dim tmpRs As Recordset
    
    Set tmpRs = New Recordset
    tmpRs.Open objAccSql.SqlDataExistsForMic(pWorkArea, pAccDt, pAccSeq), DBConn
    
    If tmpRs.EOF Then
        IsDataExistsForMic = False
    Else
        pWsCd = tmpRs.Fields("wscd").Value & ""
        pWsUnit = tmpRs.Fields("wsunit").Value & ""
        IsDataExistsForMic = True
    End If
    
    Set tmpRs = Nothing

End Function

'-----------------------------------------------------------------------------'
'   ��� : �̻��� Worksheet �������� - �̻��(2004-12-08)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Private Sub DelMicWorksheet(ByVal pWorkArea As String, ByVal pAccDt As String, _
                ByVal pAccSeq As String)
    Dim Rs          As Recordset
    Dim strWsCd     As String       'WorkSheet Code
    Dim strWsUnit   As String       'WorkSheet Unit
    Dim aryTemp()   As String       'WsCdWsUnit �迭
    Dim SQL         As String
    Dim i           As Long
    
On Error GoTo Errors
    '## 1.402���� ������ȣ�� WSCD, WSUNIT��ȸ
    Set Rs = New Recordset
    Rs.Open objAccSql.SqlDataExistsForMic(pWorkArea, pAccDt, pAccSeq), DBConn
    If Not (Rs.BOF Or Rs.EOF) Then
        ReDim aryTemp(Rs.RecordCount)
        
        Do Until Rs.EOF
            strWsCd = Rs.Fields("wscd").Value & ""
            strWsUnit = Rs.Fields("wsunit").Value & ""
            
            aryTemp(i) = strWsCd & COL_DIV & strWsUnit
            i = i + 1
            Rs.MoveNext
        Loop
    Else
        Set Rs = Nothing
        Exit Sub
    End If
    Rs.Close
    
    '## 2.402�� WORKAREA, ACCDT, ACCSEQ�� ����
    SQL = objAccSql.SqlDelRstTableForMic(T_LAB402, pWorkArea, pAccDt, pAccSeq)
    DBConn.Execute SQL
    
    '## 3.402�� WSCD, WSUNIT���� ��ȸ�� ������ 401����
    For i = 0 To UBound(aryTemp) - 1
        strWsCd = medGetP(aryTemp(i), 1, COL_DIV)
        strWsUnit = medGetP(aryTemp(i), 2, COL_DIV)
        Rs.Open objAccSql.SqlGetMicBody(strWsCd, strWsUnit)
        If Rs.BOF Or Rs.EOF Then
            SQL = objAccSql.SqlDelRstTableForMic(T_LAB401, strWsCd, strWsUnit)
            DBConn.Execute SQL
        End If
        Rs.Close
    Next i
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    MsgBox Err.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �̻��� Worksheet ���� �κл��� - �̻��(2004-12-13)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pPtId     : ȯ��ID
'       - pOrdDt    : ó������
'       - pOrdNo    : ó���ȣ
'       - pOrdSeq   : ó�����
'-----------------------------------------------------------------------------'
Private Sub DelMicWorksheetPart(ByVal pWorkArea As String, ByVal pAccDt As String, _
                ByVal pAccSeq As String, ByVal pPtid As String, ByVal pOrdDt As String, _
                ByVal pOrdNo As String, ByVal pOrdSeq As String)
    Dim Rs          As Recordset
    Dim strWsCd     As String       'WorkSheet Code
    Dim strWsUnit   As String       'WorkSheet Unit
    Dim aryTemp()   As String       'WsCdWsUnit �迭
    Dim SQL         As String
    Dim i           As Long
    
On Error GoTo Errors
    '## 1.402���� ������ȣ�� WSCD, WSUNIT��ȸ
    Set Rs = New Recordset
    Rs.Open objAccSql.SqlDataExistForMicPart(pWorkArea, pAccDt, pAccSeq, pPtid, pOrdDt, _
            pOrdNo, pOrdSeq), DBConn
    If Not (Rs.BOF Or Rs.EOF) Then
        ReDim aryTemp(Rs.RecordCount)
        
        Do Until Rs.EOF
            strWsCd = Rs.Fields("wscd").Value & ""
            strWsUnit = Rs.Fields("wsunit").Value & ""
            
            aryTemp(i) = strWsCd & COL_DIV & strWsUnit
            i = i + 1
            Rs.MoveNext
        Loop
    Else
        Set Rs = Nothing
        Exit Sub
    End If
    Rs.Close
    
    '## 2.402�� WORKAREA, ACCDT, ACCSEQ�� ����
    SQL = objAccSql.SqlDelRstTableForMicPart(pWorkArea, pAccDt, pAccSeq, pPtid, pOrdDt, pOrdNo, pOrdSeq)
    DBConn.Execute SQL
    
    '## 3.402�� WSCD, WSUNIT���� ��ȸ�� ������ 401����
    For i = 0 To UBound(aryTemp) - 1
        strWsCd = medGetP(aryTemp(i), 1, COL_DIV)
        strWsUnit = medGetP(aryTemp(i), 2, COL_DIV)
        Rs.Open objAccSql.SqlGetMicBody(strWsCd, strWsUnit)
        If Rs.BOF Or Rs.EOF Then
            SQL = objAccSql.SqlDelRstTableForMic(T_LAB401, strWsCd, strWsUnit)
            DBConn.Execute SQL
        End If
        Rs.Close
    Next i
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    MsgBox Err.Description, vbCritical, "����"
End Sub

'2009.05.20 �缺�� �߰�
'%  50. ���ڵ��ȣ�� ������ȣ���ϱ�
'%       - Calling FROM [ frm108AccCancel ] :
Public Function SqlAccOrderByBar(ByVal pSpcYY As String, ByVal pSpcNo As String) As String

    SqlAccOrderByBar = " SELECT a.workarea, a.accdt, a.accseq" & _
                       " FROM  " & T_LAB201 & " a " & _
                       " WHERE " & DBW("a.spcyy", pSpcYY, 2) & " AND   " & DBW("a.spcno", pSpcNo, 2)

End Function

