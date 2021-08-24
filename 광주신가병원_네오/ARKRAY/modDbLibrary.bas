Attribute VB_Name = "modDbLibrary"
Option Explicit


Private Function f_subSet_RefVal(ByVal strORCD As String, ByVal strSubCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap

    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "

    f_subSet_RefVal = ""
          SQL = "Select REFHIGH, REFLOW  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPNO = '" & gEquip & "' "
    SQL = SQL & "   And EXAMCODE =  '" & strORCD & "'"
'    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"

    Res = GetDBSelectColumn(gLocal, SQL)

    If Res > 0 Then
        If IsNumeric(strRSLT) And IsNumeric(Trim(gReadBuf(0))) And IsNumeric(Trim(gReadBuf(1))) Then
            If Val(strRSLT) > Val(Trim(gReadBuf(0))) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(Trim(gReadBuf(1))) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
    End If

Exit Function

ErrorTrap:
    f_subSet_RefVal = " "
'    Set AdoRs_ORACLE = Nothing
'    Call ErrMsgProc(CallForm)

End Function

''Private Function f_subSet_RefVal(ByVal strORCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
''    Dim sqlRet      As Integer
''    Dim sqlDoc      As String
''    Dim stryy, strmm, strdd, strDate  As String
''    Dim rs_svr As ADODB.Recordset
''
''On Error GoTo ErrorTrap
''
''    strRSLT = Replace(strRSLT, "<", "")
''    strRSLT = Replace(strRSLT, ">", "")
''    f_subSet_RefVal = " "
''
''    f_subSet_RefVal = ""
''    If strAge <> "" Then
''        If strAge <= 7 Then
''            SQL = "Select YMAX as MAX, YMIN as MIN "
''        Else
''            If strSex = "M" Then
''                     SQL = "Select MMAX as MAX, MMIN as MIN "
''            Else
''                     SQL = "Select WMAX as MAX, WMIN as MIN "
''            End If
''        End If
''    Else
''        SQL = "Select MMAX as MAX, MMIN as MIN "
''    End If
''
''    SQL = SQL & "  From LABMAST"
''    SQL = SQL & " Where ORCD =  '" & strORCD & "'"
''
''    Set rs_svr = cn_Ser.Execute(SQL)
''    Do Until rs_svr.EOF
''        If IsNumeric(strRSLT) And IsNumeric(rs_svr.Fields("MAX")) And IsNumeric(rs_svr.Fields("MIN")) Then
''            If Val(strRSLT) > Val(rs_svr.Fields("MAX")) Then
''                f_subSet_RefVal = "H"
''            ElseIf Val(strRSLT) < Val(rs_svr.Fields("MIN")) Then
''                f_subSet_RefVal = "L"
''            Else
''                f_subSet_RefVal = " "
''            End If
''        Else
''            f_subSet_RefVal = " "
''        End If
''        rs_svr.MoveNext
''
''    Loop
''
''Exit Function
''
''ErrorTrap:
''
''End Function

Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim strDate         As String
    Dim strInNum        As String
    Dim strGumNum       As String
    Dim VallsID         As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strSubCD        As String
    Dim strRefVal       As String
    Dim strSpcCd        As String
    Dim strSex As String
    Dim strAge  As String
    Dim strORQN As String
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    Dim strReceNo   As String
    Dim strSeqNo   As String
    
    Dim tmpREF As String
    Dim strREF As String
    Dim GumEqpCd As String * 100
    
    Dim strExamDate As String
    Dim strHospDate As String
    
    Dim strKey1     As String
    Dim strKey2     As String
    Dim strSaveSeq  As String
    Dim strSubCodes As String
    Dim strChtNum   As String
    
    Dim strInCD     As String
    Dim strInVal    As String
    Dim intTotCnt   As Integer
    
    Dim strOrderCd  As String
    
    
    Dim strOrd_Seq_No   As String
    Dim strOrd_Cd       As String
    Dim strHL           As String
    
'On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        
        strOrd_Seq_No = Trim(GetText(.vasID, argSpcRow, colPSEX))
        strOrd_Cd = Trim(GetText(.vasID, argSpcRow, colPAGE))
        
'        If Len(lsID) <> 8 Then
'            Exit Function
'        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '����ڵ�
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'  " & vbCrLf                                      '�˻���
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf       '���ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq       '�����ȣ
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colDISKNO)) & "' " & vbCrLf         'DISK ��ȣ(����˻�ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPOSNO)) & "' "                    'POS ��ȣ(��������ID)
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        sResult1 = ""
        sResult2 = ""
        strKey1 = ""
        strKey2 = ""
        
        strInCD = ""
        strInVal = ""
        intTotCnt = 0
        
        cn_Ser.BeginTrans
        
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '���(�����)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
            
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then

'>> RESA
'RESA_MEDM_ID    NVARCHAR    10          ���� �ε���     NOT NULL
'RESA_KIND       SMALLINT                0�ܷ� 1�Կ�     NOT NULL
'RESA_KEY        NVARCHAR    20          ��¥(6)+Ÿ�̸�(7)+��������(2)+��(3)+seq�����Ű���(2) ��) 16021962420800200201       NOT NULL
'RESA_SEQ        INT                     resa ������     NOT NULL
'RESA_CNT        SMALLINT                resa count      NOT NULL
'RESA_CHAM_INDEX NVARCHAR    10          ȯ�� �ε���     NULL
'RESA_GWAM_ID    NVARCHAR    3           �����      NULL
'RESA_DATE       NVARCHAR    8           ��¥        NULL
'RESA_DEPT_ID    NVARCHAR    20          ���޺μ�        NULL
'RESA_SLIP_ID    NVARCHAR    30          �������̵�      NULL
'RESA_CODE       NVARCHAR    20          �˻��ڵ�        NULL
'RESA_TIME       NVARCHAR    4           ST (ä��ð�)       NULL
'RESA_FRESULT    NVARCHAR    50          �ӻ�����ġ (�ּ�)       NULL
'RESA_TRESULT    NVARCHAR    50          �ӻ�����ġ (�ִ�)       NULL
'RESA_RESULT     NVARCHAR    50          ���        NULL
            
                      SQL = "Update E_ORDER..RESA" & Format(Now, "yyyy") & vbCrLf
                SQL = SQL & " Set "
                SQL = SQL & " RESA_RESULT     = '" & sResult & "'" & vbCrLf '�˻���
'                SQL = SQL & ",RESA_BIGO5     = '1'" & vbCrLf                '�˻����� �������̽����� ����Ǹ� íƮ��ȣ�� ���� ��ܿ� ���� ���� ��Ÿ���� �ϴ°�..
                SQL = SQL & " Where RESA_CHAM_INDEX     = '" & Val(lsID) & "'" & vbCrLf        '��ü��ȣ
                SQL = SQL & "   and RESA_DATE = '" & strHospDate & "'"
                SQL = SQL & "   and RESA_CODE = '" & strEqpCd & "'"

   
                Call SetSQLData("�������", SQL)
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                
                'LAB_HISTORY2017 ���̺� ����ð�, RESA_KEY, TAT, ��� �Է��� �� �� ���缭 �Է����ֽø� �˴ϴ�. (�˻������� ������ ����)
'LAB_HISTORY2017
'LAB_MEDM_ID LAB_DATETIME    LAB_KEY LAB_CNT LAB_USRM    LAB_MEMO    LAB_RESULT  LAB_TAT LAB_BIGO1   LAB_BIGO2   LAB_BIGO3   LAB_BIGO4   LAB_BIGO5
'0   2017-02-08 10:27:00 17020834633330102101    0   33017   AU680   0.57    50
'0   2017-02-08 10:39:00 17020834633330102101    1   33017   ����    0.57    61  62600   31  39  C3720
'
'
'insert into LAB_HISTORY2017 (LAB_MEDM_ID,LAB_DATETIME,LAB_KEY,LAB_CNT,LAB_USRM,LAB_MEMO,LAB_RESULT,LAB_TAT) values
'('0','"& format(now,"yyyy-mm-dd hh:mm:ss") &"','','0','33017','ARKRAY','')
                
                
            End If
        Next iRow
        
        cn_Ser.CommitTrans
        SaveTransDataW = 1
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    cn_Ser.RollbackTrans
    
End Function


'Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
''������ ����Ÿ ���̽��� ����
'    Dim iRow            As Integer
'    Dim lsID            As String
'    Dim lsPid           As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strEqpCd        As String
'    Dim VallsID         As String
'    Dim strDate         As String
'
'    SaveTransDataR = -1
'
'    'Local���� ȯ�ں��� ����� ��������
'    ClearSpread frmInterface.vasTemp
'
'    With frmInterface
'        lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, 2))
'        VallsID = lsID
'        lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, 5))
'        strDate = Format(CDate(.dtpExamDate.Value), "yyyymmdd")
'
'        '-- Local���� ȯ�ں��� ����� ��������
'        ClearSpread .vasTemp
'
'              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
'        SQL = SQL & "  FROM PATRESULT " & vbCrLf
'        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                            '����ڵ�
'        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf   '�˻���
'        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, 2)) & "' " & vbCrLf     '���ڵ�
'        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK ��ȣ
'        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "                    'POS ��ȣ
'
'        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'
'        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
'
'        sResult = ""
'        sResult1 = ""
'        sResult2 = ""
'
'        cn_Ser.BeginTrans
'
'        '������ ����� �����ϱ�
'        For iRow = 1 To .vasTemp.DataRowCnt
'            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
'            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '���(�����)
'            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
'
'            '-- ���������
'            If .optSaveResultR(0).Value = True Then
'                sResult = sResult1
'            Else
'                sResult = sResult2
'            End If
'
'            If sResult <> "" Then
''                If Len(VallsID) > 6 Then
''                    SQL = "Update ONIT..GUMJIN_INTERFACE" & _
''                          "   Set RESULT = '" & sResult & "'," & _
''                          "       ACT_RETURN_DATE = '" & strDate & "'" & _
''                          " Where PER_GUMJIN_DATE = '" & Mid(lsID, 1, 8) & "'" & _
''                          "   And PER_GUM_NUM = " & lsID & "" & _
''                          "   And INTERFACECODE = '" & strEqpCd & "'"
''                Else
'                    SQL = "Update onit_out..jun370_resulttb" & _
'                          "   Set Result = '" & sResult & "'" & _
'                          " Where orderorder = '" & VallsID & "'" & _
'                          "   and map2seqno = '" & strEqpCd & "'"
''                End If
'
'                Res = SendQuery(gServer, SQL)
'
'                If Res < 0 Then
'                    SaveQuery SQL
'                    cn_Ser.RollbackTrans
'                    Exit Function
'                End If
'
'            End If
'        Next iRow
'
'    End With
'
'    cn_Ser.CommitTrans
'    SaveTransDataR = 1
'
'End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'          SQL = " SELECT DISTINCT '' AS ��������"
'    SQL = SQL & ", '' AS ��Ʈ��ȣ"
'    SQL = SQL & ", '' AS ������ȣ"
'    SQL = SQL & ", '' AS �Կ�"
'    SQL = SQL & ", '' AS �̸�"
'    SQL = SQL & ", '' AS ����"
'    SQL = SQL & ", '' AS ����" & vbCrLf
'    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'    SQL = SQL & " WHERE QC_BAR_NO = '" & sBarcode & "'" & vbCrLf
    
    
                                   
     
          'SQL = "SELECT P.PbsPatNam, O.OSPCHTNUM, R.ResLabCod, E.LabShtNam, R.ResOcmNum, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
          SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, R.ResOcmNum " & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, "1", asRow, colCheckBox
        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
        'SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE       '������
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO        'íƮ��ȣ
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID            '��Ϲ�ȣ(����� �ʿ�)
        'SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT          '��/��
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPNAME          'ȯ�ڸ�
        'SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX           '����
        'SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE           '����
        
        GetSampleInfoW = 1
   
    Else
        GetSampleInfoW = -1
    End If

    frmInterface.vasID.RowHeight(-1) = 12

End Function


'-- �˻��� ���� ��������
Function GetSampleInfoW_JBUNIV(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_JBUNIV = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    strTestCd = mGetP(frmInterface.cboTest.Text, 2, "|")
    pFrDt = Format(frmInterface.dtpStartDt.Value, "yyyymmdd") & "000000"
    pToDt = Format(frmInterface.dtpStopDt.Value, "yyyymmdd") & "235959"
    pFrNo = frmInterface.txtStartNum.Text
    pToNo = frmInterface.txtStopNum.Text
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- ���ϴ뺴��  r010m.SPCCD
    SQL = ""
    SQL = SQL & "SELECT '1', '' AS SN ,'' AS ����Ͻ�, j011m.colldt AS ��������, j011m.bcno AS ���ڵ��ȣ, j010m.bcprtno AS ��Ʈ��ȣ" & vbCr
    SQL = SQL & "       , r010m.WKYMD||r010m.WKGRPCD||r010m.WKNO FLWKNO " & vbCr
    SQL = SQL & "       , r010m.WKNO AS ������ȣ" & vbCr
    SQL = SQL & "       , j011m.regno AS ������ȣ" & vbCr
    SQL = SQL & "       , j010m.patnm AS �̸�" & vbCr
    SQL = SQL & "       , j010m.age AS ����" & vbCr
    SQL = SQL & "       , j010m.sex AS ����" & vbCr
    SQL = SQL & "       , j011m.IOGBN" & vbCr
    SQL = SQL & "       , j010m.DEPTCD" & vbCr
    SQL = SQL & "       , j010m.WARDNO" & vbCr
    SQL = SQL & "       , j010m.ROOMNO" & vbCr
    SQL = SQL & "       , f72m.testcd AS ITEM" & vbCr
    SQL = SQL & "       , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "  FROM LJ011M j011m                                     " & vbCr
    SQL = SQL & "       INNER JOIN LJ010M j010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno  = j010m.bcno              " & vbCr
    SQL = SQL & "              AND j011m.regno = j010m.regno             " & vbCr
    SQL = SQL & "       INNER JOIN LR010M r010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno   = r010m.bcno             " & vbCr
    SQL = SQL & "              AND j011m.regno  = r010m.regno            " & vbCr
    SQL = SQL & "              AND NVL(r010m.rstflg,'0') = '0'           " & vbCr
    SQL = SQL & "       INNER JOIN LF072M f72m                           " & vbCr
    SQL = SQL & "               ON f72m.eqcd    = 'G0006'                " & vbCr
    SQL = SQL & "              AND f72m.testcd  = '" & strTestCd & "'    " & vbCr
    SQL = SQL & "              AND r010m.testcd = f72m.testcd            " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND r010m.wkno between '" & pFrNo & "' AND '" & pToNo & "' " & vbCr
    SQL = SQL & "   AND j011m.spcflg  = '4'                        " & vbCr
    SQL = SQL & "   AND NVL(j011m.rstflg, '0')  = '0'            " & vbCr
    SQL = SQL & "   AND j011m.bcno = '" & sBarcode & "'" & vbCr
    SQL = SQL & " UNION                                              " & vbCr
    SQL = SQL & "SELECT '1', '' AS SN ,'' AS ����Ͻ�, j011m.colldt AS ��������, j011m.bcno AS ���ڵ��ȣ, j010m.bcprtno AS ��Ʈ��ȣ " & vbCr
    SQL = SQL & "        , r010m.FLWKNO" & vbCr
    SQL = SQL & "        , r010m.WKNO AS ������ȣ" & vbCr
    SQL = SQL & "        , j011m.regno AS ������ȣ" & vbCr
    SQL = SQL & "        , j010m.patnm AS �̸�" & vbCr
    SQL = SQL & "        , j010m.age AS ����" & vbCr
    SQL = SQL & "        , j010m.sex AS ����" & vbCr
    SQL = SQL & "        , j011m.IOGBN" & vbCr
    SQL = SQL & "        , j010m.DEPTCD" & vbCr
    SQL = SQL & "        , j010m.WARDNO" & vbCr
    SQL = SQL & "        , j010m.ROOMNO" & vbCr
    SQL = SQL & "        , f72m.testcd AS ITEM" & vbCr
    SQL = SQL & "        , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "   FROM LJ011M j011m                                " & vbCr
    SQL = SQL & "        INNER JOIN LJ010M j010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno  = j010m.bcno         " & vbCr
    SQL = SQL & "               AND j011m.regno = j010m.regno        " & vbCr
    SQL = SQL & "        INNER JOIN LM010M r010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno   = r010m.bcno        " & vbCr
    SQL = SQL & "               AND j011m.regno  = r010m.regno       " & vbCr
    SQL = SQL & "               AND NVL(r010m.rstflg,'0') = '0'      " & vbCr
    SQL = SQL & "        INNER JOIN LF072M f72m                      " & vbCr
    SQL = SQL & "                ON f72m.eqcd    = 'G0006'           " & vbCr
    SQL = SQL & "                AND f72m.testcd  = '" & strTestCd & "' " & vbCr
    SQL = SQL & "               AND r010m.testcd = f72m.testcd       " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND r010m.wkno between '" & pFrNo & "' AND '" & pToNo & "' " & vbCr
    SQL = SQL & "   AND j011m.spcflg  = '4'               " & vbCr
    SQL = SQL & "   AND NVL(j011m.rstflg, '0')  = '0'     " & vbCr
    SQL = SQL & "   AND j011m.bcno = '" & sBarcode & "'"
    SQL = SQL & " ORDER BY FLWKNO  " & vbCr

    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
            
            SetText .vasID, "1", .vasID.MaxRows, colCheckBox
            SetText .vasID, Trim(RS.Fields("��������")) & "", .vasID.MaxRows, colHOSPDATE
            SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .vasID.MaxRows, colBARCODE
            SetText .vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .vasID.MaxRows, colCHARTNO
            SetText .vasID, Trim(RS.Fields("������ȣ")) & "", .vasID.MaxRows, colPID
            SetText .vasID, Trim(RS.Fields("�̸�")) & "", .vasID.MaxRows, colPNAME
            SetText .vasID, Trim(RS.Fields("����")) & "", .vasID.MaxRows, colPSEX
            SetText .vasID, Trim(RS.Fields("����")) & "", .vasID.MaxRows, colPAGE
            SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
            
            '-- ȭ�鿡 ǥ��
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_JBUNIV = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function


'-- �˻��� ���� ��������
Function GetSampleInfoW_JAINCOM(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_JAINCOM = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- ���̺� ���
          SQL = "SELECT DiSTINCT b.SCP42JDATE as ��������, a.SCP41SPMNO2 as ���ڵ��ȣ, b.SCP42IDNOA as ������ȣ, a.SCP41NAME as �̸�, a.SCP41SEX as ����, a.SCP41BIRTH as ����,b.SCP42SUGACD as ITEM"
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & sBarcode & "'"
    If frmInterface.chkSaveAll.Value = "0" Then
    '    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
        SQL = SQL & vbCrLf & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '') "
    End If

    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
            
            SetText .vasID, "1", .vasID.MaxRows, colCheckBox
            SetText .vasID, Trim(RS.Fields("��������")) & "", .vasID.MaxRows, colHOSPDATE
            SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .vasID.MaxRows, colBARCODE
            'SetText .vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .vasID.MaxRows, colCHARTNO
            SetText .vasID, Trim(RS.Fields("������ȣ")) & "", .vasID.MaxRows, colPID
            SetText .vasID, Trim(RS.Fields("�̸�")) & "", .vasID.MaxRows, colPNAME
            SetText .vasID, Trim(RS.Fields("����")) & "", .vasID.MaxRows, colPSEX
            SetText .vasID, Trim(RS.Fields("����")) & "", .vasID.MaxRows, colPAGE
            'SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
            
            '-- ȭ�鿡 ǥ��
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_JAINCOM = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function
    

'-- �˻��� ���� ��������
Function GetSampleInfoW_MSINFOTEC(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim strORQN     As String
    
    GetSampleInfoW_MSINFOTEC = -1
    strORQN = ""
    
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- ���̺� ���
    SQL = ""
    SQL = SQL & "Select DISTINCT a.ORDT as ��������,'0',b.PANM as �̸�,a.SPNO as ���ڵ��ȣ,a.OIFL,'0',b.SEXS as ����,b.AGES as ����,a.NWNO as ������ȣ,a.ORCD as ITEM,a.ORQN as ITEMSEQ " & vbCr
    SQL = SQL & "  From LRESULT a, APATINF b" & vbCr
    SQL = SQL & " Where a.SPNO =  '" & sBarcode & "'"
    SQL = SQL & "   And a.PAID = b.PAID " & vbCr
    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")" & vbCr
    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- ���Ȯ������

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        With frmInterface
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                strORQN = strORQN & Trim(RS.Fields("ITEM")) & "," & Trim(RS.Fields("ITEMSEQ")) & "|"
                
                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", asRow, colBARCODE
                'SetText .vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPSEX
                SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPAGE
                'SetText .vasID, Trim(RS.Fields("SPCCD")) & "", asRow, colDISKNO
                
                '-- ȭ�鿡 ǥ��
                For intCol = colState + 1 To .vasID.MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                        .vasID.Row = asRow
                        .vasID.Col = intCol
                        .vasID.BackColor = vbYellow
                        Exit For
                    End If
                Next
        
                RS.MoveNext
            Loop
        
            GetSampleInfoW_MSINFOTEC = 1
        
        End With
    End If
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    gOrderExam = gOrderExam & "^" & strORQN
    
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_TWIN(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim strDate     As String
    
    GetSampleInfoW_TWIN = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    strDate = Format(Now, "yyyymmdd")
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'''          SQL = " SELECT DISTINCT C.JOBDATE,C.PTNO,C.JOBNO, DECODE(C.GBIO,'I','�Կ�','O','�ܷ�') AS GBIO,C.SNAME,C.SEX,C.AGE " & vbCrLf
'''    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A," & vbCrLf
'''    SQL = SQL & "       TW_HSP_OCS.TWEXAM_MASTER  B," & vbCrLf
'''    SQL = SQL & "       TW_HSP_OCS.TWEXAM_SPECMST C" & vbCrLf
'''    SQL = SQL & " Where C.SPECNO   = '" & sBarcode & "'" & vbCrLf   ' ��ü��ȣ
'''    SQL = SQL & "   And B.EQUCODE1 = '" & gEquipCode & "'" & vbCrLf ' ����ڵ�
'''    SQL = SQL & "   AND C.STATUS   <= '3' " & vbCrLf                 ' �˻����(4 : �κпϷ�)
'''    SQL = SQL & "   And (C.SPECNO  = A.SPECNO) " & vbCrLf
'''    SQL = SQL & "   And (A.SUBCODE = B.MASTERCODE)" & vbCrLf
'''    'SQL = SQL & "   AND A.MASTERCODE IN (" & gAllExam & ")"
    
    'SetRawData "[sql]" & SQL
          SQL = " SELECT DISTINCT a.JEOBSUDT,a.PTNO,a.slipno2, '' AS GBIO,b.SNAME,a.SEX,a.AGEYY " & vbCrLf
    SQL = SQL & "  From twexam_general_sub a," & vbCrLf
    SQL = SQL & "       tw_mis_pmpa.twbas_patient b " & vbCrLf
    SQL = SQL & " Where a.ptno = '" & sBarcode & "'" & vbCrLf   ' ��ü��ȣ
    SQL = SQL & "   And a.ptno = b.ptno " & vbCrLf ' ����ڵ�
    SQL = SQL & "   And a.jeobsudt = to_date('" & strDate & "', 'yyyy/mm/dd hh24/mi/ss') " & vbCrLf
    'SQL = SQL & "   And (C.SPECNO  = A.SPECNO) " & vbCrLf
    'SQL = SQL & "   And (A.SUBCODE = B.MASTERCODE)" & vbCrLf
    SQL = SQL & "   AND a.itemcd in (" & gAllExam & ")"
        
    SetSQLData "���ڵ���ȸ", SQL
    
'SELECT a.PTNO as PatientNo, b.SNAME as PatientName,a.SEX as PatientSex ,b.JUMIN1 || b.JUMIN2,a.JEOBSUDT as ReceiptDate, a.slipno1 || '-' || a.slipno2 as ReceiptNo
'FROM twexam_general_sub a, tw_mis_pmpa.twbas_patient b
'WHERE a.ptno = '{0}'
'And a.ptno=b.ptno
'And a.jeobsudt = to_date('{2}', 'yyyy/mm/dd hh24/mi/ss')
'And a.itemcd in ({1})
'Order by a.slipno2
    
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, "1", asRow, colCheckBox
        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT
        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, colPNAME
        SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX
        SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE
        
        GetSampleInfoW_TWIN = 1
   
    Else
        GetSampleInfoW_TWIN = -1
    End If

    frmInterface.vasID.RowHeight(-1) = 12

End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_NEOSOFT(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim strDate     As String
    
    GetSampleInfoW_NEOSOFT = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    strDate = Format(Now, "yyyymmdd")
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'>> ORDER_IN/ ORDER_OUT
'MEDM_ID        NVARCHAR    10          ���� �ε���     NOT NULL
'WORK_DATE      NVARCHAR    8           ó�� ����       NOT NULL
'CHAM_INDEX     NVARCHAR    10          íƮ �ε���     NOT NULL  >> ���ڵ��ȣ
'GWAM_ID        NVARCHAR    3           �������        NOT NULL
'DOC_ID         NVARCHAR    20          �����ǻ�        NOT NULL
'SEQ            INT                     ó�� ��ȣ(�Һ�) NOT NULL
'CNT            INT                     ó�� ����       NOT NULL
'CODE           NVARCHAR    20          ���� �ڵ�       NULL
'ID             NVARCHAR    20          �ɻ� �ڵ�       NULL
'
'>> HP_CHAM
'CHAM_ID        NVARCHAR    10          ȯ�� ���̵�     NOT NULL
'CHAM_INDEX     NVARCHAR    10          íƮ ��ȣ       NOT NULL
'CHAM_NAME      NVARCHAR    20          ȯ�� ����       NULL
'CHAM_YY        NVARCHAR    2           �ֹι�ȣ  �⵵  NULL
'CHAM_JUMIN1    NVARCHAR    64          �ֹι�ȣ 1      NULL
'CHAM_JUMIN2    NVARCHAR    64          �ֹι�ȣ 2      NULL
'CHAM_SEX       SMALLINT                ����            NULL ( 0 : ���� , 1 : ����)
    
    
          SQL = " SELECT DISTINCT a.MEDM_ID,a.CHAM_INDEX,a.WORK_DATE, '�Կ�' as IO,b.CHAM_NAME,b.CHAM_SEX,b.CHAM_YY " & vbCrLf
    SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a, E_BASECODE..HP_CHAM b " & vbCrLf
    SQL = SQL & " Where a.CHAM_INDEX = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And a.CHAM_INDEX = b.CHAM_INDEX " & vbCr
    SQL = SQL & "   AND a.CODE IN (" & gAllExam & ")"
    SQL = SQL & "   AND a.TRANS = '2' " & vbCr
    SQL = SQL & " UNION ALL "
    SQL = SQL & " SELECT DISTINCT a.MEDM_ID,a.CHAM_INDEX,a.WORK_DATE, '�ܷ�' as IO,b.CHAM_NAME,b.CHAM_SEX,b.CHAM_YY " & vbCrLf
    SQL = SQL & "  From E_ORDER..ORDER_OUT" & Format(Now, "yyyy") & " a, E_BASECODE..HP_CHAM b " & vbCrLf
    SQL = SQL & " Where a.CHAM_INDEX = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And a.CHAM_INDEX = b.CHAM_INDEX " & vbCr
    SQL = SQL & "   AND a.CODE IN (" & gAllExam & ")"
    SQL = SQL & "   AND a.TRANS = '2' " & vbCr
    SQL = SQL & " ORDER BY a.WORK_DATE, IO"
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, "1", asRow, colCheckBox
        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colCHARTNO
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colHOSPDATE
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT
        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, colPNAME
        SetText frmInterface.vasID, IIf(Trim(gReadBuf(5)) & "" = "0", "M", "F"), asRow, colPSEX
        
        GetSampleInfoW_NEOSOFT = 1
   
    Else
        GetSampleInfoW_NEOSOFT = -1
    End If

    frmInterface.vasID.RowHeight(-1) = 12

'                    .MaxRows = .MaxRows + 1
'                    SetText vasID, "1", .MaxRows, colCheckBox
'                    SetText vasID, Trim(RS.Fields("WORK_DATE")) & "", .MaxRows, colHOSPDATE
'                    SetText vasID, Format(Trim(RS.Fields("CHAM_INDEX")), "0000000000"), .MaxRows, colBARCODE
'                    SetText vasID, Trim(RS.Fields("MEDM_ID")) & "", .MaxRows, colCHARTNO
'                    SetText vasID, Trim(RS.Fields("CHAM_INDEX")) & "", .MaxRows, colPID
'                    SetText vasID, Trim(RS.Fields("IO")) & "", .MaxRows, colINOUT
'                    SetText vasID, Trim(RS.Fields("CHAM_NAME")) & "", .MaxRows, colPNAME
'                    SetText vasID, IIf(Trim(RS.Fields("CHAM_SEX")) & "" = "0", "M", "F"), .MaxRows, colPSEX

End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_MCC(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim GetOrderExamCode    As String
    Dim intCol              As Integer
    Dim strTestCd           As String
    
    GetSampleInfoW_MCC = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    

'''          SQL = "SELECT DISTINCT ORD_YMD, BCODE_NO, RECEPT_NO, PTNT_NO,PTNT_NM,AGE,SEX,ORD_CD" & vbCr
'''    SQL = SQL & "  FROM MCCSI.H7LIS_BCODE_ORD " & vbCr
'''    SQL = SQL & " WHERE BCODE_NO = '" & sBarcode & "'" & vbCr
'''    SQL = SQL & "   AND ORD_CD IN (" & gAllExam & ") " & vbCr
'''    SQL = SQL & "   AND RESULT_TYPE = '20'" & vbLf & vbCr

'          SQL = "SELECT DISTINCT  a.ptnt_no, c.ptnt_nm, a.recept_no, a.spc_cd, "
'    SQL = SQL & " (select codeval1 from pm_mst_div_key1 where codediv = 'LAB01' and codekey1 = a.spc_cd) as spc_nm "
'    SQL = SQL & "      , a.sts_cd, a.acc_ymd, a.ord_cd "
'    SQL = SQL & "  FROM h3lab_result a, h1opdin b, hz_mst_ptnt c "
'    SQL = SQL & " WHERE a.recept_no = b.recept_no "
'    SQL = SQL & "   AND a.sutak_cd = ''"
'    SQL = SQL & "   AND a.ptnt_no  = c.ptnt_no"
'    SQL = SQL & "   AND a.sts_cd   = 'A'"                                                                           ' A:���� R:���"
'    SQL = SQL & "   AND a.acc_ymd between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
'    SQL = SQL & "   AND a.ord_cd IN (" & gAllExam & ") " & vbCr
'    SQL = SQL & " Order by recept_no "
    
          SQL = "select a.ptnt_no, c.ptnt_nm, a.recept_no, a.ord_no, a.ord_seq_no, a.ord_cd," & vbCr
    SQL = SQL & "       (select substr( max( apply_ymd || ord_nm),9) " & vbCr
    SQL = SQL & "          from hz_mst_lab_spc " & vbCr
    SQL = SQL & "         where ord_cd = a.ord_cd " & vbCr
    SQL = SQL & "           and spc_cd = a.spc_cd and apply_ymd <= a.acc_ymd) as ord_nm, " & vbCr
    SQL = SQL & "       a.spc_cd," & vbCr
    SQL = SQL & "       (select codeval1 " & vbCr
    SQL = SQL & "          from pm_mst_div_key1 " & vbCr
    SQL = SQL & "         where codediv = 'LAB01' and codekey1 = a.spc_cd) as spc_nm, " & vbCr
    SQL = SQL & "       a.sts_cd, a.acc_ymd, a.acc_time, a.ord_type, a.result_val, a.result_nm, a.hl_gb, a.dpa_gb, a.unit, a.vfy_ymd, a.vfy_time, a.vfy_empl_no " & vbCr
    SQL = SQL & "  from h3lab_result a, h1opdin b, hz_mst_ptnt c" & vbCr
    SQL = SQL & " where a.recept_no = b.recept_no " & vbCr
    SQL = SQL & "   and a.sutak_cd = '' " & vbCr
    SQL = SQL & "   and a.ptnt_no  = c.ptnt_no" & vbCr
    SQL = SQL & "   and a.sts_cd   = 'A' " & vbCr
    SQL = SQL & "   and a.recept_no = '" & sBarcode & "'" & vbCr
                
    Call SetSQLData("���ڵ���ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        With frmInterface
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ORD_CD")) & "',"
                
                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
                SetText .vasID, Trim(RS.Fields("acc_ymd")) & "", asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("recept_no")) & "", asRow, colBARCODE
                'SetText .vasID, Trim(RS.Fields("RECEPT_NO")) & "", asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("ptnt_no")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("ptnt_nm")) & "", asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("ord_seq_no")) & "", asRow, colPSEX
                SetText .vasID, Trim(RS.Fields("ord_no")) & "", asRow, colPAGE
                
                '-- ȭ�鿡 ǥ��
                For intCol = colState + 1 To .vasID.MaxCols
                    If Trim(RS.Fields("ord_cd")) = gArrEquip(intCol - colState, 3) Then
                        .vasID.Row = asRow
                        .vasID.Col = intCol
                        .vasID.BackColor = vbYellow
                        Exit For
                    End If
                Next
        
                RS.MoveNext
            Loop
        
            GetSampleInfoW_MCC = 1
        
        End With
    End If
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_JWINFO(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim strORQN     As String
    
    GetSampleInfoW_JWINFO = -1
    strORQN = ""
    
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- ���̺� ���
          SQL = "SELECT DISTINCT a.RECEIPTDATE as ��������, a.SPECIMENNUM as ���ڵ��ȣ, a.IPDOPD, a.RECEIPTNO as íƮ��ȣ, a.PTNO as ������ȣ, a.SNAME as �̸�, b.LABCODE as ITEM, a.ORDERCODE"
    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster a,SLA_LabResult b "
    SQL = SQL & vbCrLf & " WHERE a.SPECIMENNUM = '" & sBarcode & "'"
    SQL = SQL & vbCrLf & "   AND a.RECEIPTNO = b.RECEIPTNO "
    SQL = SQL & vbCrLf & "   AND a.OrderCode = b.OrderCode "
    SQL = SQL & vbCrLf & "   AND b.LABCODE IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND a.JSTATUS < '3'" & vbLf

    Call SetSQLData("���ڵ���ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        With frmInterface
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                'strORQN = strORQN & Trim(RS.Fields("ITEM")) & "," & Trim(RS.Fields("ITEMSEQ")) & "|"
                
                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                
                'If Trim(RS.Fields("���ڵ��ȣ")) & "" = "0" Then
                '    SetText .vasID, Trim(RS.Fields("íƮ��ȣ")) & "", asRow, colBARCODE
                'Else
                    SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", asRow, colBARCODE
                'End If
                
                SetText .vasID, Trim(RS.Fields("íƮ��ȣ")) & "", asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("ORDERCODE")) & "", asRow, colPSEX   'ORDERCODE
                SetText .vasID, IIf(Trim(RS.Fields("IPDOPD")) = 1, "�Կ�", "�ܷ�"), asRow, colINOUT
                
                
                '-- ȭ�鿡 ǥ��
                For intCol = colState + 1 To .vasID.MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                        .vasID.Row = asRow
                        .vasID.Col = intCol
                        .vasID.BackColor = vbYellow
                        Exit For
                    End If
                Next
        
                RS.MoveNext
            Loop
        
            GetSampleInfoW_JWINFO = 1
        
        End With
    End If
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    'gOrderExam = gOrderExam & "^" & strORQN
    
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_BIT(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    
    GetSampleInfoW_BIT = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
     
          'SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, E.LabShtNam, R.ResOcmNum, R.ResLabCod, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
          SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, R.ResOcmNum, R.ResLabCod AS EXAMCODE , R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    SetRawData "[SQL1]" & SQL
    
    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
            
            SetText .vasID, "1", asRow, colCheckBox
            SetText .vasID, sBarcode, asRow, colBARCODE
            SetText .vasID, Trim(RS.Fields("OSPCHTNUM")), asRow, colCHARTNO         'íƮ��ȣ(������� ����� �ʿ�)
            SetText .vasID, Trim(RS.Fields("ResOcmNum")), asRow, colPID             '��Ϲ�ȣ(���     ����� �ʿ�)
            SetText .vasID, Trim(RS.Fields("PbsPatNam")), asRow, colPNAME           'ȯ�ڸ�
            
            
            'SetText .vasID, "12345", asRow, colCHARTNO         'íƮ��ȣ
            'SetText .vasID, "67890", asRow, colPID            '��Ϲ�ȣ(����� �ʿ�)
            'SetText .vasID, "ȫ�渱", asRow, colPNAME           'ȯ�ڸ�
            
            '-- ȭ�鿡 ǥ��
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    '-- �������� SEQ
                    gArrEquip(intCol - colState, 7) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))   '�������� ��ȣ's
                    'gArrEquip(intCol - colState, 7) = "987" & "|" & "654" & "|" & "321"
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_BIT = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12

End Function

Public Function GetSameRowNum(ByVal strBarNo As String) As Integer
    Dim i As Integer

    GetSameRowNum = 0
    With frmInterface.vasID
        For i = 1 To .MaxRows
            .Row = i
            .Col = colBARCODE
            If Trim(.Text) = strBarNo Then
                GetSameRowNum = i
                Exit Function
            End If
        Next
    End With
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_GINUSDLL(ByVal asRow As Long) As Integer
    Dim pBarNo  As String
    Dim i       As Integer
    Dim intCol  As Integer
    Dim strItem As String
    
    '-- ������
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    GetSampleInfoW_GINUSDLL = -1
    
    pBarNo = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If pBarNo = "" Then
        Exit Function
    End If
    
    '-- �˻�ITEM ��������
                 strRequest = "jobs" + vbTab + "Q" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + pBarNo + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.MCD + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- ���ڵ�� �˻��� ��ȸ(https://211.172.17.66)
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    With frmInterface.vasID
        If UBound(varResponse) > 0 Then
            For i = 0 To UBound(varResponse) - 1
                SetText frmInterface.vasID, "1", asRow, colCheckBox
                SetText frmInterface.vasID, Mid(mGetP(varResponse(i), 25, vbTab), 1, 8), asRow, colHOSPDATE
                SetText frmInterface.vasID, mGetP(varResponse(i), 0, vbTab), asRow, colBARCODE
                SetText frmInterface.vasID, mGetP(varResponse(i), 7, vbTab), asRow, colPID
                SetText frmInterface.vasID, mGetP(varResponse(i), 26, vbTab), asRow, colPNAME
                
                Select Case mGetP(varResponse(i), 29, vbTab)
                    Case "O": SetText frmInterface.vasID, "�ܷ�", asRow, colINOUT
                    Case "E": SetText frmInterface.vasID, "����", asRow, colINOUT
                    Case "I": SetText frmInterface.vasID, "�Կ�", asRow, colINOUT
                End Select
                
                
                For intCol = colState + 1 To .MaxCols
                    If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colState, 3) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        '-- �������� SEQ
                        gArrEquip(intCol - colState, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                        Exit For
                    End If
                Next intCol
            Next i
        Else
            SetText frmInterface.vasID, "No Order", asRow, colState
        End If
    End With
    
    GetSampleInfoW_GINUSDLL = 1

'          SQL = " SELECT DISTINCT REQ_DT AS ��������"
'    SQL = SQL & ", LOT_NO AS ��Ʈ��ȣ"
'    SQL = SQL & ", REQ_SEQ AS ������ȣ"
'    SQL = SQL & ", '�Կ�' AS �Կ�"
'    SQL = SQL & ", 'ȫ�浿' AS �̸�"
'    SQL = SQL & ", '����' AS ����"
'    SQL = SQL & ", REQ_SEQ AS ����" & vbCrLf
'    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'    SQL = SQL & " WHERE QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
'    SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
'
'    Res = GetDBSelectColumn(gServer, SQL)
        
'    If Res = 1 Then
'        SetText frmInterface.vasID, "1", asRow, colCheckBox
'        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
'        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO
'        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID
'        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT
'        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, colPNAME
'        SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX
'        SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE
'
'        GetSampleInfoW_GINUSDLL = 1
'
'    Else
'        GetSampleInfoW_GINUSDLL = -1
'    End If
'
'    frmInterface.vasID.RowHeight(-1) = 12

End Function

'Function GetSampleInfoR(ByVal asRow As Long) As Integer
'    Dim sBarcode As String
'    Dim sSpecNo As String
'
'    GetSampleInfoR = -1
'
'    '-- ȯ������ ��������
'    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBARCODE))   '���� ���ڵ� ��ȣ
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'    '-- ���ڵ��ȣ�� ȯ������ �ҷ�����
'          SQL = "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE & vbCrLf
'    SQL = SQL & "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
'    SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
'    If gDBCOLUMN_Parm.STATUS <> "" Then
'        SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
'    End If
'    If gDBCOLUMN_Parm.RESULT <> "" Then
'        SQL = SQL + "   AND (" & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL)"
'    End If
'
'    Res = GetDBSelectColumn(gServer, SQL)
'
'    If Res = 1 Then
'        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
''        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPNAME
'        '-- ������ ������� �ֹι�ȣ�� ã��
'        'strSex = IIf(Mid(Trim(gReadBuf(4)), 7, 1) = "1", "M", "F")
'        'SetText frmInterface.vasID, strSex, colSex    '7  ����
''        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7  ����
'        '-- ���̰� ������� �ֹι�ȣ�� ã��
'        'strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
'        'SetText frmInterface.vasID, strAge, asRow, colAge
''        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSex    '8  ����
'
'        GetSampleInfoR = 1
'    Else
'
'        GetSampleInfoR = -1
'    End If
'
'End Function

Function GetEquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String

    GetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    Res = GetDBSelectVas(gLocal, SQL, frmInterface.vasTemp1)
    
    If frmInterface.vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        End If
    Next i

    'SPSLHRRST
    SQL = " Select SUCD From LRESULT " & vbCr & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    Res = GetDBSelectColumn(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        GetEquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function


Function GetGetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim strStatFg  As String
Dim sExamCd As String
Dim strItems As String
Dim strTemp As String
Dim strIntBase As String

    GetGetEquipExamCode_CA1500 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
'    argPID = "1558200030"
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    Res = GetDBSelectColumn(gServer, SQL)
    GetGetEquipExamCode_CA1500 = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and SUBSTR(exmn_cd,1,1) <> 'G'" & _
          "   and RSLT_NO IS NOT NULL"
          
    Res = GetDBSelectRow(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    Erase gReadBuf
          SQL = "Select equipcode "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    SQL = SQL & " order by equipcode    "
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strIntBase = Trim(gReadBuf(i))
            strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space$(6)
            If strIntBase <> strTemp Then
                strExamCode = strExamCode & strIntBase 'Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
                strTemp = strIntBase
            End If

            'strExamCode = strExamCode & Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    '�������� (R:Routin, E:Stat)
    'strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")
    strStatFg = "U"
    
    
'    strExamCode = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mOrder.RackNo & mOrder.TubePos & mOrder.BarNo & _
                "B" & Space(15) & strExamCode & ETX
    
    strExamCode = "" & "S2210101" & strStatFg & Space(6) & Space(4) & mResult.RackNo & mResult.TubePos & mResult.BarNo & _
                "B" & Space(15) & strExamCode & ""
    
    GetGetEquipExamCode_CA1500 = strExamCode
    
End Function

'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
    Dim RS As ADODB.Recordset
    Dim intCol As Integer

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
          

     
          SQL = "SELECT P.PbsPatNam, O.OSPCHTNUM, R.ResLabCod, E.LabShtNam, R.ResOcmNum, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & argPID & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    Set RS = cn_Ser.Execute(SQL)
    
    Do Until RS.EOF
        GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
        
        '-- ȭ�鿡 ǥ��
        With frmInterface
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = .vasID.ActiveRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
        End With
        
        RS.MoveNext
    Loop
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
    RS.Close
    
End Function

Function GetOrderExamCode_New(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_New = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    argPID = Mid(argPID, 1, 10)
    
          SQL = "SELECT DiSTINCT b.SCP42SUGACD "
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & argPID & "'"
    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
End Function


Function GetGetEquipExamCode_E411(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim sSpecNo     As String
Dim iRow        As Long
Dim SpecNo      As String
    
    GetGetEquipExamCode_E411 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- �ڰ�ü�� 11�ڸ��� ��ȸ�ϱ����Ͽ� ������ �ڸ��� ���ش�.
    argPID = Mid(argPID, 1, 10)
    
    If Mid(argPID, 1, 2) = "99" Then
        'strExamCode = Proc_Order_LX_QC(argPID)
        
        'iRow = frmInterface.vasID.DataRowCnt
        iRow = intRow
        
        SpecNo = Trim(GetText(frmInterface.vasID, iRow, colSpecNo))
        
        SQL = "SELECT QC_EXMN_CD "
        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// ��� ��ȣ
        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// �˻�� ��ȣ
        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// ���� ��ȣ
        SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & gAllExam & ") "
        SQL = SQL & vbCrLf & "  AND USE_STR_DT <= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
        SQL = SQL & vbCrLf & "  AND USE_END_DT >= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
        Res = GetDBSelectRow(gServer, SQL)
        strExamCode = ""
        
        For i = 0 To UBound(gReadBuf)
            If gReadBuf(i) <> "" Then
                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
            Else
                Exit For
            End If
        Next
        
    Else
        '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
        SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
        Res = GetDBSelectColumn(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
        '-- �˻��ڵ� ��������
        SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
              " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
              "   and RSLT_NO IS NOT NULL"
              
        Res = GetDBSelectRow(gServer, SQL)
        strExamCode = ""
        
        For i = 0 To UBound(gReadBuf)
            If gReadBuf(i) <> "" Then
                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
            Else
                Exit For
            End If
        Next
    End If
    
    If strExamCode = "" Then
'        MsgBox "������ ȯ��"
        GetGetEquipExamCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    
    '-- ������ �˻��ڵ��� ä�� ã��
          SQL = "Select distinct equipcode "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
    
        If gReadBuf(i) <> "" Then
            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    GetGetEquipExamCode_E411 = Mid(strExamCode, 2)
    
End Function



Function GetGetEquipExamCode_Architect(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_Architect = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- �˻��ڵ� ��������
'    SQL = ""
'    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.TESTCD & vbCrLf
'    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
'    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL"
'
'    Res = GetDBSelectRow(gServer, SQL)
'    strExamCode = ""
'
'    For i = 0 To UBound(gReadBuf)
'        If gReadBuf(i) <> "" Then
'            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'        Else
'            Exit For
'        End If
'    Next
'
'    If strExamCode = "" Then
'        '-- ������ȯ���̰ų� �ش���� �˻��� ����
'        GetGetEquipExamCode_Architect = ""
'        Exit Function
'    End If
'
'    '-- ������ "," �ڸ���
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = "          "
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    '-- �ش� ��� �°� ����ä�� �����ϱ� [ASTM Format >> Architect]
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    '-- ù�ڸ� "\" �ڸ���
    GetGetEquipExamCode_Architect = strExamCode
    
End Function

'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetGetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'AU480�� ��� ��񿡼� dilution ���� ���� '0'�߰�
            strExamCode = strExamCode & "0" & Trim(gReadBuf(i)) & "0"
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_AU480 = strExamCode
    
End Function


'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetGetEquipExamCode_CentaurCP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_CentaurCP = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_CentaurCP = strExamCode
    
End Function

'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetGetEquipExamCode_Hitachi7080(argEquipCode As String, argPID As String, Optional intRow As Long) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim strExamCode() As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_Hitachi7080 = ""
    j = 0
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    'strExamCode = ""

    'ReDim strExamCode(UBound(gReadBuf))
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            ReDim Preserve strExamCode(j)
            strExamCode(j) = Trim(gReadBuf(i))
            j = j + 1
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_Hitachi7080 = strExamCode
    
End Function
Function GetGetEquipExamCode(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetGetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- �ڰ�ü�� 11�ڸ��� ��ȸ�ϱ����Ͽ� ������ �ڸ��� ���ش�.
    argPID = Mid(argPID, 1, 10)
    
    SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
          " Where SPCM_NO = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    Res = GetDBSelectColumn(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
          SQL = "Select equipcode "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(argEquipCode) & ")"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & Trim(gReadBuf(i)) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    GetGetEquipExamCode = strExamCode
    
End Function


