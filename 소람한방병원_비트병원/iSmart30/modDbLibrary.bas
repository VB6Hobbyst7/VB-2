Attribute VB_Name = "modDbLibrary"
Option Explicit


'''Private Function f_subSet_RefVal(ByVal strORCD As String, ByVal strSubCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
'''    Dim sqlRet      As Integer
'''    Dim sqlDoc      As String
'''    Dim stryy, strmm, strdd, strDate  As String
'''Dim rs_svr As ADODB.Recordset
'''
'''On Error GoTo ErrorTrap
'''
'''    strRSLT = Replace(strRSLT, "<", "")
'''    strRSLT = Replace(strRSLT, ">", "")
'''    f_subSet_RefVal = " "
'''
'''    f_subSet_RefVal = ""
'''          SQL = "Select REFHIGH, REFLOW  "
'''    SQL = SQL & "  From EQPMASTER"
'''    SQL = SQL & " Where EQUIPNO = '" & gEquip & "' "
'''    SQL = SQL & "   And EXAMCODE =  '" & strORCD & "'"
''''    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"
'''
'''    Res = GetDBSelectColumn(gLocal, SQL)
'''
'''    If Res > 0 Then
'''        If IsNumeric(strRSLT) And IsNumeric(Trim(gReadBuf(0))) And IsNumeric(Trim(gReadBuf(1))) Then
'''            If Val(strRSLT) > Val(Trim(gReadBuf(0))) Then
'''                f_subSet_RefVal = "H"
'''            ElseIf Val(strRSLT) < Val(Trim(gReadBuf(1))) Then
'''                f_subSet_RefVal = "L"
'''            Else
'''                f_subSet_RefVal = " "
'''            End If
'''        Else
'''            f_subSet_RefVal = " "
'''        End If
'''    End If
'''
'''Exit Function
'''
'''ErrorTrap:
'''    f_subSet_RefVal = " "
''''    Set AdoRs_ORACLE = Nothing
''''    Call ErrMsgProc(CallForm)
'''
'''End Function

Private Function f_subSet_RefVal(ByVal strORCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
    Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap
    
    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "
    
    f_subSet_RefVal = ""
    If strAge <> "" Then
        If strAge <= 7 Then
            SQL = "Select YMAX as MAX, YMIN as MIN "
        Else
            If strSex = "M" Then
                     SQL = "Select MMAX as MAX, MMIN as MIN "
            Else
                     SQL = "Select WMAX as MAX, WMIN as MIN "
            End If
        End If
    Else
        SQL = "Select MMAX as MAX, MMIN as MIN "
    End If
    
    SQL = SQL & "  From LABMAST"
    SQL = SQL & " Where ORCD =  '" & strORCD & "'"
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        If IsNumeric(strRSLT) And IsNumeric(rs_svr.Fields("MAX")) And IsNumeric(rs_svr.Fields("MIN")) Then
            If Val(strRSLT) > Val(rs_svr.Fields("MAX")) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(rs_svr.Fields("MIN")) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
        rs_svr.MoveNext
    
    Loop
    
Exit Function

ErrorTrap:
     
End Function

Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim strDate         As String
    Dim strTime         As String
    Dim strInNum        As String
    Dim strGumNum       As String
    Dim VallsID         As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strExamCd       As String
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
    
    '-- �������
    Dim prm1    As New ADODB.Parameter
    Dim prm2    As New ADODB.Parameter
    Dim prm3    As New ADODB.Parameter
    Dim prm4    As New ADODB.Parameter
    Dim prm5    As New ADODB.Parameter
    Dim prm6    As New ADODB.Parameter
    Dim prm7    As New ADODB.Parameter
    Dim prm8    As New ADODB.Parameter
    Dim prm9    As New ADODB.Parameter
    Dim prm10   As New ADODB.Parameter
    Dim strInCD     As String
    Dim strInVal    As String
    Dim intTotCnt   As Integer
    Dim strJumin    As String
    Dim strLabSeq   As String
    Dim strPartGbn  As String
    
'On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strPartGbn = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        strSpcCd = Trim(GetText(.vasID, argSpcRow, colINOUT))
        strLabSeq = Trim(GetText(.vasID, argSpcRow, colPOSNO))
        
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        
'        If InStr(lsID, "��������") > 0 Then
'            Exit Function
'        End If
'
'        If Not IsNumeric(lsID) Then
'            Exit Function
'        End If
'
'        If Len(lsID) <> 10 Then
'            Exit Function
'        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '����ڵ�
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'  " & vbCrLf                                      '�˻���
'        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf       '���ڵ�
        SQL = SQL & "   AND CHARTNO = '" & strChtNum & "' " & vbCrLf       '���ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq       '�����ȣ
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colDISKNO)) & "' " & vbCrLf         'DISK ��ȣ(����˻�ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPOSNO)) & "' "                    'POS ��ȣ(��������ID)
            
        Call SetSQLData("�������1", SQL)
        
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
                
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 1))
            strExamCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '���(�����)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
            strSex = Trim(GetText(.vasTemp, iRow, 8))
            strAge = Trim(GetText(.vasTemp, iRow, 10))
            strORQN = Trim(GetText(.vasTemp, iRow, 14))
            strJumin = Trim(GetText(.vasTemp, iRow, 12))
            
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then
                '-- ���������Ʈ ����
                SQL = ""
                SQL = SQL & "UPDATE ME_LABDAT Set" & vbCr
                SQL = SQL & "  LABRESULT = '" & sResult & "', " & vbCr  '�˻���
                SQL = SQL & "  LABENDDEP = '2', " & vbCr                'ó������       2:����, 3:����Է�
                SQL = SQL & "  LABRSTDTE = '" & strDate & "', " & vbCr  '����Է�����   YYYY-MM-DD
                SQL = SQL & "  LABRSTTIM = '" & strTime & "', " & vbCr  '����Է�����   YYYY-MM-DD
                SQL = SQL & "  LABRSTUID = '" & gUserID & "', " & vbCr  '����Է�ID
                SQL = SQL & "  LABRSTCOM = '" & gEquip & "' " & vbCr    '����Է���ǻ�͸�
                SQL = SQL & " WHERE LABATTEND = '" & lsPid & "'" & vbCr '������ȣ
                'SQL = SQL & "   And LABBARCOD = '" & lsID & "'" & vbCr  '���ڵ�
                SQL = SQL & "   And LABODRCOD = '" & strExamCd & "'" & vbCr    '�˻��ڵ�
                SQL = SQL & "   And LABODRSTP = '" & strAge & "'"       '�˻��Ϸù�ȣ
                '-- 2016.11.28�߰�
                SQL = SQL & "  And LABODRDTE = '" & strHospDate & "'"
                
                Call SetSQLData("�������", SQL)
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    Exit Function
                End If
                    
                If Res = 1 Then
                    '-- ���º���
                    SQL = ""
                    SQL = SQL & "UPDATE ME_DAT Set " & vbCr
                    SQL = SQL & "  DATENDDEP   = '2', " & vbCr         'ó������       2:����, 3:����Է�
                    SQL = SQL & "  DATRSTDTE = '" & strDate & "', " & vbCr  '����Է�����   YYYY-MM-DD
                    SQL = SQL & "  DATRSTTIM = '" & strTime & "', " & vbCr  '����Է½ð�   hh:mm:ss
                    SQL = SQL & "  DATRSTUID = '" & gUserID & "', " & vbCr  '����Է�ID
                    SQL = SQL & "  DATRSTCOM = '" & gEquip & "' " & vbCr    '����Է���ǻ�͸�
                    SQL = SQL & " WHERE DATATTEND = '" & lsPid & "'" & vbCr '������ȣ
                    'SQL = SQL & "   And DATBARCOD = '" & lsID & "'" & vbCr  '���ڵ�
                    SQL = SQL & "   And DATODRCOD = '" & strExamCd & "'" & vbCr    '�˻��ڵ�
                    SQL = SQL & "   And DATODRSTP = '" & strAge & "'"       '�˻��Ϸù�ȣ
                    '-- 2016.11.28�߰�
                    SQL = SQL & "  And DATODRDTE = '" & strHospDate & "'"
                        
                    Call SetSQLData("��������", SQL)
                    
                    Res = SendQuery(gServer, SQL)
                    
                    If Res < 0 Then
                        SaveQuery SQL
                        Exit Function
                    End If
                    
                    SaveTransDataW = 1
                End If

            End If
        Next iRow
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    
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
Function GetSampleInfoW_SWH(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim strORQN     As String
    
    GetSampleInfoW_SWH = -1
    strORQN = ""
    
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- ���̺� ���
    SQL = ""
    SQL = SQL & "SELECT M.LABDATE as ��������, M.PARTGBN, M.LABSEQ, M.SPECIMENCD, M.NAME as �̸�, M.AGE as ����, M.IDLEFT+M.IDRIGHT AS JUMIN, M.DEPTCD, M.JUBSUGBN, M.MEDICALREMARK as ���ڵ��ȣ, " & vbCr
    SQL = SQL & "       M.RESULTENDYN, R.SUBMCD, R.TESTITEMSEQ as ITEM, R.RESULTDATE, R.RESULTTIME, R.ROUTINECD, R.MACHINECD, R.RESULT, R.DELTAGBN, R.REFGBN, R.PANICGBN, R.DELTAMARK, R.REFMARK, R.PANICMARK, R.CMCODE " & vbCr
    
    Select Case gTblNm
        Case "H": SQL = SQL & "  FROM H_JUBSU M, H_RESULT R " & vbCr
        Case "J": SQL = SQL & "  FROM J_JUBSU M, J_RESULT R " & vbCr
        Case "M": SQL = SQL & "  FROM M_JUBSU M, M_RESULT R " & vbCr
        Case "S": SQL = SQL & "  FROM S_JUBSU M, S_RESULT R " & vbCr
        Case "U": SQL = SQL & "  FROM U_JUBSU M, U_RESULT R " & vbCr
    End Select
    
    SQL = SQL & " WHERE M.MEDICALREMARK = '" & sBarcode & "'" & vbCr
    SQL = SQL & "   AND M.LABDATE= R.LABDATE " & vbCr
    SQL = SQL & "   AND M.PARTGBN = R.PARTGBN " & vbCr
    SQL = SQL & "   AND M.LABSEQ = R.LABSEQ " & vbCr
    SQL = SQL & "   AND M.SPECIMENCD = R.SPECIMENCD" & vbCr
    SQL = SQL & "   And R.TESTITEMSEQ IN (" & gAllExam & ")" & vbCr
    SQL = SQL & "   AND R.CONFIRM <> 'Y' "
    

'    SQL = ""
'    SQL = SQL & "Select DISTINCT a.ORDT as ��������,'0',b.PANM as �̸�,a.SPNO as ���ڵ��ȣ,a.OIFL,'0',b.SEXS as ����,b.AGES as ����,a.NWNO as ������ȣ,a.ORCD as ITEM,a.ORQN as ITEMSEQ " & vbCr
'    SQL = SQL & "  From LRESULT a, APATINF b" & vbCr
'    SQL = SQL & " Where a.SPNO =  '" & sBarcode & "'"
'    SQL = SQL & "   And a.PAID = b.PAID " & vbCr
'    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")" & vbCr
'    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- ���Ȯ������
    
    Call SetSQLData("���ڵ���ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        With frmInterface
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                'strORQN = strORQN & Trim(RS.Fields("ITEM")) & "," & Trim(RS.Fields("ITEMSEQ")) & "|"
                
'Public Const colCHARTNO = 6     '                   >> PARTGBN
'Public Const colPID = 7        '���Ϲ�ȣ(������ȣ)  >> TESTITEMSEQ
'Public Const colINOUT = 8      '��ü                >> SPECIMENCD
'Public Const colDISKNO = 9      '                   >> JUMIN
'Public Const colPOSNO = 10      '                   >> LABSEQ
                
                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", asRow, colBARCODE
                SetText .vasID, Trim(RS.Fields("PARTGBN")) & "", asRow, colCHARTNO
                'SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("SPECIMENCD")) & "", asRow, colINOUT
                SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPAGE
                SetText .vasID, Trim(RS.Fields("JUMIN")) & "", asRow, colDISKNO
                SetText .vasID, Trim(RS.Fields("LABSEQ")) & "", asRow, colPOSNO
                
                
'                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
'                SetText .vasID, "20160801", asRow, colHOSPDATE
'                SetText .vasID, sBarcode, asRow, colBARCODE
'                SetText .vasID, "KS-123", asRow, colCHARTNO
'                SetText .vasID, "ȫ�浿", asRow, colPNAME
'                SetText .vasID, "Serum", asRow, colINOUT
'                SetText .vasID, "24", asRow, colPAGE
'                SetText .vasID, "7103111010911", asRow, colDISKNO
'                SetText .vasID, "12345", asRow, colPOSNO
                
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
        
            GetSampleInfoW_SWH = 1
        
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
Function GetSampleInfoW_SSH(ByVal asRow As Long) As Integer
    Dim sDate       As String
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim strORQN     As String
    
    GetSampleInfoW_SSH = -1
    strORQN = ""
    
    
    sDate = Format(Trim(GetText(frmInterface.vasID, asRow, colHOSPDATE)), "####-##-##")
    'sDate = Format(sDate, "####-##-##")
    
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colCHARTNO))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- ���̺� ���
    SQL = ""
    SQL = SQL & "SELECT L.LABSERIAL, L.LABATTEND as ������ȣ, L.LABCHTNUM as íƮ��ȣ, L.LABODRDTE as ��������, M.MANADMFOR as �Կ�," & vbCrLf
    SQL = SQL & "       M.MANRESNUM as �ֹι�ȣ, M.MANPATNAM as �̸�, L.LABINSNUM as ó�����,L.LABSMPNAM as ��ü��, L.LABBARCOD as ���ڵ��ȣ, L.LABODRCOD as ITEM, L.LABODRSTP as SEQ  " & vbCrLf
    SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M" & vbCrLf
'    SQL = SQL & " WHERE L.LABBARCOD = '" & sBarcode & "' " & vbCrLf
    SQL = SQL & " WHERE L.LABCHTNUM = '" & sBarcode & "' " & vbCrLf
    SQL = SQL & "   AND L.LABODRDTE = '" & sDate & "'"
    SQL = SQL & "   AND L.LABKEYNUM = D.DATKEYNUM " & vbCrLf                    '-- ���̺���Ű��
    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND " & vbCrLf                    '-- ������ȣ
    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND " & vbCrLf                    '-- ������ȣ
    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM " & vbCrLf                    '-- íƮ��ȣ
    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM " & vbCrLf                    '-- íƮ��ȣ
    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE " & vbCrLf                    '-- ó������
    SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)" & vbCrLf
    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllExam & ")" & vbCrLf
'    SQL = SQL & "   AND L.LABSUBYON = 'Y' " & vbCrLf                            '-- �����ڵ忩�� (����Է¿� �����ڵ��̸� Y)
    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL) " & vbCrLf    '-- ��ҿ���
    
'    '-- ���������
'    If chkSaveAll.Value = "0" Then
'        SQL = SQL & "   AND L.LABENDDEP < '3' " & vbCrLf                            '-- ó������ (2:����, 3:����Է�)
'        SQL = SQL & "   AND D.DATENDDEP < '3'"                                      '-- �����μ�ó������    CHAR(2)   1:���, 2:����Է´��, 3:�Ϸ�, 9.����
'    ElseIf chkSaveAll.Value = "1" Then
'        SQL = SQL & "   AND L.LABENDDEP <= '3' " & vbCrLf                            '-- ó������ (2:����, 3:����Է�)
'        SQL = SQL & "   AND D.DATENDDEP <= '3'"                                      '-- �����μ�ó������    CHAR(2)   1:���, 2:����Է´��, 3:�Ϸ�, 9.����
'    End If
   
    SQL = SQL & " ORDER BY L.LABINSNUM, L.LABODRCOD "
   
    Call SetSQLData("���ڵ���ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        With frmInterface
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                
                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", asRow, colBARCODE
                SetText .vasID, Trim(RS.Fields("íƮ��ȣ")) & "", asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("��ü��")) & "", asRow, colSPCNM
                SetText .vasID, Trim(RS.Fields("SEQ")) & "", asRow, colPAGE
                
                Select Case Trim(Trim(RS.Fields("�Կ�")) & "")
                    Case "A":   SetText .vasID, "�ܷ�", asRow, colINOUT
                    Case "F":   SetText .vasID, "�Կ�", asRow, colINOUT
                    Case Else:  SetText .vasID, "", asRow, colINOUT
                End Select
                
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
        
            GetSampleInfoW_SSH = 1
        
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
Function GetSampleInfoW_HMHOSP(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim strORQN     As String
    
    GetSampleInfoW_HMHOSP = -1
    strORQN = ""
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
'''
'''SQL.Text:=' select (select distinct SUNAME from cpl2010 where specimen_ser = '''+TMaster.FBarCode+''') SUNAME '+#13#10+
'''' , (select distinct Bunho from cpl2010 where specimen_ser = '''+TMaster.FBarCode+''') Bunho '+#13#10+
'''' , R.HANGMOG_CODE '+#13#10+
'''' , E.GUMSA_NAME '+#13#10+
'''' , E.JANGBI_OUT_CODE '+#13#10+
'''' , E.JANGBI_CODE '+#13#10+
'''' , R.CONFIRM_YN '+#13#10+
'''' , R.CPL_RESULT '+#13#10+
'''//' , R.JANGBI_YN '+#13#10+
'''' , R.JANGBI_CODE '+#13#10+
'''' from CPL3020 R '+#13#10+
'''' , EQP1010 E '+#13#10+
'''' where R.SPECIMEN_SER ='''+TMaster.FBarCode+''' '+#13#10+
'''' and NVL(R.CONFIRM_YN, ''N'') = ''N'' '+#13#10+
'''' and E.JANGBI_CODE = '''+TGlobal.FICode+''' '+#13#10+
'''' and E.JANGBI_OUT_CODE is Not Null '+#13#10+
'''' and R.HANGMOG_CODE = E.HANGMOG_CODE '+#13#10+
'''// ' and R.SPECIMEN_CODE = E.SPECIMEN_CODE '+#13#10+
'''' order by E.jangbi_out_code ';

'      -- ���̺� ���
'    SQL = ""
'    SQL = SQL & "Select DISTINCT a.ORDT as ��������,'0',b.PANM as �̸�,a.SPNO as ���ڵ��ȣ,a.OIFL,'0',b.SEXS as ����,b.AGES as ����,a.NWNO as ������ȣ,a.ORCD as ITEM,a.ORQN as ITEMSEQ " & vbCr
'    SQL = SQL & "  From LRESULT a, APATINF b" & vbCr
'    SQL = SQL & " Where a.SPNO =  '" & sBarcode & "'"
'    SQL = SQL & "   And a.PAID = b.PAID " & vbCr
'    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")" & vbCr
'    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- ���Ȯ������


          SQL = "SELECT (SELECT DISTINCT SUNAME FROM CPL2010 WHERE specimen_ser = '" & sBarcode & "') as �̸�," & vbCr
    SQL = SQL & "       (SELECT DISTINCT Bunho  FROM CPL2010 WHERE specimen_ser = '" & sBarcode & "') as ������ȣ," & vbCr
    SQL = SQL & "       R.HANGMOG_CODE, E.GUMSA_NAME, E.JANGBI_OUT_CODE as ITEM, E.JANGBI_CODE, R.CONFIRM_YN, R.CPL_RESULT, R.JANGBI_CODE as ITEM " & vbCr
    SQL = SQL & "  FROM CPL3020 R, EQP1010 E " & vbCr
    SQL = SQL & " WHERE R.SPECIMEN_SER = '" & sBarcode & "'" & vbCr
    SQL = SQL & "   AND NVL(R.CONFIRM_YN, 'N') = 'N' " & vbCr
    SQL = SQL & "   AND E.JANGBI_CODE = '" & gEquipCode & "'" & vbCr
    SQL = SQL & "   AND E.JANGBI_OUT_CODE is Not Null " & vbCr
    SQL = SQL & "   AND R.HANGMOG_CODE = E.HANGMOG_CODE " & vbCr
    SQL = SQL & " ORDER BY E.jangbi_out_code "

'    SetRawData "[SQL]" & SQL

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        With frmInterface
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                'strORQN = strORQN & Trim(RS.Fields("ITEM")) & "," & Trim(RS.Fields("ITEMSEQ")) & "|"
                
                SetText .vasID, "1", .vasID.MaxRows, colCheckBox
                'SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                SetText .vasID, sBarcode, asRow, colBARCODE
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colCHARTNO
                'SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                'SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPSEX
                'SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPAGE
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
        
            GetSampleInfoW_HMHOSP = 1
        
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


