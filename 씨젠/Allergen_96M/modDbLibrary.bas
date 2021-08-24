Attribute VB_Name = "modDbLibrary"
Option Explicit


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
    
    Dim strKey1     As String
    Dim strKey2     As String
    Dim strSaveSeq  As String
    Dim strSubCodes As String
    Dim strChtNum   As String
    Dim strRegDate  As String
    Dim strOrdNm    As String
    Dim strOrdCd    As String
    Dim strReturn   As String
    
On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        strRegDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
        strOrdNm = Trim(GetText(.vasID, argSpcRow, colINOUT))

        Select Case strOrdNm
            Case "INHALANT":    strOrdCd = gAssayNM.INHALANT_CD
            Case "FOOD":        strOrdCd = gAssayNM.FOOD_CD
            Case "ATOPY":       strOrdCd = gAssayNM.ATOPY_CD
        End Select
        
        
        '-- Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf                            '����ڵ�
        SQL = SQL & "   AND DISKNO  = '" & strOrdNm & "'" & vbCrLf                          '����
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCrLf  '�˻���
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf                             '���ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq                                        '�����ȣ
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
        
        cn_Ser.BeginTrans
        
        '-- ������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))     '���(�����)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))     '���(�������)
            strRefVal = Trim(GetText(.vasTemp, iRow, 5))    '����
                        
            strSubCodes = Trim(GetText(.vasTemp, iRow, 14))    '����� �ڵ� : ex) 999|888|777

            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If lsID <> "" And strRegDate <> "" And sResult <> "" Then
                '-- ���������ƮD.  Interface_SetPatientResult02
                cn_Ser.Execute "Exec Interface_SetPatientResult02 '" & strRegDate & "'," & lsPid & ",'" & mGetP(strSubCodes, 1, "|") & "','" & mGetP(strSubCodes, 2, "|") & "','" & mGetP(strSubCodes, 3, "|") & "','" & sResult & "','','',0,0,0,'M010','" & strReturn & "'"
                If Res < 0 Then
                    SaveQuery SQL
                    Exit Function
                Else
                    SaveTransDataW = 1
                End If
                
            End If
        Next iRow
        
        cn_Ser.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    cn_Ser.RollbackTrans
    
End Function



