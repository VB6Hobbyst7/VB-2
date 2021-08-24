Attribute VB_Name = "modDbLibrary"
Option Explicit


Public Function db_select_Vas(argServer As Integer, argSQL As String, ByVal argSpread As vaSpread, Optional argRow As Integer = 1, Optional argcol As Integer = 1) As Integer
'���� ���� ������ �������彬Ʈ�� Display
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas = -1
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set RS = cmdSQL.Execute
  
    If argSpread.MaxCols < RS.Fields.Count + argcol - 1 Then
        argSpread.MaxCols = RS.Fields.Count + argcol - 1
    End If
    
    If RS.EOF = True Or RS.BOF = True Then
        db_select_Vas = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argRow
    While Not RS.EOF
        If argSpread.MaxRows < i Then
            argSpread.MaxRows = i
        End If
        
        For j = 0 To RS.Fields.Count - 1
            argSpread.Row = i
            argSpread.Col = j + argcol
            If IsNull(RS.Fields.Item(j).Value) Then
                argSpread.Text = ""
            Else
                argSpread.Text = Trim(CStr(RS.Fields.Item(j).Value))
            End If
        Next j
        i = i + 1
        RS.MoveNext
    Wend
    
    
    
    If argSpread.DataRowCnt = 0 Then
        db_select_Vas = 0
    Else
        db_select_Vas = i - 1
        'argSpread.MaxRows = i - 1
    End If
    
    RS.Close
    
    Exit Function
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Vas = -1
    
End Function

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'���� ���� ������ gReadbuf()�� Array�� ����
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set RS = cmdSQL.Execute
           
    If Not (RS.EOF Or RS.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Col = 0
        gReadBuf(0) = ""
        RS.Close
        Exit Function
    End If
    
    
    Do While Not RS.EOF
        For i = 0 To RS.Fields.Count - 1
            If IsNull(RS.Fields.Item(i).Value) = True Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = Trim(CStr(RS.Fields.Item(i).Value))
            End If
        Next i
        
        db_select_Col = 1
        
        RS.MoveNext
        Exit Do
    Loop
    
    RS.Close
    
    Exit Function
    
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col = -1
End Function

Private Function f_subSet_RefVal(ByVal strORCD As String, ByVal strSubCD As String, Optional ByVal strRslt As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap
    
    strRslt = Replace(strRslt, "<", "")
    strRslt = Replace(strRslt, ">", "")
    f_subSet_RefVal = " "
    
    f_subSet_RefVal = ""
          SQL = "Select REFHIGH, REFLOW  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPNO = '" & gEquip & "' "
    SQL = SQL & "   And EXAMCODE =  '" & strORCD & "'"
'    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If IsNumeric(strRslt) And IsNumeric(Trim(gReadBuf(0))) And IsNumeric(Trim(gReadBuf(1))) Then
            If Val(strRslt) > Val(Trim(gReadBuf(0))) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRslt) < Val(Trim(gReadBuf(1))) Then
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



Public Sub SaveTrans(ByVal vasIDRow As Integer)
'��������
    'Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim FindFile As String

'    If MsgBox(" " & vbCrLf & "���������� �Ͻðڽ��ϱ�?" & vbCrLf & " ", vbInformation + vbOKCancel, "�˸�:��������") = vbCancel Then
'        Exit Sub
'    End If
    
'''      SQL.Text:=' Select RESA_SEQ, RESA_TIME From Month..RESA'+YYYYMM+         #13#10+
'''                ' Where RESA_CHAM_ID = '''+BCD+'''                    '+#13#10+
'''                '   And RESA_GWAM_ID = ''01''                         '+#13#10+
'''                '   And RESA_KIND = 0                                 '+#13#10+
'''                '   And RESA_DATE = '''+DD+'''                        '+#13#10+
'''                //'   And RESA_SLIP_ID = '''+TMaster.FSlip+'''          '+#13#10+
'''                '   And RESA_LABM_ID = '''+TMaster.FExamCode+'''      '+#13#10;
'''                //'   And (RESA_RESULTS is null or RESA_RESULTS = '''') ';
'''      Open;
'''
'''      If RecordCount > 0 Then begin
'''          Seq:= FieldByName('RESA_SEQ').AsInteger;
'''          sTm:= FieldByName('RESA_TIME').AsString;
'''
'''          Close;
'''          SQL.Text:=' Update Month..RESA'+YYYYMM+ ' Set '                         +#13#10+
'''                    '       RESA_RESULTS = :RESA_RESULTS                  '+#13#10+
'''                    '     , RESA_SLIP_ID = '''+SLP+''' '+#13#10+
'''                    ' Where RESA_CHAM_ID = '''+BCD+'''       '+#13#10+
'''                    '   And RESA_GWAM_ID = ''01''                         '+#13#10+
'''                    '   And RESA_KIND = 0                                 '+#13#10+
'''                    '   And RESA_DATE = '''+DD+'''                        '+#13#10+
'''                    //'   And RESA_SLIP_ID = '''+TMaster.FSlip+'''          '+#13#10+
'''                    '   And RESA_LABM_ID = '''+TMaster.FExamCode+'''      '+#13#10+
'''                    '   And RESA_SEQ     = :RESA_SEQ                      '+#13#10+
'''                    '   And RESA_TIME    = :RESA_TIME                     '+#13#10;
'''          Parameters.ParamByName('RESA_RESULTS').Value:= TMaster.FResult;
'''          Parameters.ParamByName('RESA_SEQ').Value:= Seq;
'''          Parameters.ParamByName('RESA_TIME').Value:= sTm;
'''      End
'''      Else: begin
'''          Close;
'''          Sql.Text:=' Insert Into Month..RESA'+YYYYMM+ '(  RESA_CHAM_ID         '+#13#10+
'''                    '             , RESA_GWAM_ID         '+#13#10+
'''                    '             , RESA_KIND            '+#13#10+
'''                    '             , RESA_DATE            '+#13#10+
'''                    '             , RESA_SLIP_ID         '+#13#10+
'''                    '             , RESA_SEQ             '+#13#10+
'''                    '             , RESA_TIME            '+#13#10+
'''                    '             , RESA_GUBUN           '+#13#10+
'''                    '             , RESA_RNO             '+#13#10+
'''                    '             , RESA_LABM_ID         '+#13#10+
'''                    '             , RESA_AMT             '+#13#10+
'''                    '             , RESA_DAY             '+#13#10+
'''                    '             , RESA_BIGUB           '+#13#10+
'''                    '             , RESA_RESULTN         '+#13#10+
'''                    '             , RESA_RESULTS         '+#13#10+
'''                    '             , RESA_USGM_ID         '+#13#10+
'''                    '             , RESA_RESULT_FLAG     '+#13#10+
'''                    '             , RESA_FLAG            '+#13#10+
'''                    '             , RESA_MEMO            '+#13#10+
'''                    '             , RESA_BIGO1           '+#13#10+
'''                    '             , RESA_BIGO2           '+#13#10+
'''                    '             , RESA_BIGO3 )         '+#13#10+
'''                    ' Values                             '+#13#10+
'''                    ' (  :RESA_CHAM_ID                   '+#13#10+
'''                    '  , :RESA_GWAM_ID                   '+#13#10+
'''                    '  , :RESA_KIND                      '+#13#10+
'''                    '  , :RESA_DATE                      '+#13#10+
'''                    '  , :RESA_SLIP_ID                   '+#13#10+
'''                    '  , :RESA_SEQ                       '+#13#10+
'''                    '  , :RESA_TIME                      '+#13#10+
'''                    '  , :RESA_GUBUN                     '+#13#10+
'''                    '  , :RESA_RNO                       '+#13#10+
'''                    '  , :RESA_LABM_ID                   '+#13#10+
'''                    '  , :RESA_AMT                       '+#13#10+
'''                    '  , :RESA_DAY                       '+#13#10+
'''                    '  , :RESA_BIGUB                     '+#13#10+
'''                    '  , :RESA_RESULTN                   '+#13#10+
'''                    '  , :RESA_RESULTS                   '+#13#10+
'''                    '  , :RESA_USGM_ID                   '+#13#10+
'''                    '  , :RESA_RESULT_FLAG               '+#13#10+
'''                    '  , :RESA_FLAG                      '+#13#10+
'''                    '  , :RESA_MEMO                      '+#13#10+
'''                    '  , :RESA_BIGO1                     '+#13#10+
'''                    '  , :RESA_BIGO2                     '+#13#10+
'''                    '  , :RESA_BIGO3 )                   '+#13#10;
'''
'''
'''                Parameters.ParamByName('RESA_CHAM_ID').Value:= BCD;
'''                Parameters.ParamByName('RESA_GWAM_ID').Value:= '01';
'''                Parameters.ParamByName('RESA_KIND').Value:= 0;
'''                Parameters.ParamByName('RESA_DATE').Value:= DD;
'''                Parameters.ParamByName('RESA_SLIP_ID').Value:= SLP;
'''                Parameters.ParamByName('RESA_SEQ').Value:= TMaster.FDispSeq;
'''                Parameters.ParamByName('RESA_TIME').Value:= FormatDateTime('hhnn', now);
'''                Parameters.ParamByName('RESA_GUBUN').Value:= 1;
'''                Parameters.ParamByName('RESA_RNO').Value:=  1 ;
'''                Parameters.ParamByName('RESA_LABM_ID').Value:= TMaster.FExamCode;
'''                Parameters.ParamByName('RESA_AMT').Value:= 1;
'''                Parameters.ParamByName('RESA_DAY').Value:= 1;
'''                Parameters.ParamByName('RESA_BIGUB').Value:= 0;
'''                If StrToFloatDef(TMaster.FResult, -1) >= 0 Then
'''                    Parameters.ParamByName('RESA_RESULTN').Value:= 1
'''                Else
'''                    Parameters.ParamByName('RESA_RESULTN').Value:= 0;
'''
'''                Parameters.ParamByName('RESA_RESULTS').Value:= TMaster.FResult;
'''                Parameters.ParamByName('RESA_USGM_ID').Value:= '';
'''                Parameters.ParamByName('RESA_RESULT_FLAG').Value:= False;
'''                Parameters.ParamByName('RESA_FLAG').Value:= False;
'''                Parameters.ParamByName('RESA_MEMO').Value:=  '';
'''                Parameters.ParamByName('RESA_BIGO1').Value:= '';
'''                Parameters.ParamByName('RESA_BIGO2').Value:= '';
'''                Parameters.ParamByName('RESA_BIGO3').Value:= '';
    
    With frmInterface
        For vasResRow = 1 To .vasTemp.DataRowCnt
            .vasID.Row = vasIDRow
            .vasID.Col = 1
            If .vasID.Value = 1 Then
                .vasTemp.Row = vasResRow
                liRet = -1
                If Trim(GetText(.vasTemp, vasResRow, 3)) <> "" Then
                    liRet = Make_XML(vasResRow)
                End If
                
                If liRet = 1 Then
                    SetBackColor .vasID, vasIDRow, vasIDRow, colCHECKBOX, colState - 1, 202, 255, 112
                    SetText .vasID, "���ۿϷ�", vasIDRow, colState
                    
                    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
                    If FindFile <> "" Then
                        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '���ۿϷᰡ ������ ���������
                    End If
                          
                          SQL = " Update PATRESULT Set "
                    SQL = SQL & " SENDFLAG = '2', "
                    SQL = SQL & " SENDDATE = '" & Format(Now, "yyyymmdd") & "' "
                    SQL = SQL & " Where BARCODE  = '" & GetText(frmInterface.vasID, vasIDRow, colBARCODE) & "' "
                    SQL = SQL & "   AND SAVESEQ  = " & GetText(frmInterface.vasID, vasIDRow, colSAVESEQ)
                    SQL = SQL & "   AND MID(EXAMDATE,1,8)   = '" & Mid(GetText(frmInterface.vasID, vasIDRow, colEXAMDATE), 1, 8) & "' "
                    Res = SendQuery(gLocal, SQL)

                Else
                    SetBackColor .vasID, vasIDRow, vasIDRow, colCHECKBOX, 12, 255, 0, 0
                    'SetText vasID, "����", vasIDRow, colState
                End If
                '.vasID.Col = 1
                '.vasID.Value = "0"
            Else
            
            End If
        Next
    End With
    
    If XmlTxtHead = "" Then
        XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                     "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare�˻�����>"
    End If
    
    If XmlTxtTail = "" Then
        XmlTxtTail = "</UBCare�˻�����>"
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    SaveXMLFile XMLAllTxt
    
End Sub


Public Sub SaveXMLFile(argSQL As String, Optional argFlag As Integer = 0)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum, FilNum1
    Dim FindFile As String
    Dim TxtString1 As String
    Dim AllString1 As String
    Dim i As Long
    
    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out.xml")
    
    
    If FindFile <> "" Then
        'Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
        FilNum1 = FreeFile
        Open "C:\UBCare\SINAI\IF\ExamIF_out.xml" For Input As FilNum1
        
        Do While Not EOF(FilNum1)
            Input #FilNum1, TxtString1
            Line Input #FilNum1, TxtString1
            AllString1 = AllString1 & TxtString1
        Loop

        Close #FilNum1
        i = InStr(1, AllString1, "</UBCare�˻�����>")
        XmlBody = Mid(AllString1, 1, i - 1)
        argSQL = XmlBody & argSQL & XmlTxtTail
        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
    Else
        argSQL = XmlTxtHead & argSQL & XmlTxtTail
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    
    FilNum = FreeFile
    
    
    If argFlag = 0 Then
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Output As FilNum
    Else
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    argSQL = ""
    
End Sub


Function Make_XML(asRow) As Integer
'Dim varTmp As Variant
'Dim strTmp As String
'Dim strRslt As String

'    With frmInterface.vasTemp
'        .Row = asRow
'                    XMLAllTxt = XMLAllTxt & "<�˻�>"
'        .Col = 1:   XMLAllTxt = XMLAllTxt & "<��ü>" & Trim(.Text) & "</��ü>"
'        .Col = 2:   XMLAllTxt = XMLAllTxt & "<�������ȣ>" & Trim(.Text) & "</�������ȣ>"
'        .Col = 3:   XMLAllTxt = XMLAllTxt & "<��Ʈ��ȣ>" & Trim(.Text) & "</��Ʈ��ȣ>"
'        .Col = 4:   XMLAllTxt = XMLAllTxt & "<�����ڸ�>" & Trim(.Text) & "</�����ڸ�>"
'        .Col = 7:   XMLAllTxt = XMLAllTxt & "<�ֹε�Ϲ�ȣ>" & Trim(.Text) & "</�ֹε�Ϲ�ȣ>"
'        .Col = 8:   XMLAllTxt = XMLAllTxt & "<������ȣ>" & Trim(.Text) & "</������ȣ>"
'        .Col = 9:   XMLAllTxt = XMLAllTxt & "<�Ƿ���>" & Trim(.Text) & "</�Ƿ���>"
'        .Col = 10:  XMLAllTxt = XMLAllTxt & "<�˻��ȣ>" & Trim(.Text) & "</�˻��ȣ>"
'        .Col = 11:  XMLAllTxt = XMLAllTxt & "<�˻�ID>" & Trim(.Text) & "</�˻�ID>"
'        .Col = 12:  XMLAllTxt = XMLAllTxt & "<��ü�˻�ID></��ü�˻�ID>"
'        .Col = 13:  XMLAllTxt = XMLAllTxt & "<��ü></��ü>"
'        .Col = 14:  strRslt = Trim(.Text)
'                    XMLAllTxt = XMLAllTxt & "<���ġ>" & strRslt & "</���ġ>"
'        .Col = 15:  XMLAllTxt = XMLAllTxt & "<����ġ>" & Trim(.Text) & "</����ġ>"
'        .Col = 16:  XMLAllTxt = XMLAllTxt & "<�Ұ�>" & Trim(.Text) & "</�Ұ�>"
'        .Col = 17:  XMLAllTxt = XMLAllTxt & "<�����>" & Trim(.Text) & "</�����>"
'        .Col = 18:  XMLAllTxt = XMLAllTxt & "<�Կ��ܷ�����>" & Trim(.Text) & "</�Կ��ܷ�����>"
'                    XMLAllTxt = XMLAllTxt & "</�˻�>"
'
'    End With
'
'    Make_XML = 1
    Dim FilNum
    Dim FilNum2
    Dim TxtString As String
    Dim ResultString As String
    Dim TxtRece As String
    Dim i As Long
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    
    Dim PExamname As String
    Dim PEquipCode As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim pResult As String
    Dim pExamdate As String
    Dim pOpinion As String
    Dim TxtPat As String
    Dim IOGubun As String
    Dim TestNum As String
    
    Make_XML = -1
    
    ClearSpread frmInterface.vasResTemp
    
    SQL = "select  chartno, examcode, hospdate, barcode,pname, pjumin, examdate, '', seqno,result " & vbCrLf & _
          "  from PATRESULT " & vbCrLf & _
          " where mid(examdate,1,8) = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' " & vbCrLf & _
          "   and result <> '' " & vbCrLf & _
          "   And equipno = '" & gEquip & "' " & _
          "   and barcode = '" & Trim(GetText(frmInterface.vasID, asRow, 5)) & "'"
    SQL = SQL & "order by seqno "
    Res = db_select_Vas(gLocal, SQL, frmInterface.vasResTemp)

    For i = 1 To frmInterface.vasResTemp.DataRowCnt
        PID = Trim(GetText(frmInterface.vasResTemp, i, 1))
        PExamCode = Trim(GetText(frmInterface.vasResTemp, i, 2))
        PReceDate = Trim(GetText(frmInterface.vasResTemp, i, 3))
        PChartNum = Trim(GetText(frmInterface.vasResTemp, i, 4))
        PName = Trim(GetText(frmInterface.vasResTemp, i, 5))
        PJumin = Mid(Trim(GetText(frmInterface.vasResTemp, i, 6)), 1, 6) & "-" & Mid(Trim(GetText(frmInterface.vasResTemp, i, 6)), 7)
        pExamdate = Trim(GetText(frmInterface.vasResTemp, i, 7))
        IOGubun = Trim(GetText(frmInterface.vasResTemp, i, 8))
        TestNum = Trim(GetText(frmInterface.vasResTemp, i, 9))
        pResult = Trim(GetText(frmInterface.vasResTemp, i, 10))
        XMLAllTxt = XMLAllTxt & "<�˻�><��ü>ACK</��ü><�������ȣ>38341948</�������ȣ><��Ʈ��ȣ>" & PChartNum & "</��Ʈ��ȣ><�����ڸ�>" & PName & "</�����ڸ�><�ֹε�Ϲ�ȣ>" & PJumin & "</�ֹε�Ϲ�ȣ><������ȣ>" & PID & "</������ȣ><�Ƿ���>" & PReceDate & "</�Ƿ���><�˻��ȣ>" & TestNum & "</�˻��ȣ><�˻�ID>" & PExamCode & "</�˻�ID><��ü�˻�ID></��ü�˻�ID><��ü></��ü><���ġ>" & pResult & "</���ġ><����ġ></����ġ><�Ұ�></�Ұ�><�����>" & pExamdate & "</�����><�Կ��ܷ�����>" & IOGubun & "</�Կ��ܷ�����></�˻�>"
    Next
    
    Make_XML = 1
    
End Function

Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim strDate         As String
    Dim strInNum        As String
    Dim strGumNum       As String
    Dim VallsID         As String
    Dim lsPid           As String
    Dim strBunHo        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strSubCD        As String
    Dim strRefVal       As String
    Dim strSex As String
    Dim strAge  As String
    Dim strORQN As String
    Dim strSeqNo  As String
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    Dim strReceNo   As String
'    Dim strSeqNo   As String
    
    Dim tmpREF As String
    Dim strREF As String
    'Dim GumEqpCd As String * 100
    
    Dim strExamDate As String
    
    Dim strKey1     As String
    Dim strKey2     As String
    
    Dim strYM       As String
'    Dim strDay      As String
    Dim strResaSeq  As String
    Dim strResaTime As String
    Dim strSaveSeq  As String
    Dim strRegDate  As String
    Dim strGubun    As String
    Dim strSpcCd    As String
    
    '-- �������
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    Dim prm6 As New ADODB.Parameter
    Dim prm7 As New ADODB.Parameter
    Dim prm8 As New ADODB.Parameter
    Dim prm9 As New ADODB.Parameter
    Dim prm10 As New ADODB.Parameter
    Dim prm11 As New ADODB.Parameter
    
    
    Dim strCMnt  As String
    Dim blnSave As Boolean
    
On Error GoTo ErrHandle

    'strYM = Format(Now, "yyyymm")
    'strDay = Format(Now, "dd")
    strCMnt = ""
    blnSave = False
    
    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strBunHo = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        strRegDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
        strGubun = Trim(GetText(.vasID, argSpcRow, colINOUT))
        
     '   MsgBox "1"
        
        
     '   strYM = Format(strRegDate, "yyyymm")
     '   strDay = Format(strRegDate, "dd")
        
        '-- Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
              SQL = "SELECT DISTINCT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf                            '����ڵ�
'        SQL = SQL & "   AND DISKNO  = '" & strOrdNm & "'" & vbCrLf                          '����
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCrLf  '�˻���
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf                             '���ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq                                        '�����ȣ
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colDISKNO)) & "' " & vbCrLf         'DISK ��ȣ(����˻�ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPOSNO)) & "' "                    'POS ��ȣ(��������ID)
        SQL = SQL & "  ORDER BY EQUIPCODE "
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        'MsgBox "2"
        
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
        
        'cn_Ser.BeginTrans
        
        '-- ������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))     '���(�����)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))     '���(�������)
            strRefVal = Trim(GetText(.vasTemp, iRow, 5))    '����
            strSeqNo = Trim(GetText(.vasTemp, iRow, 9))
            strSpcCd = Trim(GetText(.vasTemp, iRow, 15))
            
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            'covid-19
            
            If Mid(strEqpCd, 1, 4) = "LZ35" Then
                strCMnt = "�ǽð� ������ ����ȿ�ҿ��������(Realtime reverse transcriptase PCR)"
            Else
                strCMnt = "�ǽð�����ȿ�ҿ��������(Realtime PCR)"
            End If
            
            If strEqpCd = "XXX01" And sResult = "Positive" Then
            '    strCMnt = "Binary Toxin : Positive"
            ElseIf strEqpCd = "XXX02" And sResult = "Positive" Then
            '    strCMnt = strCMnt & vbNewLine & "TcdC : Positive"
            Else
                If sResult <> "" Then
                    '-- �˻������� = PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEMGT_U02
                    Set cmdSQL = New ADODB.Command
                    Set cmdSQL.ActiveConnection = cn_Ser
                    With cmdSQL
                        .CommandTimeout = 15 'MEDI.
                        .CommandText = "PR_CPL_CPL0891_INSERT"  'MEDI.PR_CPL_CPL0891_INSERT
                        .CommandType = adCmdStoredProc
                        
                        Set prm1 = .CreateParameter("I_JANGBI_NAME", adVarChar, adParamInput, 40, gEquipCode)
                        .Parameters.Append prm1
                        Set prm2 = .CreateParameter("I_SAMPLE_ID", adVarChar, adParamInput, 20, lsID)
                        .Parameters.Append prm2
                        Set prm3 = .CreateParameter("I_HANGMOG_CODE", adVarChar, adParamInput, 20, strEqpCd)
                        .Parameters.Append prm3
                        Set prm4 = .CreateParameter("I_CPL_RESULT", adVarChar, adParamInput, 50, sResult)
                        .Parameters.Append prm4
                        Set prm5 = .CreateParameter("I_CHK_FLAG", adVarChar, adParamInput, 1, "N")
                        .Parameters.Append prm5
                        Set prm6 = .CreateParameter("I_CONFIRM_YN", adVarChar, adParamInput, 1, "")
                        .Parameters.Append prm6
                        Set prm7 = .CreateParameter("I_FKCPL0201", adVarChar, adParamInput, 50, lsPid)
                        .Parameters.Append prm7
                        Set prm8 = .CreateParameter("I_SPECIMEN_CODE", adVarChar, adParamInput, 2, strSpcCd)
                        .Parameters.Append prm8
                        Set prm9 = .CreateParameter("I_EMERGENCY", adVarChar, adParamInput, 1, "N")
                        .Parameters.Append prm9
                        Set prm10 = .CreateParameter("I_JANGBI_RESULT", adVarChar, adParamInput, 50, sResult)
                        .Parameters.Append prm10
                        Set prm11 = .CreateParameter("I_JANGBI_FLAG", adVarChar, adParamInput, 2000, "") 'strCMnt
                        .Parameters.Append prm11
                         
                        SetRawData "[�������]" & gEquipCode & "," & lsID & "," & strEqpCd & "," & sResult & ",N, null" & "," & lsPid & "," & strSpcCd & ",N" & "," & sResult & "," & strCMnt
    
                        .Execute
                                
                        Set cmdSQL = Nothing
                        blnSave = True
                    End With
                    
                End If
            End If
        Next iRow
        
        If blnSave = True Then
'''            SQL = ""
'''            SQL = SQL & "Update cpl0302 "
'''            SQL = SQL & "   Set mb_result_txt1  = '" & strCMnt & "' " & vbCrLf  '�˻���
''''            SQL = SQL & "     , mb_result_txt2  = ''                " & vbCrLf  '�ӻ����ؼ�
'''            SQL = SQL & " Where FKCPL0201       = '" & lsPid & "'   " & vbCrLf
'''            SQL = SQL & "   And HANGMOG_CODE    = '" & strEqpCd & "'" & vbCrLf
            
'select * from cpl0302
' where  HANGMOG_CODE = 'C5957M'
'and FKCPL0201 = (select PKCPL0201  from cpl0201
' where specimen_ser = '2010151576'
'   and bunho = '0001511225')
'
            'strCMnt = "�ǽð�����ȿ�ҿ��������(Realtime PCR)"
            
            If strEqpCd = "XXX01" Or strEqpCd = "XXX02" Then
                strEqpCd = "CZ99901"
            End If
            
            SQL = ""
            SQL = SQL & "Update cpl0302"
            SQL = SQL & "   Set mb_result_txt1  = '" & strCMnt & "' " & vbCrLf  '�˻���
            SQL = SQL & " Where HANGMOG_CODE    = '" & strEqpCd & "'" & vbCrLf
            SQL = SQL & "   and FKCPL0201       = (Select PKCPL0201 "
            SQL = SQL & "                            From cpl0201"
            SQL = SQL & "                           Where Specimen_Ser  = '" & lsID & "'"
            SQL = SQL & "                             And BUNHO         = '" & strBunHo & "')"
            
            SetRawData "[�˻�������]" & SQL
            
            Res = SendQuery(gServer, SQL)
        End If
        
        'cn_Ser.CommitTrans
        SaveTransDataW = 1
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    'cn_Ser.RollbackTrans
    
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
        SetText frmInterface.vasID, "1", asRow, colCHECKBOX
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
'    strTestCd = mGetP(frmInterface.cboTest.Text, 2, "|")
    pFrDt = Format(frmInterface.dtpStartDt.Value, "yyyymmdd") & "000000"
    pToDt = Format(frmInterface.dtpStopDt.Value, "yyyymmdd") & "235959"
'    pFrNo = frmInterface.txtStartNum.Text
'    pToNo = frmInterface.txtStopNum.Text
    
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
            
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
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
            
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
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
Function GetSampleInfoW_SWMC(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_SWMC = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- ���̺� ���
'''          SQL = "SELECT DiSTINCT b.SCP42JDATE as ��������, a.SCP41SPMNO2 as ���ڵ��ȣ, b.SCP42IDNOA as ������ȣ, a.SCP41NAME as �̸�, a.SCP41SEX as ����, a.SCP41BIRTH as ����,b.SCP42SUGACD as ITEM"
'''    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
'''    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
'''    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
'''    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
'''    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
'''    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & sBarcode & "'"
'''    If frmInterface.chkSaveAll.Value = "0" Then
'''    '    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
'''        SQL = SQL & vbCrLf & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '') "
'''    End If


          SQL = "SELECT DISTINCT PART_JUBSU_DATE as ��������, PART_JUBSU_TIME, BUNHO as ������ȣ, SUNAME as �̸�, AGE as ����, SEX as ����, "
    SQL = SQL & "       FKCPL0201, SPECIMEN_CODE, GWA, HANGMOG_CODE as ITEM, INTERFACE_YN, JANGBI_OUT_CODE, JANGBI_CODE, CONFIRM_YN, CPL_RESULT "
    SQL = SQL & "  FROM VW_CPL_INTERFACE_GUMSA_LOAD "
    SQL = SQL & " WHERE SPECIMEN_SER = '" & sBarcode & "'"
    SQL = SQL & "   AND NVL(CONFIRM_YN, 'N') = 'N'"
    SQL = SQL & "   AND HANGMOG_CODE IN (" & gAllExam & ")"
    '            //'   and JANGBI_CODE = '''+TGlobal.FICode+'''     ';

    'SetRawData "[������ȸ]" & SQL
    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
            
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
            SetText .vasID, Trim(RS.Fields("��������")) & "", .vasID.MaxRows, colHOSPDATE
'            SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .vasID.MaxRows, colBARCODE
            SetText .vasID, Trim(RS.Fields("������ȣ")) & "", .vasID.MaxRows, colCHARTNO
            SetText .vasID, Trim(RS.Fields("FKCPL0201")) & "", .vasID.MaxRows, colPID
            SetText .vasID, Trim(RS.Fields("SPECIMEN_CODE")) & "", .vasID.MaxRows, colINOUT
            SetText .vasID, Trim(RS.Fields("�̸�")) & "", .vasID.MaxRows, colPNAME
            SetText .vasID, Trim(RS.Fields("����")) & "", .vasID.MaxRows, colPSEX
            SetText .vasID, Trim(RS.Fields("����")) & "", .vasID.MaxRows, colPAGE
            'SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
            
'            SetText .vasID, "2016-07-01", .vasID.MaxRows, colHOSPDATE
'           ' SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .vasID.MaxRows, colBARCODE
'            SetText .vasID, "123456", .vasID.MaxRows, colCHARTNO
'            SetText .vasID, "S1203", .vasID.MaxRows, colPID
'            SetText .vasID, "Serum", .vasID.MaxRows, colINOUT
'            SetText .vasID, "�̹���", .vasID.MaxRows, colPNAME
'            SetText .vasID, "F", .vasID.MaxRows, colPSEX
'            SetText .vasID, "55", .vasID.MaxRows, colPAGE
            
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
    
        GetSampleInfoW_SWMC = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
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
            
            SetText .vasID, "1", asRow, colCHECKBOX
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
                SetText frmInterface.vasID, "1", asRow, colCHECKBOX
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

Function GetOrderExamCode_UBCARE(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_UBCARE = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If

          SQL = "SELECT DiSTINCT EXAMCODE "
    SQL = SQL & vbCrLf & "  FROM PATRESULT "
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & argPID & "'"
'    SQL = SQL & vbCrLf & "   AND (SAVESEQ IS NULL OR SAVESEQ = '')"
    
    Set rs_svr = cn.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_UBCARE = GetOrderExamCode_UBCARE & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_UBCARE <> "" Then
        GetOrderExamCode_UBCARE = Mid(GetOrderExamCode_UBCARE, 1, Len(GetOrderExamCode_UBCARE) - 1)
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


