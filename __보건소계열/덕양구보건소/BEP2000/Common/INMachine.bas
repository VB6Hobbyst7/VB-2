Attribute VB_Name = "ModuleMachine"
Option Explicit

Public MDIactivekey                 As Integer  'MDIform�� �̹� load �Ǿ� �ִ� ���¸� ��Ÿ���� Ű
Public OrderKey                     As Integer  '��������� �˻�� Order�� ������ ȭ��� ��� ����
Public CallLabKey                   As Integer  'Lab Manager Program�� ȣ���� ���ΰ��� ���θ� ��Ÿ���� Ű
Public SlipDigit                    As Integer  'Slip�� �� �ڸ��� �� ���ΰ� ����
Public DigitShape                   As String   'Slip�� �ڸ��� ���� ������ ���
Public FieldAddIdenTBFlag           As Integer  'Field�� IdenTB�� �߰��� ������ �ϸ� �󸶳� �� ������ ����
Public FieldAddResultTBFlag         As Integer  'Field�� ResultTB�� �߰��� ������ �ϸ� �󸶳� �� ������ ����
Public IdTBNField()
Public IdTBFieldDig()
Public IdTBFieldName()
Public RTBNField()
Public RTBFieldDig()
Public RTBFieldName()

Sub FieldADD(IdTBno As Integer, RTBno As Integer)
    
    FieldAddIdenTBFlag = IdTBno
    FieldAddResultTBFlag = RTBno
    
    If FieldAddIdenTBFlag <> 0 Then
        ReDim IdTBNField(1 To FieldAddIdenTBFlag)
        ReDim IdTBFieldDig(1 To FieldAddIdenTBFlag)
        ReDim IdTBFieldName(1 To FieldAddIdenTBFlag)
    End If
    
    If FieldAddResultTBFlag <> 0 Then
        ReDim RTBNField(1 To FieldAddResultTBFlag)
        ReDim RTBFieldDig(1 To FieldAddResultTBFlag)
        ReDim RTBFieldName(1 To FieldAddResultTBFlag)
    End If
    
    If FieldAddIdenTBFlag <> 0 Then
        IdTBNField(1) = "timeNF"
        ''
        ''

    End If
    
    If FieldAddIdenTBFlag <> 0 Then
        IdTBFieldDig(1) = 6
        
        ''
        ''

    End If

    If FieldAddIdenTBFlag <> 0 Then
        IdTBFieldName(1) = "Testtime"
        
        ''
        ''

    End If
    
        
    If FieldAddResultTBFlag <> 0 Then
        RTBNField(1) = ""
        RTBNField(2) = ""
        ''
        ''

    End If
    
    If FieldAddResultTBFlag <> 0 Then
        RTBFieldDig(1) = ""
        RTBFieldDig(2) = ""
        ''
        ''

    End If

    If FieldAddResultTBFlag <> 0 Then
        RTBFieldName(1) = ""
        RTBFieldName(2) = ""
        ''
        ''

    End If

    
End Sub

Public Sub MachineConfig()
    
    Dim RetVal As Long
    Dim sBuf As String
    Dim CallLab As String
    
    Set interfacfrm = INTface41 'INTface41 ���� ��ü�ϴ� �̸�
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "MachineName", "Machine", sBuf, 255, App.Path & "\initial.ini")
    machstr = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "Title", "�˻���", sBuf, 255, App.Path & "\initial.ini")
    Title = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "LabManagerCall", "No", sBuf, 255, App.Path & "\initial.ini")
    CallLab = Left(sBuf, RetVal)
    
    '99.12.20  YEJ �߰�
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "Data_Cut", "0", sBuf, 255, App.Path & "\initial.ini")
    p_Data_Cut = Left(sBuf, RetVal)
    '--------------------------
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "ImportPath", "No", sBuf, 255, App.Path & "\initial.ini")
    ImportPath = Left(sBuf, RetVal)
    
    sBuf = String(255, 0)
    RetVal = GetPrivateProfileString("Path", "ExportPath", "No", sBuf, 255, App.Path & "\initial.ini")
    ExportPath = Left(sBuf, RetVal)
    
    machinit = Left(machstr, 3) & "_"   ''"MIN_" '�ӽ� �̴ϼ��� ���ڷ�, ��� �� �����Ͽ�
                      '�ִ� 4�ڷ� ��
    
    fileInit = Left(machstr, 3)     '"MIN" '���ϸ� �̴ϼȷ�, ���� ���� machinit����
                     '��� �ٸ� ���� ��
    
        
    INTmain00.Caption = Title & " " & " " & "�������̽� �ʱ� ȭ��"
      
    filename = App.Path & "\" '& machstr & "\"
    commstr = "clinic\setcomm.mdb"
    codestr = "clinic\setcode.mdb"
        
    'Call delcheck(filename & "comm\", machstr)

'MDIactivekey�� MDIform�� �̹� Load �Ǿ� �ִ� ��(= true)�� ��Ÿ���� ����
'������ �ٸ� app.���� �ٽ� ȣ���Ͽ� Load �Ǵ� ���� ������ �� ����.       '
    
    ImgClickkey = False
    MDIactivekey = True

'######################## CONFIGURATION #######################################################################################

'Slip�� �� �ڸ��� �� ���ΰ��� ����, �ڸ��� ���� ������ ��� ����t
    SlipDigit = 4
    Call SlipDigitShape

'Order ���ο� ���� ���� - Order used(True), Not used Order(False)
    OrderKey = False
    
'Lab Manager Program�� ȣ���� ���̰��� ���� ���� - ȣ���(True), ��ȣ���(False)
    If CallLab = "Yes" Then
        CallLabKey = True
    Else
        CallLabKey = False
    End If
    
'Column�� �߰��� ���ΰ��� ���� �� �� ���� Column�� �߰��� ���ΰ� ����(0:�߰�����, 1: 1Col�߰�, 2: 2Col�߰�, ...., n: nCol�߰�)
    ''Call FieldADD(1, 0)
    
End Sub

'
'   Bar-Code�κ��� �Էµ� ID�� �˻��׸� ��ȸ �� ����,�۽�
'
'Sub Send_Order()
'
'    Dim Tmp   As String
'    Dim ChkS  As String
'    Dim TestDat As String       '�˻��׸�
'    Dim i   As Integer
'    Dim K   As Integer
'    Dim tData() As String
'    Dim sStr    As String
'
'    Dim iRet    As Integer
'    Dim spSampNo    As Variant
'    Dim spDate  As Variant
'    Dim spSlip  As Variant
'    Dim spID    As Variant
'    Dim spTray, spCup   As Variant
'    Dim strTray As String * 2
'    Dim strCup  As String * 2
'
'    Dim sSql    As String
'    Dim SqlData As Recordset
'
'    If FrameN > 7 Then
'        FrameN = 0
'    End If
'
'    '���̻� ���� Order�� ������ ENQ ������(����ޱ� �غ����)��...
'    If Chk_SpdColor <> True Then
'        phase = 2
'        Exit Sub
'    End If
'
'    '----- ������ Order�� �����ڵ� ����
'    For K = 1 To spdWork.MaxRows
'        spdWork.Col = 1
'        spdWork.Row = K
'        If spdWork.BackColor = &HFFFFFF Then    'white
'            iRet = spdWork.GetText(1, K, spSampNo)  'Sample No
'            iRet = spdWork.GetText(4, K, spDate)    '�۾�����
'            iRet = spdWork.GetText(9, K, spSlip)    'Slip Code
'            iRet = spdWork.GetText(5, K, spID)      '�۾���ȣ
'            iRet = spdWork.GetText(7, K, spTray)    'Tray
'            iRet = spdWork.GetText(8, K, spCup)     'Cup
'            Exit For
'        End If
'    Next K
'
'    Select Case Snd_Phase
'        Case 1      'Header Record
'            Tmp = FrameN & "H|\^|||LL^REV  4.500|||||||||" & Format$(Now, "YYYYMMDD") & Format$(Now, "HHMMSS") & Chr(13) & Chr(3)
'            Snd_Phase = 2
'
'        Case 2      'Patient Record
'            'Tmp = FrameN & "P|  1" & Chr(13) & Chr(3)
'            Tmp = FrameN & "P|" & Format(SeqPNo, "@@@") & Chr(13) & Chr(3)      '9/24
'            SeqPNo = SeqPNo + 1
'            Snd_Phase = 3
'
'        Case 3      'Order Record
'            TestDat = ""
'            '===== �Ϲ�/��� �˻��׸� ��ȸ
'            If chkRerun = False Then
'                '--- Q.C �Ǵ� �Ϲ� Sample�� ���
'                If Mid(Trim(spSampNo), 5, 1) = "9" Then        'Q.C �� ���
'                    TestDat = "^^^01A\^^^01B\^^^02A\^^^04A"
'
'                Else                        '�Ϲ� Sample �� ���
'                    '----- �˻��׸� ��ȸ
'                    SqlStr = "        Select ORDCD "
'                    SqlStr = SqlStr & " from LAB01_DB..SLB020M "
'                    SqlStr = SqlStr & "where LABDATE = '" & spDate & "'"
'                    SqlStr = SqlStr & "  and SLIPCD  = '" & spSlip & "'"
'                    SqlStr = SqlStr & "  and LABSQNO = '" & spID & "'"
'
'                    Return_cd = QSqlDBExec(SqlStr, QsqlConn)
'                    If Return_cd <> QSQL_SUCCESS Then
'                        Return_cd = QSqlSelectFree(QsqlConn)
'                        Exit Sub
'                    End If
'                    Do
'                        Return_cd = QSqlGetRow(sStr, QsqlConn)
'                        If Return_cd <> QSQL_SUCCESS Then
'                            Exit Do
'                        End If
'
'                        QSqlGetField 1, sStr, tData()
'
'                        For i = 1 To MAX_NUM
'                            If Trim(Left$(TestNameTable(i).code, 8)) = Trim$(tData(1)) Then
'                                If TestDat = "" Then
'                                    TestDat = "^^^" & Trim$(TestNameTable(i).EqCd)
'                                Else
'                                    TestDat = TestDat & "\^^^" & Trim$(TestNameTable(i).EqCd)
'                                End If
'                            End If
'                        Next i
'
'                    Loop Until (Return_cd = QSQL_ERROR)
'                    Return_cd = QSqlSelectFree(QsqlConn)
'                End If
'
'            Else            '���
'                '----- MDB �κ��� �˻��׸� ��ȸ
'                sSql = "   SELECT EqCode " _
'                        & "  FROM TbRerun " _
'                        & " WHERE Lab_ID = '" & Trim(spDate) & Trim(spSlip) & Trim(spID) & "'" _
'                        & "   AND RerunChk = '" & 1 & "'" _
'                        & " ORDER BY EqCode "
'                Set SqlData = Db.OpenRecordset(sSql, dbOpenDynaset)
'
'                If SqlData.RecordCount <> 0 Then
'                    SqlData.MoveFirst
'
'                    Do While (SqlData.EOF = False)
'                        If TestDat = "" Then
'                            TestDat = "^^^" & Trim$(SqlData!EqCode)
'                        Else
'                            TestDat = TestDat & "\^^^" & Trim$(SqlData!EqCode)
'                        End If
'                        SqlData.MoveNext
'                    Loop
'                End If
'                SqlData.Close
'            End If
'            '-------------------
'
'            strTray = " " & spTray
'            strCup = Format$(spCup, "@@")
'
'            Tmp = FrameN & "O|  1|" & spSampNo & "^" & strTray & "^" & strCup & "||"
'            Tmp = Tmp & TestDat & "|R||||||N||||S|||||||" & Format$(Now, "YYYYMMDD") & Format$(Now, "HHMMSS") & Chr(13) & Chr(3)   '??? Sample Type(S/U/O)
'            'Tmp = Tmp & TestDat & "|R||||||N||||S" & Chr(13) & Chr(3)    '??? Sample Type(S/U/O)
'            Snd_Phase = 4
'
'        Case 4      'Terminator Record
'            Tmp = FrameN & "L|  1" & Chr(13) & Chr(3)
'            Snd_Phase = 5
'
'        Case 5      'EOT
'            '----- Order ���ۿϷ�� Sample�� BackColor����
'            spdWork.Col = 1: spdWork.Col2 = 9
'            spdWork.Row = K: spdWork.Row2 = K
'            spdWork.BlockMode = True
'            spdWork.BackColorStyle = BackColorStyleUnderGrid
'            spdWork.BackColor = &HC0FFC0    '���λ�
'            spdWork.BlockMode = False
'
'            '----- �� ������ Order�� �ִ��� �˻�
'            If Chk_SpdColor() <> True Then
'                Comm1.Output = Chr(4)   'EOT
'                FrameN = 1
'                phase = 2
'                Snd_Phase = 1
'                cmdOrder.Enabled = True: cmdClear.Enabled = True
'                For i = 0 To 3
'                    cmdQC(i).Enabled = True
'                Next i
'
'                Exit Sub
'            Else
'                Tmp = FrameN & "H|\^|||LL^REV  4.500|||||||||" & Format$(Now, "YYYYMMDD") & Format$(Now, "HHMMSS") & Chr(13) & Chr(3)
'                Snd_Phase = 2
'            End If
'
'    End Select
'
'    Print #2, "<S> " & Tmp & Chr(13) & Chr(10);
'
'    ChkS = Chk_Sum(Tmp)
'    Comm1.Output = Chr(2) & Tmp & ChkS & Chr(13) & Chr(10)
'    FrameN = FrameN + 1
'
'End Sub

Sub SlipDigitShape()
    Select Case SlipDigit
    
        Case 4
            DigitShape = "0000"
        Case 5
            DigitShape = "00000"
        Case 6
            DigitShape = "000000"
        Case 7
            DigitShape = "0000000"
        Case 8
            DigitShape = "00000000"
        Case 9
            DigitShape = "000000000"
        Case 10
            DigitShape = "0000000000"
        Case 11
            DigitShape = "00000000000"
        Case 12
            DigitShape = "000000000000"
        Case 13
            DigitShape = "0000000000000"
        Case 14
            DigitShape = "00000000000000"
    
    End Select
    
    
End Sub


