Attribute VB_Name = "modAU680"
Option Explicit
Public strETBYN         As String
Public strStart         As String
Public gstrOrder                    '/���� ����� �ֱ�
'/MSCOMM1 ���� ���
Public gstrOrderType    As String   '/�������� (Q : Request)
Public gintOrderNo      As Integer  '/������ȣ No (�迭)

Public gstrSampleType   As String   '/��üŸ��(�����ٶ� �ʿ���)


Public Function FUNC_HIS_ORDER_MAKE(argBCNO As String)
    Dim strEXSEQ    As String           '/SEQ �ű��
    
    Dim EXCD                            '/ó��˻��ڵ�(�迭)
    Dim EQCD                            '/���˻��ڵ�(�迭)
    
    Dim EXCD_ALL_LIST   As String       '/ó��˻��ڵ� ����Ʈ
    Dim EXCD_LIST       As String       '/ó��˻��ڵ� ����Ʈ
    Dim EQCD_LIST       As String       '/���˻��ڵ� ����Ʈ
    Dim strOrder        As String       '/���˻��ڵ�(��ȣ)
    
    
'/1. ���ڵ��ȣ�� �������� ��ȸ
'/2. �����ڵ� ����Ʈ�� ���˻��ڵ� ��ȸ
'/3. ���˻��ڵ� + �����ڵ�� LOCAL DB �� ����
'/4. �˻���� �����
    
    EXCD_ALL_LIST = ""
    EXCD_LIST = ""
    EQCD_LIST = ""
    strOrder = ""
'/1.�������� ��ȸ----------------------------------------------------------------------------------------------------/
    If ConnDB_LOC = False Then Exit Function
                           gstrQuy = "SELECT EXCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        Do Until ADR_LOC.EOF
            EXCD_ALL_LIST = EXCD_ALL_LIST & ",'" & Trim(ADR_LOC!EXCD & "") & "' "
            
            ADR_LOC.MoveNext
        Loop
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        EXCD_ALL_LIST = Mid(EXCD_ALL_LIST, 2)
    End If
    
    Call CloseDB_LOC
'/-------------------------------------------------------------------------------------------------------------------/
    
'/2.�������� ��ȸ----------------------------------------------------------------------------------------------------/
    If ConnDB_HIS = False Then Exit Function
        
    gstrQuy = "SELECT *  "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMRES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE SPECIMENID = '" & argBCNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EXAMCODE IN (" & EXCD_ALL_LIST & ") "
    gstrQuy = gstrQuy & vbCrLf & "   AND NVL(RESEND,' ') || NVL(EXAMSTATE,' ') <> '1D' "
    If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End
    If Not ADR_HIS Is Nothing Then
        
        Do Until ADR_HIS.EOF
            EXCD_LIST = EXCD_LIST & ",'" & Trim(ADR_HIS!EXAMCODE & "") & "'"
            ADR_HIS.MoveNext
        Loop
        
        ADR_HIS.Close: Set ADR_HIS = Nothing
        
        EXCD_LIST = Mid(EXCD_LIST, 2)
    End If
    
    If EXCD_LIST = "" Then: EXCD_LIST = "''"
    
    Call CloseDB_HIS
'/-------------------------------------------------------------------------------------------------------------------/

'/2.���˻��ڵ� ��ȸ------------------------------------------------------------------------------------------------/
    If ConnDB_LOC = False Then Exit Function
                           gstrQuy = "SELECT A.EQCD, B.EXCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST A , EX_MST B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQCD = B.EQCD "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.EQORDYN = 'Y' "
        gstrQuy = gstrQuy & vbCrLf & "   AND B.EXCD IN (" & EXCD_LIST & ") "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        EXCD_LIST = "" '/���θ����� ���� �ٽ� ���� -ȿ��-
        Do Until ADR_LOC.EOF
            EQCD_LIST = EQCD_LIST & "/" & Trim(ADR_LOC!EQCD & "")
            EXCD_LIST = EXCD_LIST & "/" & Trim(ADR_LOC!EXCD & "")
            ADR_LOC.MoveNext
        Loop
        
        
        ADR_LOC.Close: Set ADR_LOC = Nothing
    End If
        
    Call CloseDB_LOC

'/-------------------------------------------------------------------------------------------------------------------/

'/3.LOCAL DB �� ����-------------------------------------------------------------------------------------------------/
    If EQCD_LIST <> "" Then                                         '/�˻� ������ ������
        EQCD = Split(EQCD_LIST, "/")
        EXCD = Split(EXCD_LIST, "/")
        For intX = 1 To UBound(EQCD)
            gtypPAT_RES.EQCD = EQCD(intX)
            gtypPAT_RES.EXAMCD = EXCD(intX)
            gtypPAT_RES.STATEFLAG = "0"
            gtypPAT_RES.EXDT = Format(Now, "YYYYMMDD")
            gtypPAT_RES.EXTM = Format(Now, "HHMMSS")
            
            If strEXSEQ <> "Y" Then
                gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/��ü��ȣ(Barcode)�� �˻�ȸ��
                strEXSEQ = "Y"
            End If
            
            
            If FUNC_LOC_SAVE_PAT_RES = False Then: Exit Function
        Next intX
    Else                                                            '/�˻������ ������ ���ٴ� ��ȣ�� ����
        
    End If
'/-------------------------------------------------------------------------------------------------------------------/
'/4.�˻���� �����--------------------------------------------------------------------------------------------------/
    If EQCD_LIST <> "" Then                                         '/�˻� ������ ������
        gstrSampleType = " "
        EQCD = Split(EQCD_LIST, "/")
        For intX = 1 To UBound(EQCD)
            If EQCD(intX) = "100" Then gstrSampleType = "W"
            strOrder = strOrder & EQCD(intX) & "0"
        Next intX
        'gintOrderNo = gintOrderNo + 1
        gintOrderNo = 0
        '/Head         1H|\^&|||ASTM-Host(CR)59
        
        gstrOrder(gintOrderNo) = chrSTX & "S " & gtypPAT_RES.DISKNO & gtypPAT_RES.POSNO & gstrSampleType & gtypPAT_RES.SAMPLENO
        gstrOrder(gintOrderNo) = gstrOrder(gintOrderNo) & TEXT_RSET(gtypPAT_RES.BARCD, 20) & "    "
        gstrOrder(gintOrderNo) = gstrOrder(gintOrderNo) & "E" & strOrder & chrETX
        
        frmEQ_Main.MSComm1.Output = gstrOrder(gintOrderNo)
        SaveData "[Tx] : " & gstrOrder(gintOrderNo)
''        gstrOrder(0) = chrSTX & gstrOrder(0) & CheckSum(gstrOrder(0)) & vbCrLf
''
''        '/Patient       2P|1||200807250520(CR)96
''        gstrOrder(1) = "2P|1||" & gtypPAT_RES.BARCD & chrCR & chrETX
''        gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & vbCrLf
''
''        '/Order         3O|1|00000031|^4^3|^^^410^0\^^^900^0|R||||||N||||||||||||||O(CR)44
''        gstrOrder(2) = "3O|1|" & gtypPAT_RES.BARCD & "^" & gtypPAT_RES.SAMPLENO & "^" & gtypPAT_RES.DISKNO & "^" & gtypPAT_RES.POSNO & "^^SAMPLE^NORMAL|"
''        gstrOrder(2) = gstrOrder(2) & strOrder & "|R||||||N||||||||||||||O" & chrCR & chrETX
''        gstrOrder(2) = chrSTX & gstrOrder(2) & CheckSum(gstrOrder(2)) & vbCrLf
''
''        '/terminater    4L|1(CR)3D
''        gstrOrder(3) = "4L|1" & chrCR & chrETX
''        gstrOrder(3) = chrSTX & gstrOrder(3) & CheckSum(gstrOrder(3)) & vbCrLf
'
'        '/EOT           
'
'    Else                                                            '/�˻������ ������ ���ٴ� ��ȣ�� ����
'        ReDim gstrOrder(4)
'        '/Head         1H|\^&|||ASTM-Host(CR)59
'        gstrOrder(0) = "1H|\^&|||ASTM-Host" & chrCR & chrETX
'        gstrOrder(0) = chrSTX & gstrOrder(0) & CheckSum(gstrOrder(0)) & vbCrLf
'
'        '/Patient       2P|1||200807250520(CR)96
'        gstrOrder(1) = "2P|1||" & gtypPAT_RES.BARCD & chrCR & chrETX
'        gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & vbCrLf
'
'        '/Order         3O|1|00000031|^4^3|^^^410^0\^^^900^0|R||||||N||||||||||||||O(CR)44
'        gstrOrder(2) = "3O|1|" & gtypPAT_RES.BARCD & "^" & gtypPAT_RES.SAMPLENO & "^" & gtypPAT_RES.DISKNO & "^" & gtypPAT_RES.POSNO & "^^SAMPLE^NORMAL|"
'        gstrOrder(2) = gstrOrder(2) & "|R||||||N||||||||||||||O" & chrCR & chrETX
'        gstrOrder(2) = chrSTX & gstrOrder(2) & CheckSum(gstrOrder(2)) & vbCrLf
'
'        '/terminater    4L|1(CR)3D
'        gstrOrder(3) = "4L|1" & chrCR & chrETX
'        gstrOrder(3) = chrSTX & gstrOrder(3) & CheckSum(gstrOrder(3)) & vbCrLf
'
'        '/EOT           
'        gstrOrder(4) = chrEOT
    End If
'/-------------------------------------------------------------------------------------------------------------------/


'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    
End Function

Public Function FUNC_HIS_ORDER_VIEW() As Boolean
    Dim EXCD_LIST   As String
    
    FUNC_HIS_ORDER_VIEW = False
    gtypPAT_RES.EXAMCD = ""
    '/���� �׸�
    'gtypPAT_RES.EXAMCD  '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
    'gtypPAT_RES.ORDDT   '/ORDDT(ó������)
    'gtypPAT_RES.ORDGB   '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����)
On Error GoTo RTN_ERR
    
'/-------------------------------------------------------------------------------------------------------------------/
    
    If ConnDB_LOC = True Then
        '/����ڵ庰 ó���ڵ� ��������
        gstrQuy = "SELECT EXCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & gtypPAT_RES.EQCD & "' "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
        If Not ADR_LOC Is Nothing Then
            Do Until ADR_LOC.EOF
                EXCD_LIST = EXCD_LIST & ",'" & Trim(ADR_LOC!EXCD & "") & "'"
                ADR_LOC.MoveNext
            Loop
        End If
        
        Call CloseDB_LOC
        
        EXCD_LIST = Mid(EXCD_LIST, 2)
    End If
    
    If EXCD_LIST = "" Then: Exit Function

'/----------------------------------------------------------------------------------------------------/
    
'/2.�������� ��ȸ----------------------------------------------------------------------------------------------------/
    If ConnDB_HIS = True Then
                           gstrQuy = "SELECT *  "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMRES"
        gstrQuy = gstrQuy & vbCrLf & " WHERE SPECIMENID = '" & gtypPAT_RES.BARCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXAMCODE IN (" & EXCD_LIST & ") "
        gstrQuy = gstrQuy & vbCrLf & "   AND (NVL(RESEND,' ') || NVL(EXAMSTATE,' ')) <> '1D' "
        gstrQuy = gstrQuy & vbCrLf & "   AND LABRECYN = 'Y' "
        If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End
    
        If Not ADR_HIS Is Nothing Then
            gtypPAT_RES.EXAMCD = Trim(ADR_HIS!EXAMCODE & "")
            ADR_HIS.Close: Set ADR_HIS = Nothing
        End If
        
        Call CloseDB_HIS
    End If
'/----------------------------------------------------------------------------------------------------/

    FUNC_HIS_ORDER_VIEW = True
    Exit Function
RTN_ERR:
    
End Function

Public Function FUNC_HIS_PATIENT() As Boolean
    FUNC_HIS_PATIENT = False
    
On Error GoTo RTN_ERR
    
    '/Patient ID �� ���ڵ��� ��� ���Ϲ�ȣ�� ã�´�.
    gtypPAT_RES.PATNO = "" '/PATNO(���Ϲ�ȣ)

    '/����
    gtypPAT_RES.PATNM = ""  '/PATNM(�����ڸ�)
    gtypPAT_RES.PATSEX = "" '/PATSEX(����)
    gtypPAT_RES.PATAGE = "" '/PATAGE(����)
    
    '/���������� ������ ���� ��--------------------------------------------------/

    If ConnDB_HIS = False Then Exit Function
                           gstrQuy = "SELECT A.*  "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PATIENT A, EXAMRES B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID = B.PID "
        gstrQuy = gstrQuy & vbCrLf & "   AND  B.SPECIMENID = '" & gtypPAT_RES.BARCD & "' "
    If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End

    If Not ADR_HIS Is Nothing Then
        
        Do Until ADR_HIS.EOF
            gtypPAT_RES.PATNO = Trim(ADR_HIS!PID & "")
            gtypPAT_RES.PATNM = Trim(ADR_HIS!PNAME & "")    '/PATNM(�����ڸ�)
            gtypPAT_RES.PATSEX = Trim(ADR_HIS!SEX & "")  '/PATSEX(����)
            'gtypPAT_RES.PATAGE = Trim(ADR_HIS!PID & "")  '/PATAGE(����)
            ADR_HIS.MoveNext
        Loop
        
        ADR_HIS.Close: Set ADR_HIS = Nothing

    End If
    
    Call CloseDB_HIS
    
    FUNC_HIS_PATIENT = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    
End Function

Public Function FUNC_HIS_RESULT_JUDGMENT()
    Dim strPOINT        As String
    Dim strAHREF        As String
    Dim strALREF        As String
    
    Dim strPHREF        As String
    Dim strPLREF        As String
    
    Dim strDBEFORRSLT   As String
    Dim strDHREF        As String
    Dim strDLREF        As String
    Dim strDDATE        As String
    Dim strDGUBUN       As String
    Dim strDVALUE       As String
    
    Dim strStartDate    As String
    Dim strEndDate      As String
    
    gtypPAT_RES.AFLAG = "" '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
    gtypPAT_RES.PFLAG = "" '/PFLAG(Panic)
    gtypPAT_RES.DFLAG = "" '/DFLAG(Delta)
    
    strStartDate = Format(DateAdd("Y", Now, -30), "yyyy-mm-dd")
    strEndDate = Format(Now, "yyyy-mm-dd")
    
    '/���������� �ҷ�����--------------------------------------------------------------------------------------------
    If ConnDB_HIS = False Then Exit Function
    
    gstrQuy = ""
    gstrQuy = gstrQuy & vbCrLf & "SELECT VS_RESULT, VS_RECENO "
    gstrQuy = gstrQuy & vbCrLf & "  FROM V_EXAMRES"
    gstrQuy = gstrQuy & vbCrLf & " WHERE VS_REQDATE  >= '" & strStartDate & "'"
    gstrQuy = gstrQuy & vbCrLf & "   AND VS_REQDATE  <= '" & strEndDate & "'"
    gstrQuy = gstrQuy & vbCrLf & "   AND VS_RECENO < "
    gstrQuy = gstrQuy & vbCrLf & "              (SELECT MAX(VS_REQDATE)"
    gstrQuy = gstrQuy & vbCrLf & "                 FROM V_EXAMRES"
    gstrQuy = gstrQuy & vbCrLf & "                WHERE VS_PID = '" & gtypPAT_RES.PATNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "                  AND VS_EXAMCODE = '" & gtypPAT_RES.EXAMCD & "') "
    gstrQuy = gstrQuy & vbCrLf & "   AND VS_PID       = '" & gtypPAT_RES.PATNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND VS_EXAMCODE  = '" & gtypPAT_RES.EXAMCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND NVL(VS_RESULT, ' ') <> ' '"
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY VS_REQDATE DESC"
    If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End

    If Not ADR_HIS Is Nothing Then
        strDBEFORRSLT = Trim(ADR_HIS!VS_RESULT & "")
        ADR_HIS.Close: Set ADR_HIS = Nothing
    End If
    '/���������� �ҷ�����--------------------------------------------------------------------------------------------
    
    If gtypPAT_RES.PATSEX = "F" Then
        If ConnDB_HIS = False Then Exit Function
                               gstrQuy = "SELECT * " 'POINT, RES_M_HIGH, RES_M_LOW, DELTAHIGH, DELTALOW, DELTACAL, PANIC_M_HIGH, PANIC_M_LOW "
            gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMMASTER "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EXAMCODE = '" & gtypPAT_RES.EXAMCD & "' "
    
        If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End
    
        If Not ADR_HIS Is Nothing Then
            
            strPOINT = Trim(ADR_HIS!Point & "")
            strAHREF = Trim(ADR_HIS!RES_F_HIGH & "")
            strALREF = Trim(ADR_HIS!RES_F_LOW & "")
            
            strPHREF = Trim(ADR_HIS!PANIC_F_HIGH & "")
            strPLREF = Trim(ADR_HIS!PANIC_F_LOW & "")
            
            strDHREF = Trim(ADR_HIS!DELTAHIGH & "")
            strDLREF = Trim(ADR_HIS!DELTALOW & "")
            strDDATE = "30"
            strDGUBUN = Trim(ADR_HIS!DELTACAL & "")
            'ADR_HIS.MoveNext
            
            ADR_HIS.Close: Set ADR_HIS = Nothing
            
        End If
    Else
        If ConnDB_HIS = False Then Exit Function
                               gstrQuy = "SELECT POINT, RES_M_HIGH, RES_M_LOW, DELTAHIGH, DELTALOW, DELTACAL, PANIC_M_HIGH, PANIC_M_LOW "
            gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMMASTER "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EXAMCODE = '" & gtypPAT_RES.EXAMCD & "' "
    
        If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End
    
        If Not ADR_HIS Is Nothing Then
            
            strPOINT = Trim(ADR_HIS!Point & "")
            strAHREF = Trim(ADR_HIS!RES_M_HIGH & "")
            strALREF = Trim(ADR_HIS!RES_M_LOW & "")
            
            strPHREF = Trim(ADR_HIS!PANIC_M_HIGH & "")
            strPLREF = Trim(ADR_HIS!PANIC_M_LOW & "")
            
            strDHREF = Trim(ADR_HIS!DELTAHIGH & "")
            strDLREF = Trim(ADR_HIS!DELTALOW & "")
            strDDATE = "30"
            strDGUBUN = Trim(ADR_HIS!DELTACAL & "")
            'ADR_HIS.MoveNext
            
            ADR_HIS.Close: Set ADR_HIS = Nothing
            
        End If
    End If
    Call CloseDB_HIS
    
    If IsNumeric(Trim(gtypPAT_RES.EQRESULT & "")) = False Then Exit Function
    
    '/��� �Ҽ����ڸ� ����
    Select Case strPOINT
        Case "0": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0")
        Case "1": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0.0")
        Case "2": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0.#0")
        Case "3": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0.##0")
        Case "4": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0.###0")
        Case "5": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0.####0")
        Case "6": gtypPAT_RES.Result = Format(gtypPAT_RES.EQRESULT, "###0.#####0")
        Case Else
            gtypPAT_RES.Result = gtypPAT_RES.EQRESULT
    End Select
     '/--------------------------------------------------
     
     
    '/�˻��ڵ带 ���ҷ��ͼ� ������ ������ ��쿡�� �����õ����͸� �־���
    If gtypPAT_RES.EXAMCD = "" Then Exit Function
        
    '/Abnomal �Ǻ�
    If IsNumeric(Trim(strALREF)) = True And IsNumeric(Trim(strAHREF)) = True Then
        If CDbl(gtypPAT_RES.EQRESULT) < CDbl(strALREF) Then
            gtypPAT_RES.AFLAG = "L"
        ElseIf CDbl(gtypPAT_RES.EQRESULT) > CDbl(strAHREF) Then
            gtypPAT_RES.AFLAG = "H"
        End If
        
    ElseIf IsNumeric(Trim(strALREF)) = False And IsNumeric(Trim(strAHREF)) = True Then
        If CDbl(gtypPAT_RES.EQRESULT) > CDbl(strAHREF) Then
            gtypPAT_RES.AFLAG = "H"
        End If
        
    ElseIf IsNumeric(Trim(strALREF)) = True And IsNumeric(Trim(strAHREF)) = False Then
        If CDbl(gtypPAT_RES.EQRESULT) < CDbl(strALREF) Then
            gtypPAT_RES.AFLAG = "L"
        End If
                
    ElseIf IsNumeric(Trim(strALREF)) = False And IsNumeric(Trim(strAHREF)) = False Then
        
    End If
    '/--------------------------------------------------
    
    '/Delta �Ǻ� (����)
    If IsNumeric(gtypPAT_RES.EQRESULT) = True And IsNumeric(strDBEFORRSLT) = True Then
        Select Case strDGUBUN
            Case ""     '0 ������
                strDVALUE = ""
            Case "1"     '1 ��ȭ�� = ������ - �������"
                strDVALUE = ""
                strDVALUE = CDbl(gtypPAT_RES.EQRESULT) - CDbl(strDBEFORRSLT)                    '��ȭ��
            Case "2"     '2 ��ȭ���� = ��ȭ�� / ������� * 100"
                strDVALUE = ""
                strDVALUE = CDbl(gtypPAT_RES.EQRESULT) - CDbl(strDBEFORRSLT)                    '��ȭ��
                strDVALUE = (CDbl(strDVALUE) / CDbl(strDBEFORRSLT)) * 100               '��ȭ����
            Case "3"     '3 �Ⱓ�� ��ȭ���� = ��ȭ���� / �Ⱓ"
                strDVALUE = ""
                strDVALUE = CDbl(gtypPAT_RES.EQRESULT) - CDbl(strDBEFORRSLT)                    '��ȭ��
                strDVALUE = (CDbl(strDVALUE) / CDbl(strDBEFORRSLT)) * 100                '��ȭ����
                strDVALUE = strDVALUE / CCur(strDDATE)          '�Ⱓ�� ��ȭ����
            Case "4"     '4 �Ⱓ�� ��ȭ�� = ��ȭ�� / �Ⱓ"
                strDVALUE = ""
                strDVALUE = CDbl(gtypPAT_RES.EQRESULT) - CDbl(strDBEFORRSLT)                    '��ȭ��
                strDVALUE = CDbl(strDVALUE) / CCur(strDDATE)   '�Ⱓ�� ��ȭ��
            Case "5"     '�´��� Ȯ���ؾ��� -> 5 ���뺯ȭ���� = ��ȭ�� / �������"
                strDVALUE = ""
'                strDVALUE = CDbl(gtypPAT_RES.EQRESULT) - CDbl(strDBEFORRSLT)                    '��ȭ��
'                strDVALUE = CDbl(strDVALUE) / CDbl(strDBEFORRSLT)                        '���뺯ȭ����
            Case Else
                strDVALUE = ""
        End Select
        
        If IsNumeric(strDHREF) And IsNumeric(strDLREF) And IsNumeric(strDVALUE) Then
            If (CDbl(strDVALUE) > strDHREF Or CCur(strDVALUE) < strDLREF) Then
                gtypPAT_RES.DFLAG = "D"
            Else
                gtypPAT_RES.DFLAG = ""
            End If
        Else
            gtypPAT_RES.DFLAG = ""
        End If
    End If
    '/--------------------------------------------------
    
    '/Panic �Ǻ�
    If IsNumeric(Trim(strPLREF)) = True And IsNumeric(Trim(strPHREF)) = True Then
        If CDbl(gtypPAT_RES.EQRESULT) < CDbl(strPLREF) Then
            gtypPAT_RES.PFLAG = "P"
        ElseIf CDbl(gtypPAT_RES.EQRESULT) > CDbl(strPHREF) Then
            gtypPAT_RES.PFLAG = "P"
        End If
        
    ElseIf IsNumeric(Trim(strPLREF)) = False And IsNumeric(Trim(strPHREF)) = True Then
        If CDbl(gtypPAT_RES.EQRESULT) > CDbl(strPHREF) Then
            gtypPAT_RES.PFLAG = "P"
        End If
        
    ElseIf IsNumeric(Trim(strPLREF)) = True And IsNumeric(Trim(strPHREF)) = False Then
        If CDbl(gtypPAT_RES.EQRESULT) < CDbl(strPLREF) Then
            gtypPAT_RES.PFLAG = "P"
        End If
                
    ElseIf IsNumeric(Trim(strPLREF)) = False And IsNumeric(Trim(strPHREF)) = False Then
        
    End If
    '/--------------------------------------------------
    
End Function

Public Function FUNC_HIS_SAVE() As Boolean
    FUNC_HIS_SAVE = False
    
On Error GoTo RTN_ERR
    
    If ConnDB_HIS = False Then Exit Function
    
    ADC_HIS.BeginTrans
    
                           gstrQuy = "UPDATE EXAMRES "
        gstrQuy = gstrQuy & vbCrLf & "   SET RESULT = '" & gtypPAT_RES.Result & "' "
        gstrQuy = gstrQuy & vbCrLf & "       ,PANICFLAG = '" & gtypPAT_RES.PFLAG & "' "
        gstrQuy = gstrQuy & vbCrLf & "       ,DELTAFLAG = '" & gtypPAT_RES.DFLAG & "' "
        gstrQuy = gstrQuy & vbCrLf & "       ,DECISION = '" & gtypPAT_RES.AFLAG & "' "
        gstrQuy = gstrQuy & vbCrLf & "       ,EXAMUID = '" & gtypUSER.USERID & "' "
        gstrQuy = gstrQuy & vbCrLf & "       ,EXAMDATE = SYSDATE "
        gstrQuy = gstrQuy & vbCrLf & "       ,EXAMSTATE = "
        gstrQuy = gstrQuy & vbCrLf & "                 (CASE "
        'gstrQuy = gstrQuy & vbCrLf & "                  WHEN (SELECT NVL(EXAMSTATE,' ') FROM EXAMRES WHERE NVL(RESEND,' ')= '' AND NVL(EXAMSTATE,' ') = 'B' AND SPECIMENID = '" & gtypPAT_RES.BARCD & "') = 'B' "
        gstrQuy = gstrQuy & vbCrLf & "                  WHEN (NVL(EXAMSTATE,' ')  = 'B' OR NVL(EXAMSTATE,' ')  = ' ') "
        gstrQuy = gstrQuy & vbCrLf & "                       THEN 'D'"
        gstrQuy = gstrQuy & vbCrLf & "                  ELSE EXAMSTATE "
        gstrQuy = gstrQuy & vbCrLf & "                  END) "
        gstrQuy = gstrQuy & vbCrLf & " WHERE SPECIMENID = '" & gtypPAT_RES.BARCD & "'"
        gstrQuy = gstrQuy & vbCrLf & "   AND EXAMCODE = '" & gtypPAT_RES.EXAMCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND NVL(RESEND,' ') || NVL(EXAMSTATE,' ') <> '1D' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXAMSTATE <> 'Q' "
        gstrQuy = gstrQuy & vbCrLf & "   AND LABRECYN = 'Y' "
    If RunSQL_HIS(gstrQuy) = False Then ADC_HIS.RollbackTrans: Call CloseDB_HIS: Exit Function
    
    ADC_HIS.CommitTrans
    
    FUNC_HIS_SAVE = True
    gtypPAT_RES.SENDFLAG = "1"
Exit Function
    
'/----------------------------------------------------------------------------------------------------/
    
RTN_ERR:

End Function

Public Sub SaveData(ByVal ArgSQL As String, Optional argFlag As Integer = 0)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
        
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    Open App.Path & "\Log\" & Format(Date, "yyyy-mm-dd") & ".txt" For Append As FilNum
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & ArgSQL
    Close FilNum
End Sub


