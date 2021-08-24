Attribute VB_Name = "modK4500"
Option Explicit
Public strETBYN         As String
Public strStart         As String
Public gstrOrder                    '/���� ����� �ֱ�
'/MSCOMM1 ���� ���
Public gstrOrderType    As String   '/�������� (Q : Request)
Public gintOrderNo      As Integer  '/������ȣ No (�迭)


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
        gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMRES"
        gstrQuy = gstrQuy & vbCrLf & " WHERE SPECIMENID = '" & argBCNO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXAMCODE IN (" & EXCD_ALL_LIST & ") "
        gstrQuy = gstrQuy & vbCrLf & "   AND (NVL(RESEND,' ') <> '1' "
        gstrQuy = gstrQuy & vbCrLf & "        OR (RESEND = '1' AND EXAMSTATE = 'E'))"
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
            If strEXSEQ <> "Y" Then
                gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/��ü��ȣ(Barcode)�� �˻�ȸ��
                strEXSEQ = "Y"
            End If
            
            gtypPAT_RES.STATEFLAG = "0"
            gtypPAT_RES.EXDT = Format(Now, "YYYYMMDD")
            gtypPAT_RES.EXTM = Format(Now, "HHMMSS")
            If FUNC_LOC_SAVE_PAT_RES = False Then: Exit Function
        Next intX
    Else                                                            '/�˻������ ������ ���ٴ� ��ȣ�� ����
        
    End If
'/-------------------------------------------------------------------------------------------------------------------/

'/4.�˻���� �����--------------------------------------------------------------------------------------------------/
    If EQCD_LIST <> "" Then                                         '/�˻� ������ ������
        EQCD = Split(EQCD_LIST, "/")
        For intX = 1 To UBound(EQCD)
            strOrder = strOrder & "^^^" & EQCD(intX) & "^0\"
        Next intX
        
        strOrder = Mid(strOrder, 1, Len(strOrder) - 1)
        
        ReDim gstrOrder(4)
        '/Head         1H|\^&|||ASTM-Host(CR)59
        gstrOrder(0) = "1H|\^&|||ASTM-Host" & chrCR & chrETX
        gstrOrder(0) = chrSTX & gstrOrder(0) & CheckSum(gstrOrder(0)) & vbCrLf
        
        '/Patient       2P|1||200807250520(CR)96
        gstrOrder(1) = "2P|1||" & gtypPAT_RES.BARCD & chrCR & chrETX
        gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & vbCrLf
        
        '/Order         3O|1|00000031|^4^3|^^^410^0\^^^900^0|R||||||N||||||||||||||O(CR)44
        gstrOrder(2) = "3O|1|" & gtypPAT_RES.BARCD & "^" & gtypPAT_RES.SAMPLENO & "^" & gtypPAT_RES.DISKNO & "^" & gtypPAT_RES.POSNO & "^^SAMPLE^NORMAL|"
        gstrOrder(2) = gstrOrder(2) & strOrder & "|R||||||N||||||||||||||O" & chrCR & chrETX
        gstrOrder(2) = chrSTX & gstrOrder(2) & CheckSum(gstrOrder(2)) & vbCrLf
        
        '/terminater    4L|1(CR)3D
        gstrOrder(3) = "4L|1" & chrCR & chrETX
        gstrOrder(3) = chrSTX & gstrOrder(3) & CheckSum(gstrOrder(3)) & vbCrLf
        
        '/EOT           
        gstrOrder(4) = chrEOT
        
    Else                                                            '/�˻������ ������ ���ٴ� ��ȣ�� ����
        ReDim gstrOrder(4)
        '/Head         1H|\^&|||ASTM-Host(CR)59
        gstrOrder(0) = "1H|\^&|||ASTM-Host" & chrCR & chrETX
        gstrOrder(0) = chrSTX & gstrOrder(0) & CheckSum(gstrOrder(0)) & vbCrLf
        
        '/Patient       2P|1||200807250520(CR)96
        gstrOrder(1) = "2P|1||" & gtypPAT_RES.BARCD & chrCR & chrETX
        gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & vbCrLf
        
        '/Order         3O|1|00000031|^4^3|^^^410^0\^^^900^0|R||||||N||||||||||||||O(CR)44
        gstrOrder(2) = "3O|1|" & gtypPAT_RES.BARCD & "^" & gtypPAT_RES.SAMPLENO & "^" & gtypPAT_RES.DISKNO & "^" & gtypPAT_RES.POSNO & "^^SAMPLE^NORMAL|"
        gstrOrder(2) = gstrOrder(2) & "|R||||||N||||||||||||||O" & chrCR & chrETX
        gstrOrder(2) = chrSTX & gstrOrder(2) & CheckSum(gstrOrder(2)) & vbCrLf
        
        '/terminater    4L|1(CR)3D
        gstrOrder(3) = "4L|1" & chrCR & chrETX
        gstrOrder(3) = chrSTX & gstrOrder(3) & CheckSum(gstrOrder(3)) & vbCrLf
        
        '/EOT           
        gstrOrder(4) = chrEOT
    End If
'/-------------------------------------------------------------------------------------------------------------------/


'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    
End Function

Public Function FUNC_HIS_ORDER_VIEW() As Boolean
    Dim EXCD_LIST   As String
    
    FUNC_HIS_ORDER_VIEW = False
    
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
        gstrQuy = gstrQuy & vbCrLf & "   AND (NVL(RESEND,' ') || NVL(EXAMSTATE,' ') <> '1D') "
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
    
    gtypPAT_RES.AFLAG = "" '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
    gtypPAT_RES.PFLAG = "" '/PFLAG(Panic)
    gtypPAT_RES.DFLAG = "" '/DFLAG(Delta)
    
    
    If gtypPAT_RES.PATSEX = "F" Then
        If ConnDB_HIS = False Then Exit Function
                               gstrQuy = "SELECT POINT, RES_F_HIGH, RES_F_LOW, DELTAHIGH, DELTALOW, PANIC_F_HIGH, PANIC_F_LOW "
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
            
            'ADR_HIS.MoveNext
            
            ADR_HIS.Close: Set ADR_HIS = Nothing
            
        End If
    Else
        If ConnDB_HIS = False Then Exit Function
                               gstrQuy = "SELECT POINT, RES_M_HIGH, RES_M_LOW, DELTAHIGH, DELTALOW, PANIC_M_HIGH, PANIC_M_LOW "
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
            
            'ADR_HIS.MoveNext
            
            ADR_HIS.Close: Set ADR_HIS = Nothing
            
        End If
    End If
    Call CloseDB_HIS
    
    If IsNumeric(Trim(gtypPAT_RES.EQRESULT & "")) = False Then gtypPAT_RES.Result = gtypPAT_RES.EQRESULT: Exit Function
    
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
'    If IsNumeric(Trim(strALREF)) = True And IsNumeric(Trim(strAHREF)) = True Then
'        If gtypPAT_RES.EQRESULT < strALREF Then
'            gtypPAT_RES.AFLAG = "L"
'        ElseIf gtypPAT_RES.EQRESULT > strAHREF Then
'            gtypPAT_RES.AFLAG = "H"
'        End If
'
'    ElseIf IsNumeric(Trim(strALREF)) = "" And IsNumeric(Trim(strAHREF)) = True Then
'        If gtypPAT_RES.EQRESULT > strAHREF Then
'            gtypPAT_RES.AFLAG = "H"
'        End If
'
'    ElseIf IsNumeric(Trim(strALREF)) = True And IsNumeric(Trim(strAHREF)) = "" Then
'        If gtypPAT_RES.EQRESULT < strALREF Then
'            gtypPAT_RES.AFLAG = "L"
'        End If
'
'    ElseIf IsNumeric(Trim(strALREF)) = "" And IsNumeric(Trim(strAHREF)) = "" Then
'
'    End If
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
        'gstrQuy = gstrQuy & vbCrLf & "                  WHEN NVL(EXAMSTATE,' ')  = 'B' "
        gstrQuy = gstrQuy & vbCrLf & "                  WHEN (NVL(EXAMSTATE,' ')  = 'B' OR NVL(EXAMSTATE,' ')  = ' ') "
        gstrQuy = gstrQuy & vbCrLf & "                       THEN 'D'"
        gstrQuy = gstrQuy & vbCrLf & "                  ELSE EXAMSTATE "
        gstrQuy = gstrQuy & vbCrLf & "                  END) "
        gstrQuy = gstrQuy & vbCrLf & " WHERE SPECIMENID = '" & gtypPAT_RES.BARCD & "'"
        gstrQuy = gstrQuy & vbCrLf & "   AND EXAMCODE = '" & gtypPAT_RES.EXAMCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND NVL(RESEND,' ') || NVL(EXAMSTATE,' ') <> '1D' "
        gstrQuy = gstrQuy & vbCrLf & "   AND LABRECYN = 'Y' "
    If RunSQL_HIS(gstrQuy) = False Then ADC_HIS.RollbackTrans: Call CloseDB_HIS: Exit Function
    
    ADC_HIS.CommitTrans
    Call CloseDB_HIS
    FUNC_HIS_SAVE = True
    gtypPAT_RES.SENDFLAG = "1"
Exit Function
    
'/----------------------------------------------------------------------------------------------------/
    
RTN_ERR:

End Function


Public Function FUNC_HIS_SAVE_MANUAL(argRow As Integer)
    
On Error GoTo RTN_ERR
    If ConnDB_LOC = False Then Exit Function
                       gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD     = '" & gtypPAT_RES.BARCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ     =  " & Val(gtypPAT_RES.EXSEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO  = '" & gtypPAT_RES.SAMPLENO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO    = '" & gtypPAT_RES.DISKNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND POSNO     = '" & gtypPAT_RES.POSNO & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        Do Until ADR_LOC.EOF
            gtypPAT_RES.Result = Trim(ADR_LOC!Result & "")      '/���
            gtypPAT_RES.EQCD = Trim(ADR_LOC!EQCD & "")          '/���˻��ڵ�
            gtypPAT_RES.EXAMCD = Trim(ADR_LOC!EXAMCD & "")            '/ó��˻��ڵ�
            gtypPAT_RES.DFLAG = Trim(ADR_LOC!DFLAG & "")
            gtypPAT_RES.PFLAG = Trim(ADR_LOC!PFLAG & "")
            gtypPAT_RES.AFLAG = Trim(ADR_LOC!AFLAG & "")

            If FUNC_HIS_SAVE = True Then '/HIS�� ��� ����
                Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS�� ��� ����
                Call SET_CELL(frmEQ_�˻�������.sprLResult, 8, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "�Ϸ�", "���"))
            End If
            
            ADR_LOC.MoveNext
        Loop
        
    End If
   
    Call CloseDB_LOC
Exit Function
    
'/----------------------------------------------------------------------------------------------------/
    
RTN_ERR:

End Function


Public Function NOVA_SAVE(argBARCD As String, argRow As Integer) As Boolean
    '/�� �˻� �� ��Ī�� ���
    Dim intCol      As Integer
    Dim strEXSEQ    As String
    NOVA_SAVE = False
With frmEQ_Main
    For intCol = gintEQ_StartCol To .sprLResult.MaxCols
        If GET_CELL(.sprLResult, intCol, argRow) <> "" Then
            gtypPAT_RES.BARCD = argBARCD
            gtypPAT_RES.EQCD = GET_CELL(.sprLResult, intCol, -1000)
            gtypPAT_RES.EQRESULT = GET_CELL(.sprLResult, intCol, argRow)
            If IsNumeric(gtypPAT_RES.EQRESULT) = True Then
                Call FUNC_HIS_RESULT_JUDGMENT '/��� ����
            Else
                gtypPAT_RES.Result = gtypPAT_RES.EQRESULT
            End If
            
            'gtypPAT_RES.Result = GET_CELL(.sprLResult, intCol, argROW)
            
            gtypPAT_RES.EXDT = Format(Now, "YYYYMMDD")
            gtypPAT_RES.EXTM = Format(Now, "HHMMSS")
            gtypPAT_RES.RCDT = Format(Now, "YYYYMMDD")      '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
            gtypPAT_RES.RCTM = Format(Now, "HHMMSS")        '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
            gtypPAT_RES.STATEFLAG = "1"                     '/STATEFLAG(���������� (0:ó��, 1:���))
            gtypPAT_RES.SENDFLAG = "0"
            '/�˻�SEQ ã��
            If strEXSEQ <> "Y" Then
                gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/��ü��ȣ(Barcode)�� �˻�ȸ��
                strEXSEQ = "Y"
            End If
            Call FUNC_HIS_PATIENT
            Call FUNC_HIS_ORDER_VIEW
            If gtypPAT_RES.EXAMCD <> "" Then
                If FUNC_LOC_SAVE_PAT_RES = True Then
                    If .mnuJobModeAuto.Checked = True And gtypPAT_RES.BARCD <> "" Then   '/���۹���� �ڵ������̸�...
                        If FUNC_HIS_SAVE = True Then '/HIS�� ��� ����
                        Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS�� ��� ����
                        Call SET_CELL(.sprLResult, 7, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "�Ϸ�", "���"))
                        End If
                        
                    End If
                End If
                
                gtypPAT_RES.EXAMCD = ""
            Else
                Call SET_CELL(.sprLResult, intCol, argRow, "")
                
            End If
        End If
    Next intCol
    gtypPAT_RES.BARCD = ""
    gtypPAT_RES.EXSEQ = ""            '/EXSEQ(��ü��ȣ(Barcode)�� �˻�ȸ��)
    gtypPAT_RES.EQCD = ""             '/EQCD(���˻��ڵ�)
    gtypPAT_RES.EXAMCD = ""           '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
    gtypPAT_RES.EXDT = ""             '/EXDT(�˻�ó����������(YYYYMMDD) HIEQ->�Ƿ����)
    gtypPAT_RES.EXTM = ""             '/EXTM(�˻�ó�����۽ð�(24HHMMSS) HIEQ->�Ƿ����)
    gtypPAT_RES.RCDT = ""             '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
    gtypPAT_RES.RCTM = ""             '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
    gtypPAT_RES.SDDT = ""             '/SDDT(�˻�����������(YYYYMMDD) HIEQ->HIS)
    gtypPAT_RES.SDTM = ""             '/SDTM(�˻������۽ð�(24HHMMSS) HIEQ->HIS)
    gtypPAT_RES.Result = ""           '/RESULT(�˻���(������ ���))
    gtypPAT_RES.EQRESULT = ""         '/EQRESULT(�����ð��)
    gtypPAT_RES.AFLAG = ""            '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
    gtypPAT_RES.PFLAG = ""            '/PFLAG(Panic)
    gtypPAT_RES.DFLAG = ""            '/DFLAG(Delta)
    gtypPAT_RES.SAMPLENO = ""         '/Sample No(AU2700, Uriscan � ���)
    gtypPAT_RES.DISKNO = ""           '/DISKNO(��ũ��ȣ or ����ȣ)
    gtypPAT_RES.POSNO = ""            '/POSNO(��ġ��ȣ)
    gtypPAT_RES.ORDDT = ""            '/ORDDT(ó������)
    gtypPAT_RES.ORDGB = ""            '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))
    gtypPAT_RES.PATNO = ""            '/PATNO(���Ϲ�ȣ)
    gtypPAT_RES.PATNM = ""            '/PATNM(�����ڸ�)
    gtypPAT_RES.PATSEX = ""           '/PATSEX(����)
    gtypPAT_RES.PATAGE = ""           '/PATAGE(����)
    gtypPAT_RES.SENDFLAG = ""         '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
    gtypPAT_RES.STATEFLAG = ""        '/STATEFLAG(���������� (0:ó��, 1:���))
    
End With
    NOVA_SAVE = True
End Function
