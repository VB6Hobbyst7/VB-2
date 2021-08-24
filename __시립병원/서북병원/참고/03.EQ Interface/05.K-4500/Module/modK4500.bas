Attribute VB_Name = "modK4500"
Option Explicit
Public strETBYN         As String
Public strStart         As String
Public gstrOrder                    '/오더 만든거 넣기
'/MSCOMM1 에서 사용
Public gstrOrderType    As String   '/오더구분 (Q : Request)
Public gintOrderNo      As Integer  '/오더신호 No (배열)


Public Function FUNC_HIS_ORDER_MAKE(argBCNO As String)
    Dim strEXSEQ    As String           '/SEQ 매기기
    
    Dim EXCD                            '/처방검사코드(배열)
    Dim EQCD                            '/장비검사코드(배열)
    
    Dim EXCD_ALL_LIST   As String       '/처방검사코드 리스트
    Dim EXCD_LIST       As String       '/처방검사코드 리스트
    Dim EQCD_LIST       As String       '/장비검사코드 리스트
    Dim strOrder        As String       '/장비검사코드(신호)
    
    
'/1. 바코드번호로 오더정보 조회
'/2. 오더코드 리스트로 장비검사코드 조회
'/3. 장비검사코드 + 오더코드로 LOCAL DB 에 저장
'/4. 검사오더 만들기
    
    EXCD_ALL_LIST = ""
    EXCD_LIST = ""
    EQCD_LIST = ""
    strOrder = ""
'/1.오더정보 조회----------------------------------------------------------------------------------------------------/
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
    
'/2.오더정보 조회----------------------------------------------------------------------------------------------------/
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

'/2.장비검사코드 조회------------------------------------------------------------------------------------------------/
    If ConnDB_LOC = False Then Exit Function
                           gstrQuy = "SELECT A.EQCD, B.EXCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST A , EX_MST B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQCD = B.EQCD "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.EQORDYN = 'Y' "
        gstrQuy = gstrQuy & vbCrLf & "   AND B.EXCD IN (" & EXCD_LIST & ") "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        EXCD_LIST = "" '/새로만들지 말고 다시 쓰자 -효준-
        Do Until ADR_LOC.EOF
            EQCD_LIST = EQCD_LIST & "/" & Trim(ADR_LOC!EQCD & "")
            EXCD_LIST = EXCD_LIST & "/" & Trim(ADR_LOC!EXCD & "")
            ADR_LOC.MoveNext
        Loop
        
        
        ADR_LOC.Close: Set ADR_LOC = Nothing
    End If
        
    Call CloseDB_LOC

'/-------------------------------------------------------------------------------------------------------------------/

'/3.LOCAL DB 에 저장-------------------------------------------------------------------------------------------------/
    If EQCD_LIST <> "" Then                                         '/검사 오더가 있을때
        EQCD = Split(EQCD_LIST, "/")
        EXCD = Split(EXCD_LIST, "/")
        For intX = 1 To UBound(EQCD)
            gtypPAT_RES.EQCD = EQCD(intX)
            gtypPAT_RES.EXAMCD = EXCD(intX)
            If strEXSEQ <> "Y" Then
                gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/검체번호(Barcode)별 검사회차
                strEXSEQ = "Y"
            End If
            
            gtypPAT_RES.STATEFLAG = "0"
            gtypPAT_RES.EXDT = Format(Now, "YYYYMMDD")
            gtypPAT_RES.EXTM = Format(Now, "HHMMSS")
            If FUNC_LOC_SAVE_PAT_RES = False Then: Exit Function
        Next intX
    Else                                                            '/검사오더가 없을때 없다는 신호를 보냄
        
    End If
'/-------------------------------------------------------------------------------------------------------------------/

'/4.검사오더 만들기--------------------------------------------------------------------------------------------------/
    If EQCD_LIST <> "" Then                                         '/검사 오더가 있을때
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
        
    Else                                                            '/검사오더가 없을때 없다는 신호를 보냄
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
    
    '/정의 항목
    'gtypPAT_RES.EXAMCD  '/EXAMCD(처방코드(HIS or LIS의 검사코드))
    'gtypPAT_RES.ORDDT   '/ORDDT(처방일자)
    'gtypPAT_RES.ORDGB   '/ORDGB(처방종류(O.외래, I.입원, G.건강검진)
On Error GoTo RTN_ERR
    
'/-------------------------------------------------------------------------------------------------------------------/
    
    If ConnDB_LOC = True Then
        '/장비코드별 처방코드 가져오기
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
    
'/2.오더정보 조회----------------------------------------------------------------------------------------------------/
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
    
    '/Patient ID 가 바코드일 경우 병록번호를 찾는다.
    gtypPAT_RES.PATNO = "" '/PATNO(병록번호)


    '/공통
    gtypPAT_RES.PATNM = ""  '/PATNM(수검자명)
    gtypPAT_RES.PATSEX = "" '/PATSEX(성별)
    gtypPAT_RES.PATAGE = "" '/PATAGE(연령)
    
    '/적용기관별로 로직을 정할 것--------------------------------------------------/

    If ConnDB_HIS = False Then Exit Function
                           gstrQuy = "SELECT A.*  "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PATIENT A, EXAMRES B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID = B.PID "
        gstrQuy = gstrQuy & vbCrLf & "   AND  B.SPECIMENID = '" & gtypPAT_RES.BARCD & "' "
    If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End

    If Not ADR_HIS Is Nothing Then
        
        Do Until ADR_HIS.EOF
            gtypPAT_RES.PATNO = Trim(ADR_HIS!PID & "")
            gtypPAT_RES.PATNM = Trim(ADR_HIS!PNAME & "")    '/PATNM(수검자명)
            gtypPAT_RES.PATSEX = Trim(ADR_HIS!SEX & "")  '/PATSEX(성별)
            'gtypPAT_RES.PATAGE = Trim(ADR_HIS!PID & "")  '/PATAGE(연령)
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
    
    gtypPAT_RES.AFLAG = "" '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
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
    
    '/결과 소숫점자리 적용
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
     
     
    '/검사코드를 못불러와서 정보를 못받은 경우에는 장비원시데이터를 넣어줌
    If gtypPAT_RES.EXAMCD = "" Then Exit Function
        
    '/Abnomal 판별
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
    
    '/Delta 판별 (보류)
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
    
    '/Panic 판별
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
            gtypPAT_RES.Result = Trim(ADR_LOC!Result & "")      '/결과
            gtypPAT_RES.EQCD = Trim(ADR_LOC!EQCD & "")          '/장비검사코드
            gtypPAT_RES.EXAMCD = Trim(ADR_LOC!EXAMCD & "")            '/처방검사코드
            gtypPAT_RES.DFLAG = Trim(ADR_LOC!DFLAG & "")
            gtypPAT_RES.PFLAG = Trim(ADR_LOC!PFLAG & "")
            gtypPAT_RES.AFLAG = Trim(ADR_LOC!AFLAG & "")

            If FUNC_HIS_SAVE = True Then '/HIS에 결과 전송
                Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS에 결과 전송
                Call SET_CELL(frmEQ_검사결과관리.sprLResult, 8, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "완료", "대기"))
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
    '/선 검사 후 매칭일 경우
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
                Call FUNC_HIS_RESULT_JUDGMENT '/결과 판정
            Else
                gtypPAT_RES.Result = gtypPAT_RES.EQRESULT
            End If
            
            'gtypPAT_RES.Result = GET_CELL(.sprLResult, intCol, argROW)
            
            gtypPAT_RES.EXDT = Format(Now, "YYYYMMDD")
            gtypPAT_RES.EXTM = Format(Now, "HHMMSS")
            gtypPAT_RES.RCDT = Format(Now, "YYYYMMDD")      '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
            gtypPAT_RES.RCTM = Format(Now, "HHMMSS")        '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
            gtypPAT_RES.STATEFLAG = "1"                     '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
            gtypPAT_RES.SENDFLAG = "0"
            '/검사SEQ 찾기
            If strEXSEQ <> "Y" Then
                gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/검체번호(Barcode)별 검사회차
                strEXSEQ = "Y"
            End If
            Call FUNC_HIS_PATIENT
            Call FUNC_HIS_ORDER_VIEW
            If gtypPAT_RES.EXAMCD <> "" Then
                If FUNC_LOC_SAVE_PAT_RES = True Then
                    If .mnuJobModeAuto.Checked = True And gtypPAT_RES.BARCD <> "" Then   '/전송방식이 자동전송이면...
                        If FUNC_HIS_SAVE = True Then '/HIS에 결과 전송
                        Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS에 결과 전송
                        Call SET_CELL(.sprLResult, 7, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "완료", "대기"))
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
    gtypPAT_RES.EXSEQ = ""            '/EXSEQ(검체번호(Barcode)별 검사회차)
    gtypPAT_RES.EQCD = ""             '/EQCD(장비검사코드)
    gtypPAT_RES.EXAMCD = ""           '/EXAMCD(처방코드(HIS or LIS의 검사코드))
    gtypPAT_RES.EXDT = ""             '/EXDT(검사처방전송일자(YYYYMMDD) HIEQ->의료장비)
    gtypPAT_RES.EXTM = ""             '/EXTM(검사처방전송시간(24HHMMSS) HIEQ->의료장비)
    gtypPAT_RES.RCDT = ""             '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
    gtypPAT_RES.RCTM = ""             '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
    gtypPAT_RES.SDDT = ""             '/SDDT(검사결과전송일자(YYYYMMDD) HIEQ->HIS)
    gtypPAT_RES.SDTM = ""             '/SDTM(검사결과전송시간(24HHMMSS) HIEQ->HIS)
    gtypPAT_RES.Result = ""           '/RESULT(검사결과(변형된 결과))
    gtypPAT_RES.EQRESULT = ""         '/EQRESULT(장비원시결과)
    gtypPAT_RES.AFLAG = ""            '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
    gtypPAT_RES.PFLAG = ""            '/PFLAG(Panic)
    gtypPAT_RES.DFLAG = ""            '/DFLAG(Delta)
    gtypPAT_RES.SAMPLENO = ""         '/Sample No(AU2700, Uriscan 등에 사용)
    gtypPAT_RES.DISKNO = ""           '/DISKNO(디스크번호 or 렉번호)
    gtypPAT_RES.POSNO = ""            '/POSNO(위치번호)
    gtypPAT_RES.ORDDT = ""            '/ORDDT(처방일자)
    gtypPAT_RES.ORDGB = ""            '/ORDGB(처방종류(O.외래, I.입원, G.건강검진))
    gtypPAT_RES.PATNO = ""            '/PATNO(병록번호)
    gtypPAT_RES.PATNM = ""            '/PATNM(수검자명)
    gtypPAT_RES.PATSEX = ""           '/PATSEX(성별)
    gtypPAT_RES.PATAGE = ""           '/PATAGE(연령)
    gtypPAT_RES.SENDFLAG = ""         '/SENDFLAG(HIS 전송 FLAG (0:대기, 1:완료))
    gtypPAT_RES.STATEFLAG = ""        '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
    
End With
    NOVA_SAVE = True
End Function
