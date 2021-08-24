Attribute VB_Name = "modHUBIQUANPRO"
Option Explicit

Public Function FUNC_HIS_ORDER_MAKE() As String
    FUNC_HIS_ORDER_VIEW = False
    
On Error GoTo RTN_ERR
    
    FUNC_HIS_ORDER_VIEW = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    
End Function

Public Function FUNC_HIS_ORDER_VIEW() As Boolean
    FUNC_HIS_ORDER_VIEW = False
    
    '/정의 항목
    'gtypPAT_RES.EXAMCD  '/EXAMCD(처방코드(HIS or LIS의 검사코드))
    'gtypPAT_RES.ORDDT   '/ORDDT(처방일자)
    'gtypPAT_RES.ORDGB   '/ORDGB(처방종류(O.외래, I.입원, G.건강검진)
On Error GoTo RTN_ERR
    
    If ConnDB_LOC = True Then
        '/장비코드별 처방코드 가져오기
        gstrQuy = "SELECT EXCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & gtypPAT_RES.EQCD & "' "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End

        If Not ADR_LOC Is Nothing Then
            gtypPAT_RES.EXAMCD = Trim(ADR_LOC!EXCD & "") '/EXAMCD(처방코드(HIS or LIS의 검사코드))
            
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        Call CloseDB_LOC
    End If
    
    gtypPAT_RES.ORDDT = ""  '/ORDDT(처방일자)
    gtypPAT_RES.ORDGB = ""  '/ORDGB(처방종류(O.외래, I.입원, G.건강검진))

    FUNC_HIS_ORDER_VIEW = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

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

    FUNC_HIS_PATIENT = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    
End Function

Public Function FUNC_HIS_RESULT_JUDGMENT()
    gtypPAT_RES.AFLAG = "" '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
    gtypPAT_RES.PFLAG = "" '/PFLAG(Panic)
    gtypPAT_RES.DFLAG = "" '/DFLAG(Delta)
End Function

Public Function FUNC_HIS_SAVE() As Boolean
    FUNC_HIS_SAVE = False
    
On Error GoTo RTN_ERR

    FUNC_HIS_SAVE = True
    
Exit Function
    
'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function
