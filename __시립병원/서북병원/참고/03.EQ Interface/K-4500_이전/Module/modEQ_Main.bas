Attribute VB_Name = "modEQ_Main"
Option Explicit

'/임의 변수
Public intX                 As Integer
Public intY                 As Integer
Public strTemp              As String

Public gstrHOS_CUSCD        As String

Type USER_INFO
    USERID          As String
    USERNM          As String
    USERPW          As String
End Type
Public gtypUSER     As USER_INFO

Public gstrUSERID           As String
Public gstrUSERNM           As String
Public gstrUSERPW           As String

'/PAT_RES Table Data 변수
Type PAT_RES
    BARCD           As String   '/BARCD(검체번호(Barcode))
    EXSEQ           As String   '/EXSEQ(검체번호(Barcode)별 검사회차)
    EQCD            As String   '/EQCD(장비검사코드)
    EXAMCD          As String   '/EXAMCD(처방코드(HIS or LIS의 검사코드))
    EXDT            As String   '/EXDT(검사처방전송일자(YYYYMMDD) HIEQ->의료장비)
    EXTM            As String   '/EXTM(검사처방전송시간(24HHMMSS) HIEQ->의료장비)
    RCDT            As String   '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
    RCTM            As String   '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
    SDDT            As String   '/SDDT(검사결과전송일자(YYYYMMDD) HIEQ->HIS)
    SDTM            As String   '/SDTM(검사결과전송시간(24HHMMSS) HIEQ->HIS)
    Result          As String   '/RESULT(검사결과(변형된 결과))
    EQRESULT        As String   '/EQRESULT(장비원시결과)
    AFLAG           As String   '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
    PFLAG           As String   '/PFLAG(Panic)
    DFLAG           As String   '/DFLAG(Delta)
    SAMPLENO        As String   '/Sample No(AU2700, Uriscan 등에 사용)
    DISKNO          As String   '/DISKNO(디스크번호 or 렉번호)
    POSNO           As String   '/POSNO(위치번호)
    ORDDT           As String   '/ORDDT(처방일자)
    ORDGB           As String   '/ORDGB(처방종류(O.외래, I.입원, G.건강검진))
    PATNO           As String   '/PATNO(병록번호)
    PATNM           As String   '/PATNM(수검자명)
    PATSEX          As String   '/PATSEX(성별)
    PATAGE          As String   '/PATAGE(연령)
    SENDFLAG        As String   '/SENDFLAG(HIS 전송 FLAG (0:대기, 1:완료))
    STATEFLAG       As String   '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
End Type
Public gtypPAT_RES  As PAT_RES

Global Const gintEQ_StartCol    As Integer = 16                 '/실시간 검사리스트(frmEQ_Main) 검사항목 처음 Column
Public gstrMSCOMM_Buff          As String                       '/MSComm Input String
Public gstrEQORDREAD            As String                       '/검사수행 가능한 처방들

'/기본프린터 지정
Private Const HWND_BROADCAST    As Long = &HFFFF&
Private Const WM_WININICHANGE   As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

'/입력/조회 Form 연결 변수
Public gstrInputUpdate     As String  '/1.Input, 2.Update
Public gstrInputUpdateYN   As Boolean '/True.변화있음, False.변화없음
Public gstrArgTemp1        As String  '/PopUp Form 전달 변수1
Public gstrArgTemp2        As String  '/PopUp Form 전달 변수2
Public gstrArgTemp3        As String  '/PopUp Form 전달 변수3
Public gstrArgTemp4        As String  '/PopUp Form 전달 변수3
Public gstrArgTemp5        As String  '/PopUp Form 전달 변수3

Public Function FUNC_GET_EXCD() As Boolean
    FUNC_GET_EXCD = False
    gstrEQORDREAD = ""
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT EXCD "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQORDREADYN = 'Y' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        Do Until ADR_LOC.EOF
            gstrEQORDREAD = gstrEQORDREAD & ",'" & Trim(ADR_LOC!EXCD & "") & "'"
            
            ADR_LOC.MoveNext
        Loop
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        gstrEQORDREAD = Mid(gstrEQORDREAD, 2)
    End If
    
    Call CloseDB_LOC
    
    FUNC_GET_EXCD = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_GET_EXSEQ(argBARCD As String) As String
    FUNC_GET_EXSEQ = "1"

    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT EXSEQ, STATEFLAG "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD     = '" & gtypPAT_RES.BARCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO  = '" & gtypPAT_RES.SAMPLENO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO    = '" & gtypPAT_RES.DISKNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND POSNO     = '" & gtypPAT_RES.POSNO & "' "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY EXSEQ DESC "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        If Trim(ADR_LOC!STATEFLAG & "") = "1" Then '/결과진행상태 (0:처방, 1:결과)
            FUNC_GET_EXSEQ = Val(ADR_LOC!EXSEQ & "") + 1 '/1증가
        Else
            FUNC_GET_EXSEQ = Val(ADR_LOC!EXSEQ & "") '/그대로
        End If

        ADR_LOC.Close: Set ADR_LOC = Nothing
    End If
    
    Call CloseDB_LOC
End Function

Public Function FUNC_LOC_BARCD_DELETE(ArgSection As String) As Boolean
    FUNC_LOC_BARCD_DELETE = False
    
On Error GoTo ERR_RTN
    
    
    FUNC_LOC_BARCD_DELETE = True

Exit Function


'/----------------------------------------------------------------------------------------------------/

ERR_RTN:
    MsgBox "삭제 오류!!!", vbCritical, "확인"
End Function

Public Function FUNC_LOC_BARCD_UPDATE(ArgSection As String) As Boolean
    FUNC_LOC_BARCD_UPDATE = False
    
On Error GoTo ERR_RTN
    
    
    FUNC_LOC_BARCD_UPDATE = True

Exit Function


'/----------------------------------------------------------------------------------------------------/

ERR_RTN:
    MsgBox "삭제 오류!!!", vbCritical, "확인"
End Function

Public Function FUNC_LOC_SAVE_PAT_RES() As Boolean
    '/1.결과신호 받을 때
    '/2.처방내역 발생 때
    '/3.장비 특성에 따라 WHERE 절 조건을 달리한다. 단 BARCD,EXSEQ,EQCD는 유지한다.
    
    FUNC_LOC_SAVE_PAT_RES = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
                       gstrQuy = "SELECT BARCD "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD     = '" & gtypPAT_RES.BARCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ     =  " & Val(gtypPAT_RES.EXSEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQCD      = '" & gtypPAT_RES.EQCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO  = '" & gtypPAT_RES.SAMPLENO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO    = '" & gtypPAT_RES.DISKNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND POSNO     = '" & gtypPAT_RES.POSNO & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
    
                           gstrQuy = "UPDATE PAT_RES SET "
        gstrQuy = gstrQuy & vbCrLf & "       RCDT      = '" & gtypPAT_RES.RCDT & "', "      '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "       RCTM      = '" & gtypPAT_RES.RCTM & "', "      '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "       RESULT    = '" & gtypPAT_RES.Result & "', "    '/RESULT(검사결과(변형된 결과))
        gstrQuy = gstrQuy & vbCrLf & "       EQRESULT  = '" & gtypPAT_RES.EQRESULT & "', "  '/EQRESULT(장비원시결과)
        gstrQuy = gstrQuy & vbCrLf & "       AFLAG     = '" & gtypPAT_RES.AFLAG & "', "     '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
        gstrQuy = gstrQuy & vbCrLf & "       PFLAG     = '" & gtypPAT_RES.PFLAG & "', "     '/PFLAG(Panic)
        gstrQuy = gstrQuy & vbCrLf & "       DFLAG     = '" & gtypPAT_RES.DFLAG & "', "     '/DFLAG(Delta)
        gstrQuy = gstrQuy & vbCrLf & "       SENDFLAG  = '0', "                             '/SENDFLAG(HIS 전송 FLAG (0:대기, 1:완료))
        gstrQuy = gstrQuy & vbCrLf & "       STATEFLAG = '1' "                              '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
        gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD     = '" & gtypPAT_RES.BARCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ     =  " & Val(gtypPAT_RES.EXSEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQCD      = '" & gtypPAT_RES.EQCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO  = '" & gtypPAT_RES.SAMPLENO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO    = '" & gtypPAT_RES.DISKNO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND POSNO     = '" & gtypPAT_RES.POSNO & "' "
    Else
                           gstrQuy = "INSERT INTO PAT_RES "
        gstrQuy = gstrQuy & vbCrLf & " (BARCD,      EXSEQ,      EQCD,   EXAMCD, EXDT, "
        gstrQuy = gstrQuy & vbCrLf & "  EXTM,       RCDT,       RCTM,   SDDT,   SDTM, "
        gstrQuy = gstrQuy & vbCrLf & "  RESULT,     EQRESULT,   AFLAG,  PFLAG,  DFLAG, "
        gstrQuy = gstrQuy & vbCrLf & "  SAMPLENO,   DISKNO,     POSNO,  ORDDT,  ORDGB, "
        gstrQuy = gstrQuy & vbCrLf & "  PATNO,      PATNM,      PATSEX, PATAGE, SENDFLAG, "
        gstrQuy = gstrQuy & vbCrLf & "  STATEFLAG) "
        gstrQuy = gstrQuy & vbCrLf & " VALUES "
        gstrQuy = gstrQuy & vbCrLf & " ('" & gtypPAT_RES.BARCD & "', "      '/BARCD(검체번호(Barcode))
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(gtypPAT_RES.EXSEQ) & ", "  '/EXSEQ(검체번호(Barcode)별 검사회차)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EQCD & "', "       '/EQCD(장비검사코드)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EXAMCD & "', "     '/EXAMCD(처방코드(HIS or LIS의 검사코드))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EXDT & "', "       '/EXDT(검사처방전송일자(YYYYMMDD) HIEQ->의료장비)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EXTM & "', "       '/EXTM(검사처방전송시간(24HHMMSS) HIEQ->의료장비)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.RCDT & "', "       '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.RCTM & "', "       '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SDDT & "', "       '/SDDT(검사결과전송일자(YYYYMMDD) HIEQ->HIS)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SDTM & "', "       '/SDTM(검사결과전송시간(24HHMMSS) HIEQ->HIS)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.Result & "', "     '/RESULT(검사결과(변형된 결과))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EQRESULT & "', "   '/EQRESULT(장비원시결과)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.AFLAG & "', "      '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PFLAG & "', "      '/PFLAG(Panic)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.DFLAG & "', "      '/DFLAG(Delta)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SAMPLENO & "', "   '/Sample No(AU2700, Uriscan 등에 사용)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.DISKNO & "', "     '/DISKNO(디스크번호 or 렉번호)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.POSNO & "', "      '/POSNO(위치번호)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.ORDDT & "', "      '/ORDDT(처방일자)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.ORDGB & "', "      '/ORDGB(처방종류(O.외래, I.입원, G.건강검진))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATNO & "', "      '/PATNO(병록번호)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATNM & "', "      '/PATNM(수검자명)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATSEX & "', "     '/PATSEX(성별)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATAGE & "', "     '/PATAGE(연령)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SENDFLAG & "', "   '/SENDFLAG(HIS 전송 FLAG (0:대기, 1:완료))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.STATEFLAG & "') "  '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
    End If
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
'''    gstrQuy = "DELETE FROM PAT_RES "
'''    gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD    = '" & gtypPAT_RES.BARCD & "' "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ    =  " & Val(gtypPAT_RES.EXSEQ) & " "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND EQCD     = '" & gtypPAT_RES.EQCD & "' "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO = '" & gtypPAT_RES.SAMPLENO & "' "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO   = '" & gtypPAT_RES.DISKNO & "' "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND POSNO    = '" & gtypPAT_RES.POSNO & "' "
'''    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function

    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_LOC_SAVE_PAT_RES = True
    
Exit Function
    
'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_LOC_SAVE_SEND(argBARCD As String, argEXSEQ As String, argEQCD As String, argSAMPLENO As String, argDISKNO As String, argPOSNO As String, argSENDFLAG As String) As Boolean
    FUNC_LOC_SAVE_SEND = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
    gstrQuy = "UPDATE PAT_RES SET "
    gstrQuy = gstrQuy & vbCrLf & "       SDDT      = '" & Format(Now, "YYYYMMDD") & "', "   '/SDDT(검사결과전송일자(YYYYMMDD) HIEQ->HIS)
    gstrQuy = gstrQuy & vbCrLf & "       SDTM      = '" & Format(Now, "HHMMSS") & "', "     '/SDTM(검사결과전송시간(24HHMMSS) HIEQ->HIS)
    gstrQuy = gstrQuy & vbCrLf & "       SENDFLAG  = '" & argSENDFLAG & "' "                '/SENDFLAG(HIS 전송 FLAG (0:대기, 1:완료))
    gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD     = '" & argBARCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ     =  " & Val(gtypPAT_RES.EXSEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQCD      = '" & argEQCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO  = '" & argSAMPLENO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO    = '" & argDISKNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND POSNO     = '" & argPOSNO & "' "
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_LOC_SAVE_SEND = True
Exit Function
    
'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_RESULT_CHANGE(argECCD As String, argEQRESULT As String)
    FUNC_RESULT_CHANGE = gtypPAT_RES.EQRESULT




End Function

Public Sub Main()
    frmEQ_Main.Show
End Sub
