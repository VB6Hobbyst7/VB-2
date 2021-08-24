Attribute VB_Name = "modEQ_Main"
Option Explicit

'/���� ����
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

'/PAT_RES Table Data ����
Type PAT_RES
    BARCD           As String   '/BARCD(��ü��ȣ(Barcode))
    EXSEQ           As String   '/EXSEQ(��ü��ȣ(Barcode)�� �˻�ȸ��)
    EQCD            As String   '/EQCD(���˻��ڵ�)
    EXAMCD          As String   '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
    EXDT            As String   '/EXDT(�˻�ó����������(YYYYMMDD) HIEQ->�Ƿ����)
    EXTM            As String   '/EXTM(�˻�ó�����۽ð�(24HHMMSS) HIEQ->�Ƿ����)
    RCDT            As String   '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
    RCTM            As String   '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
    SDDT            As String   '/SDDT(�˻�����������(YYYYMMDD) HIEQ->HIS)
    SDTM            As String   '/SDTM(�˻������۽ð�(24HHMMSS) HIEQ->HIS)
    Result          As String   '/RESULT(�˻���(������ ���))
    EQRESULT        As String   '/EQRESULT(�����ð��)
    AFLAG           As String   '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
    PFLAG           As String   '/PFLAG(Panic)
    DFLAG           As String   '/DFLAG(Delta)
    SAMPLENO        As String   '/Sample No(AU2700, Uriscan � ���)
    DISKNO          As String   '/DISKNO(��ũ��ȣ or ����ȣ)
    POSNO           As String   '/POSNO(��ġ��ȣ)
    ORDDT           As String   '/ORDDT(ó������)
    ORDGB           As String   '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))
    PATNO           As String   '/PATNO(���Ϲ�ȣ)
    PATNM           As String   '/PATNM(�����ڸ�)
    PATSEX          As String   '/PATSEX(����)
    PATAGE          As String   '/PATAGE(����)
    SENDFLAG        As String   '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
    STATEFLAG       As String   '/STATEFLAG(���������� (0:ó��, 1:���))
End Type
Public gtypPAT_RES  As PAT_RES

Global Const gintEQ_StartCol    As Integer = 16                 '/�ǽð� �˻縮��Ʈ(frmEQ_Main) �˻��׸� ó�� Column
Public gstrMSCOMM_Buff          As String                       '/MSComm Input String
Public gstrEQORDREAD            As String                       '/�˻���� ������ ó���

'/�⺻������ ����
Private Const HWND_BROADCAST    As Long = &HFFFF&
Private Const WM_WININICHANGE   As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

'/�Է�/��ȸ Form ���� ����
Public gstrInputUpdate     As String  '/1.Input, 2.Update
Public gstrInputUpdateYN   As Boolean '/True.��ȭ����, False.��ȭ����
Public gstrArgTemp1        As String  '/PopUp Form ���� ����1
Public gstrArgTemp2        As String  '/PopUp Form ���� ����2
Public gstrArgTemp3        As String  '/PopUp Form ���� ����3
Public gstrArgTemp4        As String  '/PopUp Form ���� ����3
Public gstrArgTemp5        As String  '/PopUp Form ���� ����3

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
        If Trim(ADR_LOC!STATEFLAG & "") = "1" Then '/���������� (0:ó��, 1:���)
            FUNC_GET_EXSEQ = Val(ADR_LOC!EXSEQ & "") + 1 '/1����
        Else
            FUNC_GET_EXSEQ = Val(ADR_LOC!EXSEQ & "") '/�״��
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
    MsgBox "���� ����!!!", vbCritical, "Ȯ��"
End Function

Public Function FUNC_LOC_BARCD_UPDATE(ArgSection As String) As Boolean
    FUNC_LOC_BARCD_UPDATE = False
    
On Error GoTo ERR_RTN
    
    
    FUNC_LOC_BARCD_UPDATE = True

Exit Function


'/----------------------------------------------------------------------------------------------------/

ERR_RTN:
    MsgBox "���� ����!!!", vbCritical, "Ȯ��"
End Function

Public Function FUNC_LOC_SAVE_PAT_RES() As Boolean
    '/1.�����ȣ ���� ��
    '/2.ó�泻�� �߻� ��
    '/3.��� Ư���� ���� WHERE �� ������ �޸��Ѵ�. �� BARCD,EXSEQ,EQCD�� �����Ѵ�.
    
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
        gstrQuy = gstrQuy & vbCrLf & "       RCDT      = '" & gtypPAT_RES.RCDT & "', "      '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "       RCTM      = '" & gtypPAT_RES.RCTM & "', "      '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "       RESULT    = '" & gtypPAT_RES.Result & "', "    '/RESULT(�˻���(������ ���))
        gstrQuy = gstrQuy & vbCrLf & "       EQRESULT  = '" & gtypPAT_RES.EQRESULT & "', "  '/EQRESULT(�����ð��)
        gstrQuy = gstrQuy & vbCrLf & "       AFLAG     = '" & gtypPAT_RES.AFLAG & "', "     '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
        gstrQuy = gstrQuy & vbCrLf & "       PFLAG     = '" & gtypPAT_RES.PFLAG & "', "     '/PFLAG(Panic)
        gstrQuy = gstrQuy & vbCrLf & "       DFLAG     = '" & gtypPAT_RES.DFLAG & "', "     '/DFLAG(Delta)
        gstrQuy = gstrQuy & vbCrLf & "       SENDFLAG  = '0', "                             '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
        gstrQuy = gstrQuy & vbCrLf & "       STATEFLAG = '1' "                              '/STATEFLAG(���������� (0:ó��, 1:���))
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
        gstrQuy = gstrQuy & vbCrLf & " ('" & gtypPAT_RES.BARCD & "', "      '/BARCD(��ü��ȣ(Barcode))
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(gtypPAT_RES.EXSEQ) & ", "  '/EXSEQ(��ü��ȣ(Barcode)�� �˻�ȸ��)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EQCD & "', "       '/EQCD(���˻��ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EXAMCD & "', "     '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EXDT & "', "       '/EXDT(�˻�ó����������(YYYYMMDD) HIEQ->�Ƿ����)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EXTM & "', "       '/EXTM(�˻�ó�����۽ð�(24HHMMSS) HIEQ->�Ƿ����)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.RCDT & "', "       '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.RCTM & "', "       '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SDDT & "', "       '/SDDT(�˻�����������(YYYYMMDD) HIEQ->HIS)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SDTM & "', "       '/SDTM(�˻������۽ð�(24HHMMSS) HIEQ->HIS)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.Result & "', "     '/RESULT(�˻���(������ ���))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.EQRESULT & "', "   '/EQRESULT(�����ð��)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.AFLAG & "', "      '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PFLAG & "', "      '/PFLAG(Panic)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.DFLAG & "', "      '/DFLAG(Delta)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SAMPLENO & "', "   '/Sample No(AU2700, Uriscan � ���)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.DISKNO & "', "     '/DISKNO(��ũ��ȣ or ����ȣ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.POSNO & "', "      '/POSNO(��ġ��ȣ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.ORDDT & "', "      '/ORDDT(ó������)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.ORDGB & "', "      '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATNO & "', "      '/PATNO(���Ϲ�ȣ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATNM & "', "      '/PATNM(�����ڸ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATSEX & "', "     '/PATSEX(����)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.PATAGE & "', "     '/PATAGE(����)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.SENDFLAG & "', "   '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypPAT_RES.STATEFLAG & "') "  '/STATEFLAG(���������� (0:ó��, 1:���))
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
    gstrQuy = gstrQuy & vbCrLf & "       SDDT      = '" & Format(Now, "YYYYMMDD") & "', "   '/SDDT(�˻�����������(YYYYMMDD) HIEQ->HIS)
    gstrQuy = gstrQuy & vbCrLf & "       SDTM      = '" & Format(Now, "HHMMSS") & "', "     '/SDTM(�˻������۽ð�(24HHMMSS) HIEQ->HIS)
    gstrQuy = gstrQuy & vbCrLf & "       SENDFLAG  = '" & argSENDFLAG & "' "                '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
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
