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
    
    '/���� �׸�
    'gtypPAT_RES.EXAMCD  '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
    'gtypPAT_RES.ORDDT   '/ORDDT(ó������)
    'gtypPAT_RES.ORDGB   '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����)
On Error GoTo RTN_ERR
    
    If ConnDB_LOC = True Then
        '/����ڵ庰 ó���ڵ� ��������
        gstrQuy = "SELECT EXCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & gtypPAT_RES.EQCD & "' "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End

        If Not ADR_LOC Is Nothing Then
            gtypPAT_RES.EXAMCD = Trim(ADR_LOC!EXCD & "") '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
            
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        Call CloseDB_LOC
    End If
    
    gtypPAT_RES.ORDDT = ""  '/ORDDT(ó������)
    gtypPAT_RES.ORDGB = ""  '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))

    FUNC_HIS_ORDER_VIEW = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

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

    FUNC_HIS_PATIENT = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    
End Function

Public Function FUNC_HIS_RESULT_JUDGMENT()
    gtypPAT_RES.AFLAG = "" '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
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
