Attribute VB_Name = "modSRL"
Option Explicit

'--- 2004-03-17 KHS ADD (SRLDEVC.dll�� �ܺ��Լ��� �ҷ��� CALL�Ѵ�)
Declare Function AS400DOWNF$ Lib "C:\Windows\System\SRLDEVC.dll" (ByVal DEVC$, ByVal FROMDATE$, ByVal TODATE$)
Declare Function AS400UPF$ Lib "C:\Windows\System\SRLDEVC.dll" (ByVal DEVC$, ByVal NOWDATE$)
'test
Declare Function AS400DOWNSF$ Lib "C:\Windows\System\SRLDEVC2.dll" (ByVal DEVC$, ByVal JDATE$, ByVal JNO$)


Public Sub EditBarCode_SRL(ByVal sBarCd As String, ByVal sCurDate As String, _
                        ByRef sJDate As String, ByRef sJNo As String)
    On Error GoTo ErrEdit
    
    Dim tmpDate As String
    
    sJDate = "": sJNo = ""
    
    If Len(sBarCd) < 11 Then
        Exit Sub
    End If
    
    tmpDate = Left(sCurDate, 3) & Mid(sBarCd, 1, 1) & "-01-01"
    
    sJDate = Format(DateAdd("d", Val(Mid(sBarCd, 2, 3)) - 1, tmpDate), "YYYYMMDD")
    sJNo = Mid(sBarCd, 5, 5)
    
ErrEdit:
    If Err <> 0 Then
        ViewMsg "EditBarCode_SRL Err - " & Err.Description
    End If
End Sub


Public Function fOrdDownAS400_Realtime(ByVal sDevCd As String, ByVal sJDate As String, ByVal sJNo As String) As Boolean
    On Error GoTo ErrAS400
    
    Dim sRet    As String
    
    fOrdDownAS400_Realtime = False
   
    '--- 2004-03-17 KHS ADD (SRLDEVC2.dll�� �ܺ��Լ��� �ҷ��� CALL�Ѵ�)
    sRet = AS400DOWNSF(sDevCd, sJDate, sJNo)
        
    Select Case Trim(sRet)
        Case "E1"   '����ڵ� Error
            ViewMsg "E1: ����ڵ� Error"
        Case "E2"   'FromDate��¥ Error
            ViewMsg "E2: ���ڵ��ȣ Error"
        Case "O"    '����
            fOrdDownAS400_Realtime = True
        Case "N"    '������ ����
            ViewMsg "N: ������ ����"
        Case "E"    '�Լ�����Error
            ViewMsg "E: �Լ�����Error"
    End Select

ErrAS400:
    If Err <> 0 Then
        ViewMsg "fOrdDownAS400_RT Err - " & Err.Description
    End If
End Function


Public Function fOrdDownAS400_Batch(ByVal sDevCd As String, ByVal sDate1 As String, ByVal sDate2 As String) As Boolean
    On Error GoTo ErrAS400
    
    Dim sRet    As String
    
    fOrdDownAS400_Batch = False
    
    '--- 2004-03-17 KHS ADD (SRLDEVC.dll�� �ܺ��Լ��� �ҷ��� CALL�Ѵ�)
    sRet = AS400DOWNF(sDevCd, sDate1, sDate2)
        
    Select Case Trim(sRet)
        Case "E1"   '����ڵ� Error
            ViewMsg "E1: ����ڵ� Error"
        Case "E2"   'FromDate��¥ Error
            ViewMsg "E2: FromDate��¥ Error"
        Case "E3"   'ToDate��¥ Error
            ViewMsg "E3: ToDate��¥ Error"
        Case "O"    '����
            fOrdDownAS400_Batch = True
        Case "N"    '������ ����
            ViewMsg "N: ������ ����"
        Case "E"    '�Լ�����Error
            ViewMsg "E: �Լ�����Error"
    End Select

ErrAS400:
    If Err <> 0 Then
        ViewMsg "fOrdDownAS400_Batch Err - " & Err.Description
    End If
End Function
