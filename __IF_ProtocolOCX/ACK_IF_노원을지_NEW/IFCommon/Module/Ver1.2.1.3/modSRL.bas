Attribute VB_Name = "modSRL"
Option Explicit

'--- 2004-03-17 KHS ADD (SRLDEVC.dll을 외부함수로 불러서 CALL한다)
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
   
    '--- 2004-03-17 KHS ADD (SRLDEVC2.dll을 외부함수로 불러서 CALL한다)
    sRet = AS400DOWNSF(sDevCd, sJDate, sJNo)
        
    Select Case Trim(sRet)
        Case "E1"   '장비코드 Error
            ViewMsg "E1: 장비코드 Error"
        Case "E2"   'FromDate날짜 Error
            ViewMsg "E2: 바코드번호 Error"
        Case "O"    '성공
            fOrdDownAS400_Realtime = True
        Case "N"    '데이터 없음
            ViewMsg "N: 데이터 없음"
        Case "E"    '함수실행Error
            ViewMsg "E: 함수실행Error"
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
    
    '--- 2004-03-17 KHS ADD (SRLDEVC.dll을 외부함수로 불러서 CALL한다)
    sRet = AS400DOWNF(sDevCd, sDate1, sDate2)
        
    Select Case Trim(sRet)
        Case "E1"   '장비코드 Error
            ViewMsg "E1: 장비코드 Error"
        Case "E2"   'FromDate날짜 Error
            ViewMsg "E2: FromDate날짜 Error"
        Case "E3"   'ToDate날짜 Error
            ViewMsg "E3: ToDate날짜 Error"
        Case "O"    '성공
            fOrdDownAS400_Batch = True
        Case "N"    '데이터 없음
            ViewMsg "N: 데이터 없음"
        Case "E"    '함수실행Error
            ViewMsg "E: 함수실행Error"
    End Select

ErrAS400:
    If Err <> 0 Then
        ViewMsg "fOrdDownAS400_Batch Err - " & Err.Description
    End If
End Function
