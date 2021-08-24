VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCO0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type tPrt진료과정보
    sDeptNm     As String
End Type

Private Type tPrt루틴코드정보
    sRtNm      As String
End Type

Private Type tPrt출력Font
    sFontSize   As String
End Type

Private Type tPrt환자정보
    sField      As String   ' SQL 사용시 필드명
    sMaxLen     As String   ' 각 ITEM의 문자열 최대길이
    sAlign      As String   ' 좌우 정렬값
    sXPos       As String
    sYPos       As String
    sRealStr    As String   ' 실제로 출력시 사용할 내용(당연히 환자마다 내용이 다르겠지요)
End Type

'-- 2001.08.09(osw)추가 : 검사명출력을 위한..
Private Type tPrt검사정보
    sMaxLen     As String   ' 각 ITEM의 문자열 최대길이
    sAlign      As String   ' 좌우 정렬값
    sXPos       As String
    sYPos       As String
    sRealStr    As String   ' 결과
    sFlagStr    As String   ' 플레그
End Type

'-- 2001.08.13(osw)추가 : 참고치출력을 위한..
Private Type tPrt참고정보
    sMaxLen     As String   ' 각 ITEM의 문자열 최대길이
    sAlign      As String   ' 좌우 정렬값
    sXPos       As String
    sYPos       As String
    sRealStr    As String   ' 결과
    sFlagStr    As String   ' 플레그
End Type

Private Type tPrt결과정보
    sMaxLen     As String   ' 각 ITEM의 문자열 최대길이
    sAlign      As String   ' 좌우 정렬값
    sXPos       As String
    sYPos       As String
    sRealStr    As String   ' 결과
    sFlagStr    As String   ' 플레그
End Type

' 결과지
Dim iUserNmLen          As Integer      ' 보고자 이름의 길이
Dim iUserNmXPos         As Integer
Dim iUserNmYPos         As Integer
Dim sUserNm             As String
Dim sUserNmAlign        As String      ' 좌우정렬

' 결과대장
Dim iCntPrt검체정보     As Integer      ' 검체코드정보의 갯수
Dim iCntPrt환자정보     As Integer      ' 환자정보의 갯수
Dim iCntPrt결과정보     As Integer      ' 검사결과의 갯수
Dim iCntPrt진료과정보   As Integer

'-- 2001.08.09(osw)
Dim iCntPrt검사정보     As Integer      ' 환자정보의 갯수

'-- 2001.08.13(osw)
Dim iCntPrt참고정보     As Integer      ' 참고정보의 갯수


'Dim tPrtDept(99)        As tPrt진료과정보 '배열화 된 tPrt진료과정보
'KTY 수정 19990305
Dim tPrtDept(2340)        As tPrt진료과정보 '배열화 된 tPrt진료과정보
Dim tPrtRtcd(999)       As tPrt루틴코드정보 ' 배열화 된 tPrt루틴코드정보
Dim tPrtFontSize(10)    As tPrt출력Font ' 배열화 된 tPrt출력Font
Dim tPrtPatient(30)     As tPrt환자정보 ' 배열화 된 tPrt환자정보
Dim tPrtResult(100)     As tPrt결과정보 ' 배열화 된 tPrt결과정보

'-- 2001.08.09(osw)
Dim tPrtItem(100)       As tPrt검사정보 ' 배열화 된 tPrt검사정보

'-- 2001.08.13(osw)
Dim tPrtRef(100)       As tPrt참고정보 ' 배열화 된 tPrt검사정보


Dim iMaxLenComment      As Integer      ' Comment 출력시 문자열 최대길이
Dim iMaxLineComment     As Integer      ' Comment 출력시 최대라인
Dim sAlignComment       As String       ' Comment 출력시 정렬방식
Dim sXPos               As String       ' Comment 출력시 X위치
Dim sYPos               As String       ' Comment 출력시 Y위치
Dim sComment            As String       ' Comment 출력시 내용

' 작업대장
Dim iPageLineMax        As Integer      ' 출력시 한 페이지의 최대 라인
Dim iNoCnt              As Integer      ' 출력시 Seq번호를 처리하는 변수
Dim iNowRow             As Integer      ' 현재 출력 Line수
Dim iColMax             As Integer      ' 검사항목 인쇄시 최대 갯수
Dim iColNow             As Integer      ' 검사항목 현재 인쇄 Col위치
Dim sSlipNm             As String       ' Slip명
Dim sDate               As String


' 결과지 출력 관련

' U1) 결과를 출력한 후 출력을 했다는 표시를 하는 처리

Function Edit_PrintFlag(sLabDate As String, sSlipCd As String, sLabSeq As String) As String

' 결과출력 FLAG update
    Dim AdoConn     As New ADODB.Connection
    Dim sSql        As String
        
    On Error GoTo ErrHandler
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = " UPDATE H_JUBSU " & _
           "    SET SLIPYN = '1' " & _
           "  WHERE LABDATE = '" & sLabDate & "'" & _
           "    AND PARTGBN = '" & Mid(sSlipCd, 2, 2) & "'" & _
           "    AND LABSEQ  = '" & sLabSeq & "'"

    AdoConn.Execute sSql
    AdoConn.Close
    
    Edit_PrintFlag = "SUCCESS"
Exit Function
ErrHandler:
    Edit_PrintFlag = Str(Err.Number)
    MsgBox "EDIT (DLL) - " & Err.Description
    
End Function

' U2) 모든 결과가 등록되어 있으면, 자동으로 결과완료정보를 접수TABLE에 기록함
Function Edit_ResultEnd(sLabDate As String, sSlipCd As String, sLabSeq As String) As String

' 결과출력 FLAG update
    Dim AdoConn     As New ADODB.Connection
    Dim DCJ0101     As DCJ0101
    Dim sNowDate    As String
    Dim sNowTime    As String
    Dim sSql        As String
    Dim sPartInit   As String
    Dim sPartGbn    As String
    
    On Error GoTo ErrHandler
    
    'SYSTEM DATE/TIME취득
    Set DCJ0101 = New DCJ0101
        sNowDate = DCJ0101.Get_Date("AS")
    Set DCJ0101 = Nothing
    
    sNowTime = Right(Trim(sNowDate), 6)
    sNowDate = Left(Trim(sNowDate), 8)
    
    sPartInit = Left$(sSlipCd, 1)
    sPartGbn = Mid$(sSlipCd, 2, 2)
    
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    sSql = " UPDATE H_JUBSU AS A " & _
           "    SET RESULTENDDATE = '" & sNowDate & "', " & _
           "        RESULTENDTIME = '" & sNowTime & "', " & _
           "        RESULTENDYN = '1' " & _
           "  WHERE ( SELECT COUNT(C.LABSEQ) FROM H_RESULT AS C " & _
           "           WHERE C.LABDATE = '" & sLabDate & "' " & _
           "             AND C.PARTGBN = '" & sPartGbn & "' " & _
           "             AND C.LABSEQ  = '" & sLabSeq & "' " & _
           "             AND (C.RESULT = '' OR C.RESULT = NULL) ) = 0 " & _
           "    AND A.LABDATE = '" & sLabDate & "' " & _
           "    AND A.PARTGBN= '" & sPartGbn & "' " & _
           "    AND A.LABSEQ = '" & sLabSeq & "' "

    AdoConn.Execute sSql
    AdoConn.Close
    
    Edit_ResultEnd = "SUCCESS"
Exit Function
ErrHandler:
    Edit_ResultEnd = Str(Err.Number)
    MsgBox "EDIT (DLL) - " & Err.Description

End Function

' S1) 출력구간안에 있는 검체번호를 취득하는 루틴
Function Get_LabResult_Info(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String) As String

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim sInfo           As String
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    Dim sSql            As String
    Dim sPrtLabSeq      As String
    
    sInfo = ""
    
    On Error GoTo ErrHandler
    
    If Trim(sSlipCd) = "All" Then
        sSql = "SELECT Distinct LABSEQ " & _
               " FROM H_JUBSU " & _
               "WHERE LABDATE = '" & sLabDate & "'"
    Else
        sSql = "SELECT LABSEQ " & _
               " FROM H_JUBSU " & _
               "WHERE LABDATE = '" & sLabDate & "'"
    End If
    
    If Trim(sSlipCd) = "1" Then
        sSql = sSql & "  AND PARTGBN in ('01','02') "

    ElseIf Trim(sSlipCd) = "2" Then
        sSql = sSql & "  AND PARTGBN = '04' "
    ElseIf Trim(sSlipCd) = "All" Then
        sSql = sSql & "  AND PARTGBN in ('01','02','04') "
    
    Else
        sSql = sSql & "  AND PARTGBN = '" & Mid(sSlipCd, 2, 2) & "'"
    End If
    
    sSql = sSql & "  AND SLIPYN  = '" & sPrtGbn & "'" & _
                  "  AND LABSEQ BETWEEN '" & sLabSeqS & "' AND '" & sLabSeqE & "'" & _
                  "  Order By LABSEQ "
                  
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 검체갯수 | {검체번호}n
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        sPrtLabSeq = ""
        
        For iCnt = 1 To RCnt
            sPrtLabSeq = sPrtLabSeq & AdoRecord("LABSEQ") & "|"
            AdoRecord.MoveNext
        Next iCnt
        sInfo = Str(RCnt) & "|" & sPrtLabSeq
    End If
    AdoRecord.Close
    
    Get_LabResult_Info = sInfo
   
Exit Function
ErrHandler:
    MsgBox "GetItem(출력 DLL) - " & Err.Description
    AdoRecord.Close
    
End Function

' S2) 출력위치와 크기를 미리 정의해 놓은 곳에서 읽어옴
Function Get_Print_Info(PrtKind As String, sSlipCd As String) As Integer

    Dim AdoRecord       As New ADODB.Recordset
    Dim sTmp            As String
    Dim iRtnCd          As Integer
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    Dim sSql            As String
    Dim sPrtLabSeq      As String
    
    iRtnCd = 0
    
    On Error GoTo ErrHandler
    
    ' 진료과 정보를 MEMORY에...
    sSql = " SELECT DEPTCD, DEPTNM " & _
           "   FROM DEPT " & _
           "  WHERE DEPTCD > ''"
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        iCntPrt진료과정보 = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        For iCnt = 1 To iCntPrt진료과정보
            sTmp = AdoRecord("DEPTCD") & ""
            tPrtDept(Asc(UCase$(Left$(sTmp, 1))) * 26 + Asc(UCase$(Right$(sTmp, 1)))).sDeptNm = AdoRecord("DEPTNM") & ""
            AdoRecord.MoveNext
        Next iCnt
    End If
    AdoRecord.Close
    
    Select Case PrtKind
    Case "00"   ' 결과지
        ' 출력정보 MASTER
        sSql = "SELECT MASTERGBN, REALGBN, MAXLEN, ALIGN, PRXPOS, PRYPOS " & _
                " FROM PRINTINFO " & _
                "WHERE PRINTGBN = '" & PrtKind & "' " & _
                "  AND MASTERGBN IN ('0','1','3','4')" & _
                "  AND MAXLEN <> '0'"
            
        AdoRecord.CursorType = adOpenStatic
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
        If AdoRecord.RecordCount <> 0 Then
        ' 출력 자료 구조로 관리
            RCnt = AdoRecord.RecordCount
            AdoRecord.MoveFirst
        
            sPrtLabSeq = ""
        
            iCntPrt환자정보 = 0
            For iCnt = 1 To RCnt
                If AdoRecord("MASTERGBN") & "" = "0" Then
                    tPrtFontSize(Val(AdoRecord("REALGBN") & "")).sFontSize = AdoRecord("MAXLEN") & ""
                ElseIf AdoRecord("MASTERGBN") & "" = "1" Then
                    iCntPrt환자정보 = iCntPrt환자정보 + 1
                    tPrtPatient(iCntPrt환자정보).sField = Mid(AdoRecord("REALGBN") & "", 3)
                    tPrtPatient(iCntPrt환자정보).sMaxLen = AdoRecord("MAXLEN") & ""
                    tPrtPatient(iCntPrt환자정보).sAlign = AdoRecord("ALIGN") & ""
                    tPrtPatient(iCntPrt환자정보).sXPos = AdoRecord("PRXPOS") & ""
                    tPrtPatient(iCntPrt환자정보).sYPos = AdoRecord("PRYPOS") & ""
                ElseIf AdoRecord("MASTERGBN") & "" = "3" Then
                    iMaxLineComment = Val(AdoRecord("REALGBN") & "")
                    iMaxLenComment = Val(AdoRecord("MAXLEN") & "")
                    sAlignComment = AdoRecord("ALIGN") & ""
                    sXPos = AdoRecord("PRXPOS")
                    sYPos = AdoRecord("PRYPOS")
                ElseIf AdoRecord("MASTERGBN") & "" = "4" Then
                    sUserNm = fCurUserNm
                    iUserNmLen = Val(AdoRecord("MAXLEN") & "")
                    sUserNmAlign = AdoRecord("ALIGN") & ""
                    iUserNmXPos = Val(AdoRecord("PRXPOS"))
                    iUserNmYPos = Val(AdoRecord("PRYPOS"))
                    'iUserNmXPos = 210
                    'iUserNmYPos = 297
                End If
                AdoRecord.MoveNext
            Next iCnt
        End If
        AdoRecord.Close
    Case "01"   ' Slip별 작업대장
    ' 진료과 정보를 MEMORY에...
        sSql = " SELECT DISTINCT ROUTINECD, ROUTINENM " & _
               "   FROM ROUTINE " & _
               "  WHERE PARTCD = '" & Left(sSlipCd, 1) & "' " & _
               "    AND PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' "
               
        AdoRecord.CursorType = adOpenStatic
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
        If AdoRecord.RecordCount <> 0 Then
            RCnt = AdoRecord.RecordCount
            AdoRecord.MoveFirst
        
            For iCnt = 1 To RCnt
                tPrtRtcd(Val(AdoRecord("ROUTINECD") & "")).sRtNm = AdoRecord("ROUTINENM") & ""
                AdoRecord.MoveNext
            Next iCnt
        End If
        AdoRecord.Close
    
    Case "02"   ' Slip별 결과대장
        RCnt = 1
    
    End Select
    iRtnCd = RCnt
    Get_Print_Info = iRtnCd
    
Exit Function
ErrHandler:
    MsgBox "GetItem(출력 DLL) - " & Err.Description
    'AdoRecord.Close

End Function

' P1) 각각 환자당 검사결과 출력 루틴 ("00" 결과지)
Function Print_LabResult(sLabDate As String, sSlipCd As String, sLabSeq As String) As Integer

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    Dim sSelect_Item    As String
    Dim sFromStr        As String
    Dim sFromStr1       As String
    Dim sFromStr2       As String
    Dim sFromStr3       As String
    Dim sWhereStr       As String
    Dim sSql            As String
    Dim sPrtLabSeq      As String
    Dim sField          As String
    Dim sRstval         As String

    Dim sOtherGbn       As String   ' 출력시 판정값을 출력할찌 결정(Register에서 읽어오기)
    Dim P       As Object
    Dim jCNT            As Integer
    Dim kCNT            As Integer
    Dim TitleGbn        As String
    
    Dim ypos(100)           As Single
    
    Dim yCnt            As Integer
    Dim tmp_Cnt         As Integer
    
    Dim tmp_Part        As String
    Dim pnt_Title       As Boolean
    
    '-- Y position
    sSql = "Select PRYPOS From PRINTINFO " & _
           " Where PrintGbn  = '00' " & _
           "   and MasterGBN = '2' " & _
           "   and RealGbn   = 'H01001001NNNN' "
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        ypos(0) = AdoRecord("PRYPOS")
    Else
        ypos(0) = 54
    End If
        
    AdoRecord.Close
    
    For kCNT = 1 To 100
        ypos(kCNT) = ypos(kCNT - 1) + 4
    Next
    
    For Each P In Printers
        If P.DeviceName = "대련MTS SemiLIS Printer" Then
            Set Printer = P
            Exit For
        End If
    Next P

    Print_LabResult = 0
    sOtherGbn = ""
' 먼저 환자 신상의 출력을 셋팅

    sSelect_Item = ""   ' Select절 조립
    sFromStr1 = "": sFromStr2 = "": sFromStr3 = ""

    For iCnt = 1 To iCntPrt환자정보
        If Trim(tPrtPatient(iCnt).sField) = "LABTIME" Or Trim(tPrtPatient(iCnt).sField) = "WARD" Or Trim(tPrtPatient(iCnt).sField) = "RESULTENDDATE" Or Trim(tPrtPatient(iCnt).sField) = "RESULTENDTIME" Or Trim(tPrtPatient(iCnt).sField) = "MEDICALREMARK" Then
            sFromStr1 = "H_JUBSU"
            sFromStr = sFromStr1
        ElseIf Trim(tPrtPatient(iCnt).sField) = "DEPTNM" Or Trim(tPrtPatient(iCnt).sField) = "DEPTCD" Then
            tPrtPatient(iCnt).sField = "DEPTCD"
            sFromStr3 = "H_JUBSU"
            sFromStr = sFromStr3
        Else
            sFromStr3 = "H_RESULT"    ' Join의 기준이 되는 DB
            sFromStr = sFromStr3
        End If

        If iCnt = iCntPrt환자정보 Then
            sSelect_Item = sSelect_Item & sFromStr & "." & tPrtPatient(iCnt).sField
        Else
            sSelect_Item = sSelect_Item & sFromStr & "." & tPrtPatient(iCnt).sField & ","
        End If
    Next iCnt

    sFromStr = ""   ' From절 조립
    If sFromStr1 <> "" Then
        sFromStr = sFromStr1 & ","
    End If
    If sFromStr2 <> "" Then
        sFromStr = sFromStr & sFromStr2 & ","
    End If
    If sFromStr3 <> "" Then
        sFromStr = sFromStr & sFromStr3 & ","
    End If
    sFromStr = Left(sFromStr, Len(Trim(sFromStr)) - 1) & " "

    sWhereStr = ""      ' Where절 추가
    
    If sSlipCd = "All" Then
        If sFromStr1 <> "" Then
            sWhereStr = " AND " & sFromStr1 & ".LABDATE = '" & sLabDate & "'" & _
                        " AND " & sFromStr1 & ".PARTGBN in ('01','02','04') "
            
            For jCNT = 1 To Len(sLabSeq) Step 5
                If jCNT = 1 Then
                    sWhereStr = sWhereStr & " AND " & sFromStr1 & ".LABSEQ = '" & Mid(sLabSeq, 1, 5) & "'"
                Else
                    sWhereStr = sWhereStr & " Or " & sFromStr1 & ".LABSEQ = '" & Mid(sLabSeq, jCNT, jCNT + 4) & "'"
                End If
            Next
        End If
    
    Else
        If sFromStr1 <> "" Then
            sWhereStr = " AND " & sFromStr1 & ".LABDATE = '" & sLabDate & "'" & _
                        " AND " & sFromStr1 & ".PARTGBN = '" & Mid(sSlipCd, 2, 2) & "'" & _
                        " AND " & sFromStr1 & ".LABSEQ = '" & sLabSeq & "'"
        End If
    End If
    
    If sFromStr2 <> "" Then
        sWhereStr = sWhereStr & " AND " & sFromStr2 & ".DEPTCD = " & sFromStr1 & ".DEPTCD"
    End If
    
    On Error GoTo ErrHandler
    
    If sSlipCd = "All" Then
        sSql = " SELECT TOP 1 " & sSelect_Item & _
               "   FROM " & sFromStr & _
               "  WHERE " & sFromStr3 & ".LABDATE = '" & sLabDate & "' " & _
               "    AND " & sFromStr3 & ".PARTGBN in ('01','02','04') " & _
               "    AND " & sFromStr3 & ".LABSEQ  = '" & sLabSeq & "' "
    Else
        sSql = " SELECT TOP 1 " & sSelect_Item & _
               "   FROM " & sFromStr & _
               "  WHERE " & sFromStr3 & ".LABDATE = '" & sLabDate & "' " & _
               "    AND " & sFromStr3 & ".PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' " & _
               "    AND " & sFromStr3 & ".LABSEQ  = '" & sLabSeq & "' "
    End If
    
    sSql = sSql & sWhereStr
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
    ' SELECT_ITEM에 해당하는 내용을 취득
        For iCnt = 1 To iCntPrt환자정보
            sField = Trim(Right(Trim(tPrtPatient(iCnt).sField), 4))
            sRstval = Trim(AdoRecord(tPrtPatient(iCnt).sField) & "")
            Select Case sField
            Case "DATE"
                tPrtPatient(iCnt).sRealStr = Left(sRstval, 4) & "/" & Mid(sRstval, 5, 2) & "/" & Mid(sRstval, 7, 2)
            Case "TIME"
                tPrtPatient(iCnt).sRealStr = Left(sRstval, 2) & ":" & Mid(sRstval, 3, 2) & ":" & Mid(sRstval, 5, 2)
            Case "SEX"
                tPrtPatient(iCnt).sRealStr = Trim(sRstval)
                If Trim(tPrtPatient(iCnt).sRealStr) = "1" Or Trim(tPrtPatient(iCnt).sRealStr) = "3" Then
                    tPrtPatient(iCnt).sRealStr = "남"
                ElseIf Trim(tPrtPatient(iCnt).sRealStr) = "2" Or Trim(tPrtPatient(iCnt).sRealStr) = "4" Then
                    tPrtPatient(iCnt).sRealStr = "여"
                End If
            Case "PTCD" ' DEPT코드
                tPrtPatient(iCnt).sRealStr = ""
                If Trim(sRstval) <> "" Then
                    tPrtPatient(iCnt).sRealStr = tPrtDept(Asc(UCase$(Left$(sRstval, 1))) * 26 + Asc(UCase$(Right$(sRstval, 1)))).sDeptNm
                End If
            Case Else
                tPrtPatient(iCnt).sRealStr = sRstval
            End Select
        Next iCnt
    End If
    AdoRecord.Close

    ' 검사명을 가져오기
    sSql = "SELECT PARTGBN, TESTITEMSEQ,TEST.PRINTNM, PRN.MAXLEN, PRN.ALIGN, PRN.PRXPOS, PRN.PRYPOS " & _
            " FROM TESTITEM AS TEST, PRINTINFO AS PRN " & _
            "WHERE TEST.PARTCD = '" & Left(sSlipCd, 1) & "' " & _
            "  AND TEST.PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' " & _
            "  AND PRN.PRINTGBN = '00' " & _
            "  AND PRN.REALGBN  = 'H01001' + TESTITEMSEQ + 'NNNN' " & _
            "  AND PRN.MASTERGBN ='5' " & _
            " Order by PARTGBN, TESTITEMSEQ "
        
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        iCntPrt검사정보 = AdoRecord.RecordCount
        If iCntPrt검사정보 = 0 Then Exit Function
        ReDim p_Pos(4, AdoRecord.RecordCount)
        For iCnt = 1 To iCntPrt검사정보
            If iCnt = 1 Then
                tPrtItem(iCnt).sRealStr = "[생화학검사]  " & AdoRecord("PRINTNM") & ""
                pnt_Title = True
            ElseIf AdoRecord("PARTGBN") = "02" Then
                If pnt_Title = True Then
                    tPrtItem(iCnt).sRealStr = "[혈액학검사]  " & AdoRecord("PRINTNM") & ""
                    pnt_Title = False
                Else
                    tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
                End If
            ElseIf AdoRecord("PARTGBN") = "04" Then
                If pnt_Title = False Then
                    tPrtItem(iCnt).sRealStr = "[뇨화학검사]  " & AdoRecord("PRINTNM") & ""
                    pnt_Title = True
                Else
                    tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
                End If
            Else
                tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
            End If
            
'            If iCnt = 1 Then
'                If Trim(sSlipCd) = "C01" Then
'                    tPrtItem(iCnt).sRealStr = "[생화학검사]  " & AdoRecord("PRINTNM") & ""
'                ElseIf Trim(sSlipCd) = "H02" Then
'                    tPrtItem(iCnt).sRealStr = "[혈액학검사]  " & AdoRecord("PRINTNM") & ""
'                Else
'                    tPrtItem(iCnt).sRealStr = "[뇨화학검사]  " & AdoRecord("PRINTNM") & ""
'                End If
'            Else
'                tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
'            End If
'            tPrtItem(iCnt).sRealStr = AdoRecord("PRINTNM") & ""
            tPrtItem(iCnt).sMaxLen = AdoRecord("MAXLEN") & ""
            tPrtItem(iCnt).sAlign = AdoRecord("ALIGN") & ""
            tPrtItem(iCnt).sXPos = AdoRecord("PRXPOS") & ""
            tPrtResult(iCnt).sYPos = ypos(iCnt) & ""
            p_Pos(CInt(AdoRecord("PARTGBN")), CInt(AdoRecord("TESTITEMSEQ"))) = ypos(iCnt)
            AdoRecord.MoveNext
        Next iCnt
    End If
    AdoRecord.Close
    
    ' 환자 결과를 가져오기
    sSql = "SELECT TESTITEMSEQ,RESULT, REFMARK, MAXLEN, ALIGN, PRXPOS, PRYPOS, PARTGBN " & _
            " FROM H_RESULT AS RESULT, PRINTINFO " & _
            "WHERE LABDATE  = '" & sLabDate & "' " & _
            "  AND PARTGBN  = '" & Mid(sSlipCd, 2, 2) & "' " & _
            "  AND LABSEQ   = '" & sLabSeq & "' " & _
            "  AND PRINTGBN = '00' " & _
            "  AND REALGBN  = 'H01001' + RESULT.TESTITEMSEQ + 'NNNN' " & _
            "  AND MASTERGBN ='2' " & _
            " Order by PARTGBN, RESULT.TESTITEMSEQ "
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        iCntPrt결과정보 = AdoRecord.RecordCount
        If iCntPrt결과정보 = 0 Then Exit Function ' 결과정보가 없는 경우에 스킵
        For iCnt = 1 To iCntPrt결과정보
            tPrtResult(iCnt).sRealStr = AdoRecord("RESULT") & ""
            tPrtResult(iCnt).sFlagStr = AdoRecord("REFMARK") & ""
            tPrtResult(iCnt).sMaxLen = AdoRecord("MAXLEN") & ""
            tPrtResult(iCnt).sAlign = AdoRecord("ALIGN") & ""
            tPrtResult(iCnt).sXPos = AdoRecord("PRXPOS") & ""
            'tPrtResult(iCnt).sYPos = ypos(iCnt) & ""
            tPrtResult(iCnt).sYPos = p_Pos(CInt(AdoRecord("PARTGBN")), CInt(AdoRecord("TESTITEMSEQ"))) & ""
            AdoRecord.MoveNext
        Next iCnt
    End If
    AdoRecord.Close
    
    ' 참고치를 가져오기
    sSql = "SELECT PARTGBN, TESTITEMSEQ,(TEST.REFLOWM + ' ~ ' + TEST.REFHIGHM) AS TESTREF, PRN.MAXLEN, PRN.ALIGN, PRN.PRXPOS, PRN.PRYPOS " & _
            " FROM TESTITEM AS TEST, PRINTINFO AS PRN " & _
            "WHERE TEST.PARTCD = '" & Left(sSlipCd, 1) & "' " & _
            "  AND TEST.PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' " & _
            "  AND PRN.PRINTGBN = '00' " & _
            "  AND PRN.REALGBN  = 'H01001' + TESTITEMSEQ + 'NNNN' " & _
            "  AND PRN.MASTERGBN ='6' " & _
            " Order by PARTGBN, TESTITEMSEQ "
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        iCntPrt참고정보 = AdoRecord.RecordCount
        If iCntPrt참고정보 = 0 Then Exit Function ' 결과정보가 없는 경우에 스킵
        For iCnt = 1 To iCntPrt참고정보
            tPrtRef(iCnt).sRealStr = AdoRecord("TESTREF") & ""
            tPrtRef(iCnt).sMaxLen = AdoRecord("MAXLEN") & ""
            tPrtRef(iCnt).sAlign = AdoRecord("ALIGN") & ""
            tPrtRef(iCnt).sXPos = AdoRecord("PRXPOS") & ""
            'tPrtResult(iCnt).sYPos = ypos(iCnt) & ""
            tPrtRef(iCnt).sYPos = p_Pos(CInt(AdoRecord("PARTGBN")), CInt(AdoRecord("TESTITEMSEQ"))) & ""
            AdoRecord.MoveNext
        Next iCnt
    End If
    AdoRecord.Close
    
    ' Comment 출력 셋팅
    sSql = "SELECT COMMENT " & _
            " FROM COMMENT " & _
            "WHERE LABDATE = '" & sLabDate & "' " & _
            "  AND PARTCD = '" & Left(sSlipCd, 1) & "' " & _
            "  AND PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' "
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    sComment = ""
    If AdoRecord.RecordCount <> 0 Then
    ' 결과와 출력정보 취득
        For iCnt = 1 To AdoRecord.RecordCount
            sComment = sComment & AdoRecord("COMMENT") & " "
        Next iCnt
    End If
    AdoRecord.Close

' 실제 출력
    Printer.EndDoc
    
    Printer.FontName = "굴림체"
        
    ' 환자신상 출력 폰트
    Printer.ScaleMode = vbMillimeters
    If Trim(tPrtFontSize(1).sFontSize) = "" Then
        Printer.FontSize = 10
    Else
        Printer.FontSize = tPrtFontSize(1).sFontSize
    End If
    For iCnt = 1 To iCntPrt환자정보
        Printer.CurrentX = Val(Trim(tPrtPatient(iCnt).sXPos))
        Printer.CurrentY = Val(Trim(tPrtPatient(iCnt).sYPos))
        If Trim(tPrtPatient(iCnt).sAlign) = "R" Then
            Printer.Print Spc(Val(Trim(tPrtPatient(iCnt).sMaxLen)) - (LenH(Trim(tPrtPatient(iCnt).sRealStr)))); Trim(tPrtPatient(iCnt).sRealStr)
        ElseIf Trim(tPrtPatient(iCnt).sAlign) = "L" Then
            Printer.Print Trim(tPrtPatient(iCnt).sRealStr)
        End If
    Next iCnt
    
    '-- 검사 정보
    For iCnt = 1 To iCntPrt검사정보
        Printer.CurrentX = Val(Trim(tPrtItem(iCnt).sXPos))
        Printer.CurrentY = ypos(iCnt)
        If Trim(tPrtItem(iCnt).sAlign) = "R" Then
            Printer.Print Spc(Val(Trim(tPrtItem(iCnt).sMaxLen)) - (LenH(Trim(tPrtItem(iCnt).sRealStr)))); Trim(tPrtItem(iCnt).sRealStr);
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Spc(1); Trim(tPrtItem(iCnt).sFlagStr)
            End If
        ElseIf Trim(tPrtItem(iCnt).sAlign) = "L" Then
            Printer.Print tPrtItem(iCnt).sRealStr; Spc(Val(Trim(tPrtItem(iCnt).sMaxLen)) - (LenH(Trim(tPrtItem(iCnt).sRealStr))));
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Trim(tPrtItem(iCnt).sFlagStr)
            End If
        End If
    Next iCnt
    
    '-- 검사결과
    If Trim(tPrtFontSize(2).sFontSize) = "" Then
        Printer.FontSize = 9
    Else
        Printer.FontSize = tPrtFontSize(2).sFontSize
    End If
    
    For iCnt = 1 To iCntPrt결과정보
        Printer.CurrentX = Val(Trim(tPrtResult(iCnt).sXPos))
        'Printer.CurrentY = ypos(iCnt)
        For yCnt = 1 To 32
            If tPrtResult(iCnt).sYPos <> "" Then
                If tPrtResult(iCnt).sYPos = ypos(yCnt) Then
                    'tPrtResult(iCnt).sYPos = ypos(yCnt) & ""
                    Printer.CurrentY = tPrtResult(iCnt).sYPos
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        
        If Trim(tPrtResult(iCnt).sAlign) = "R" Then
            Printer.Print Spc(Val(Trim(tPrtResult(iCnt).sMaxLen)) - (LenH(Trim(tPrtResult(iCnt).sRealStr)))); Trim(tPrtResult(iCnt).sRealStr);
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Spc(1); Trim(tPrtResult(iCnt).sFlagStr)
            End If
        ElseIf Trim(tPrtResult(iCnt).sAlign) = "L" Then
            Printer.Print Trim(tPrtResult(iCnt).sRealStr); Spc(Val(Trim(tPrtResult(iCnt).sMaxLen)) - (LenH(Trim(tPrtResult(iCnt).sRealStr))));
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Trim(tPrtResult(iCnt).sFlagStr)
            End If
        End If
    Next iCnt

    
    '-- 참고치 정보
    For iCnt = 1 To iCntPrt참고정보
        Printer.CurrentX = Val(Trim(tPrtRef(iCnt).sXPos))
        Printer.CurrentY = ypos(iCnt)
        If Trim(tPrtRef(iCnt).sAlign) = "R" Then
            If Trim(tPrtRef(iCnt).sRealStr) <> "~" Then
                Printer.Print Spc(Val(Trim(tPrtRef(iCnt).sMaxLen)) - (LenH(Trim(tPrtRef(iCnt).sRealStr)))); Trim(tPrtRef(iCnt).sRealStr);
            End If
'            Printer.Print Spc(Val(Trim(tPrtRef(iCnt).sMaxLen)) - (LenH(Trim(tPrtRef(iCnt).sRealStr)))); Trim(tPrtRef(iCnt).sRealStr);
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Spc(1); Trim(tPrtRef(iCnt).sFlagStr)
            End If
        ElseIf Trim(tPrtRef(iCnt).sAlign) = "L" Then
            If Trim(tPrtRef(iCnt).sRealStr) <> "~" Then
                Printer.Print Trim(tPrtRef(iCnt).sRealStr); Spc(Val(Trim(tPrtRef(iCnt).sMaxLen)) - (LenH(Trim(tPrtRef(iCnt).sRealStr))));
            End If
'            Printer.Print Trim(tPrtRef(iCnt).sRealStr); Spc(Val(Trim(tPrtRef(iCnt).sMaxLen)) - (LenH(Trim(tPrtRef(iCnt).sRealStr))));
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Trim(tPrtRef(iCnt).sFlagStr)
            End If
        End If
    Next iCnt

    '-- Comment를 출력
    If Trim(tPrtFontSize(3).sFontSize) = "" Then
        Printer.FontSize = 10
    Else
        Printer.FontSize = tPrtFontSize(3).sFontSize
    End If
        
    Printer.CurrentY = Val(sYPos)
        
    For iCnt = 1 To iMaxLineComment
        Printer.CurrentX = Val(sXPos)
        If sAlignComment = "R" Then
            Printer.Print Spc(Val(Trim(Left(Trim(sComment), iMaxLenComment))) - (LenH(Trim(Left(Trim(sComment), iMaxLenComment))))); Trim(Left(Trim(sComment), iMaxLenComment))
        ElseIf sAlignComment = "L" Then
            Printer.Print Left(sComment, iMaxLenComment)
        End If
        sComment = Mid(sComment, iMaxLenComment + 1)
    Next iCnt
    
    '-- 보고차 출력
    If Trim(tPrtFontSize(1).sFontSize) = "" Then
        Printer.FontSize = 10
    Else
        Printer.FontSize = tPrtFontSize(1).sFontSize
    End If
    
    Printer.CurrentX = iUserNmXPos
    Printer.CurrentY = iUserNmYPos
    
    If sUserNmAlign = "R" Then
        Printer.Print Spc(Val(Trim(Left(Trim(sUserNm), iUserNmLen))) - (LenH(Trim(Left(Trim(sUserNm), iUserNmLen))))); Trim(Left(Trim(sUserNm), iUserNmLen))
    ElseIf sUserNmAlign = "L" Then
        Printer.Print Left(sUserNm, iUserNmLen)
    End If
    
    Printer.EndDoc
    
    Print_LabResult = 1
    
Exit Function
ErrHandler:
    MsgBox "Print_Result(출력 DLL) - " & Err.Description
    Print_LabResult = -1
    
End Function

' P1) 각각 환자당 검사결과 출력 루틴 ("00" 결과지)
Function Print_AllLabResult(sLabDate As String, sSlipCd As String, sLabSeq As String) As Integer

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    Dim sSelect_Item    As String
    Dim sFromStr        As String
    Dim sFromStr1       As String
    Dim sFromStr2       As String
    Dim sFromStr3       As String
    Dim sWhereStr       As String
    Dim sSql            As String
    Dim sPrtLabSeq      As String
    Dim sField          As String
    Dim sRstval         As String

    Dim sOtherGbn       As String   ' 출력시 판정값을 출력할찌 결정(Register에서 읽어오기)
    Dim P               As Object
    Dim jCNT            As Integer
    Dim kCNT            As Integer
    Dim TitleGbn        As String
    
    Dim ypos(100)       As Single
    Dim iM              As Integer
    Dim iCntPrt         As Integer
    Dim opt_Nxt         As Integer
    Dim st_Item         As Integer
    Dim yCnt            As Integer
    Dim tmp_Cnt         As Integer
    
    Dim tmp_Part        As String
    Dim pnt_Title       As Boolean
    
    Erase tPrtItem, tPrtRef, tPrtResult
    
    '-- Y positin
    sSql = "Select PRYPOS From PRINTINFO " & _
           " Where PrintGbn  = '00' " & _
           "   and MasterGBN = '2' " & _
           "   and RealGbn   = 'H01001001NNNN' "
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        ypos(0) = AdoRecord("PRYPOS")
    Else
        ypos(0) = 67
    End If
        
    AdoRecord.Close
    
    For kCNT = 1 To 100
        ypos(kCNT) = ypos(kCNT - 1) + 4
    Next
    
    For Each P In Printers
        If P.DeviceName = "대련MTS SemiLIS Printer" Then
            Set Printer = P
            Exit For
        End If
    Next P

    Print_AllLabResult = 0
    sOtherGbn = ""
    
    '-- 환자 신상
    sSelect_Item = ""   ' Select절 조립
    sFromStr1 = "": sFromStr2 = "": sFromStr3 = ""

    For iCnt = 1 To iCntPrt환자정보
        If Trim(tPrtPatient(iCnt).sField) = "LABTIME" Or Trim(tPrtPatient(iCnt).sField) = "WARD" Or Trim(tPrtPatient(iCnt).sField) = "RESULTENDDATE" Or Trim(tPrtPatient(iCnt).sField) = "RESULTENDTIME" Or Trim(tPrtPatient(iCnt).sField) = "MEDICALREMARK" Then
            sFromStr1 = "H_JUBSU"
            sFromStr = sFromStr1
        ElseIf Trim(tPrtPatient(iCnt).sField) = "DEPTNM" Or Trim(tPrtPatient(iCnt).sField) = "DEPTCD" Then
            tPrtPatient(iCnt).sField = "DEPTCD"
            sFromStr3 = "H_JUBSU"
            sFromStr = sFromStr3
        Else
            sFromStr3 = "H_RESULT"    ' Join의 기준이 되는 DB
            sFromStr = sFromStr3
        End If

        If iCnt = iCntPrt환자정보 Then
            sSelect_Item = sSelect_Item & sFromStr & "." & tPrtPatient(iCnt).sField
        Else
            sSelect_Item = sSelect_Item & sFromStr & "." & tPrtPatient(iCnt).sField & ","
        End If
    Next iCnt

    sFromStr = ""   ' From절 조립
    If sFromStr1 <> "" Then
        sFromStr = sFromStr1 & ","
    End If
    If sFromStr2 <> "" Then
        sFromStr = sFromStr & sFromStr2 & ","
    End If
    If sFromStr3 <> "" Then
        sFromStr = sFromStr & sFromStr3 & ","
    End If
    sFromStr = Left(sFromStr, Len(Trim(sFromStr)) - 1) & " "

    sWhereStr = ""      ' Where절 추가
    
    If sSlipCd = "All" Then
        If sFromStr1 <> "" Then
            sWhereStr = " AND " & sFromStr1 & ".LABDATE = '" & sLabDate & "'" & _
                        " AND " & sFromStr1 & ".PARTGBN in ('01','02','04') "
            
            For jCNT = 1 To Len(sLabSeq) Step 5
                If jCNT = 1 Then
                    sWhereStr = sWhereStr & " AND " & sFromStr1 & ".LABSEQ = '" & Mid(sLabSeq, 1, 5) & "'"
                Else
                    sWhereStr = sWhereStr & " Or " & sFromStr1 & ".LABSEQ = '" & Mid(sLabSeq, jCNT, jCNT + 4) & "'"
                End If
            Next
        End If
    
    Else
        If sFromStr1 <> "" Then
            sWhereStr = " AND " & sFromStr1 & ".LABDATE = '" & sLabDate & "'" & _
                        " AND " & sFromStr1 & ".PARTGBN = '" & Mid(sSlipCd, 2, 2) & "'" & _
                        " AND " & sFromStr1 & ".LABSEQ = '" & sLabSeq & "'"
        End If
    End If
    
    If sFromStr2 <> "" Then
        sWhereStr = sWhereStr & " AND " & sFromStr2 & ".DEPTCD = " & sFromStr1 & ".DEPTCD"
    End If
    
    On Error GoTo ErrHandler
    
    If sSlipCd = "All" Then
        sSql = " SELECT TOP 1 " & sSelect_Item & _
               "   FROM " & sFromStr & _
               "  WHERE " & sFromStr3 & ".LABDATE = '" & sLabDate & "' " & _
               "    AND " & sFromStr3 & ".PARTGBN in ('01','02','04') " & _
               "    AND " & sFromStr3 & ".LABSEQ  = '" & sLabSeq & "' "
    Else
        sSql = " SELECT TOP 1 " & sSelect_Item & _
               "   FROM " & sFromStr & _
               "  WHERE " & sFromStr3 & ".LABDATE = '" & sLabDate & "' " & _
               "    AND " & sFromStr3 & ".PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' " & _
               "    AND " & sFromStr3 & ".LABSEQ  = '" & sLabSeq & "' "
    End If
    
    sSql = sSql & sWhereStr
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
    ' SELECT_ITEM에 해당하는 내용을 취득
        For iCnt = 1 To iCntPrt환자정보
            sField = Trim(Right(Trim(tPrtPatient(iCnt).sField), 4))
            sRstval = Trim(AdoRecord(tPrtPatient(iCnt).sField) & "")
            Select Case sField
            Case "DATE"
                tPrtPatient(iCnt).sRealStr = Left(sRstval, 4) & "/" & Mid(sRstval, 5, 2) & "/" & Mid(sRstval, 7, 2)
            Case "TIME"
                tPrtPatient(iCnt).sRealStr = Left(sRstval, 2) & ":" & Mid(sRstval, 3, 2) & ":" & Mid(sRstval, 5, 2)
            Case "SEX"
                tPrtPatient(iCnt).sRealStr = Trim(sRstval)
                If Trim(tPrtPatient(iCnt).sRealStr) = "1" Or Trim(tPrtPatient(iCnt).sRealStr) = "3" Then
                    tPrtPatient(iCnt).sRealStr = "남"
                ElseIf Trim(tPrtPatient(iCnt).sRealStr) = "2" Or Trim(tPrtPatient(iCnt).sRealStr) = "4" Then
                    tPrtPatient(iCnt).sRealStr = "여"
                End If
            Case "PTCD" ' DEPT코드
                tPrtPatient(iCnt).sRealStr = ""
                If Trim(sRstval) <> "" Then
                    tPrtPatient(iCnt).sRealStr = tPrtDept(Asc(UCase$(Left$(sRstval, 1))) * 26 + Asc(UCase$(Right$(sRstval, 1)))).sDeptNm
                End If
            Case Else
                tPrtPatient(iCnt).sRealStr = sRstval
            End Select
        Next iCnt
    End If
    AdoRecord.Close

'** 2001.08.29(osw)
'** 검사장비수만큼 Loop하면서 배열에 저장
'** 장비가 비뀔때마다 Title Print

    '-- 2001.08.09(OSW)
    ' 검사항목과 출력 DB를 조인하여 결과 출력을 셋팅
    iCntPrt검체정보 = 0
        
    ' 검사명을 가져오기
    sSql = "SELECT PARTGBN, TESTITEMSEQ, TEST.PRINTNM, PRN.MAXLEN, PRN.ALIGN, PRN.PRXPOS, PRN.PRYPOS " & _
            " FROM TESTITEM AS TEST, PRINTINFO AS PRN " & _
            " WHERE TEST.PARTGBN IN ('01','02','04')" & _
            "   AND PRN.PRINTGBN = '00' " & _
            "   AND PRN.REALGBN  = 'H01001' + TESTITEMSEQ + 'NNNN' " & _
            "   AND PRN.MASTERGBN ='5' " & _
            "   AND TEST.PARTCD IN ('C','H','U') "
            
    sSql = sSql & " Order by PARTGBN, TESTITEMSEQ "
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        If AdoRecord.RecordCount = 0 Then Exit Function
        ReDim p_Pos(4, AdoRecord.RecordCount)
        pnt_Title = False
        '-- 검사항목
        For iCnt = 1 To AdoRecord.RecordCount
            If iCnt = 1 Then
                tPrtItem(iCnt).sRealStr = "[생화학검사]  " & AdoRecord("PRINTNM") & ""
                pnt_Title = True
            ElseIf AdoRecord("PARTGBN") = "02" Then
                If pnt_Title = True Then
                    tPrtItem(iCnt).sRealStr = "[혈액학검사]  " & AdoRecord("PRINTNM") & ""
                    pnt_Title = False
                Else
                    tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
                End If
            ElseIf AdoRecord("PARTGBN") = "04" Then
                If pnt_Title = False Then
                    tPrtItem(iCnt).sRealStr = "[뇨화학검사]  " & AdoRecord("PRINTNM") & ""
                    pnt_Title = True
                Else
                    tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
                End If
            Else
                tPrtItem(iCnt).sRealStr = Space$(14) & AdoRecord("PRINTNM") & ""
            End If
            
            tPrtItem(iCnt).sMaxLen = AdoRecord("MAXLEN") & ""
            tPrtItem(iCnt).sAlign = AdoRecord("ALIGN") & ""
            tPrtItem(iCnt).sXPos = AdoRecord("PRXPOS") & ""
            tPrtItem(iCnt).sYPos = ypos(iCnt) & ""
            p_Pos(CInt(AdoRecord("PARTGBN")), CInt(AdoRecord("TESTITEMSEQ"))) = ypos(iCnt)
            
            AdoRecord.MoveNext
        
        Next iCnt
    End If
            
    iCntPrt검체정보 = AdoRecord.RecordCount
            
    AdoRecord.Close
                
    sSql = "SELECT TESTITEMSEQ, RESULT, REFMARK, MAXLEN, ALIGN, PRXPOS, PRYPOS, PARTGBN " & _
            " FROM H_RESULT AS RESULT, PRINTINFO " & _
            "WHERE LABDATE  = '" & sLabDate & "' " & _
            "  AND PARTGBN IN ('01','02','04') " & _
            "  AND PRINTGBN = '00' " & _
            "  AND REALGBN IN ('H01001' + RESULT.TESTITEMSEQ + 'NNNN','H01002' + RESULT.TESTITEMSEQ + 'NNNN','H01004' + RESULT.TESTITEMSEQ + 'NNNN') " & _
            "  AND MASTERGBN ='2' " & _
            "  AND LABSEQ = '" & sLabSeq & "'" & _
            " Order by PARTGBN, RESULT.TESTITEMSEQ "
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        If AdoRecord.RecordCount = 0 Then Exit Function
        '-- 검사결과
        For iCnt = 1 To AdoRecord.RecordCount
            tPrtResult(iCnt).sRealStr = AdoRecord("RESULT")
            tPrtResult(iCnt).sFlagStr = AdoRecord("REFMARK") & ""
            tPrtResult(iCnt).sMaxLen = AdoRecord("MAXLEN") & ""
            tPrtResult(iCnt).sAlign = AdoRecord("ALIGN") & ""
            tPrtResult(iCnt).sXPos = AdoRecord("PRXPOS") & ""
            tPrtResult(iCnt).sYPos = p_Pos(CInt(AdoRecord("PARTGBN")), CInt(AdoRecord("TESTITEMSEQ"))) & ""
            
            AdoRecord.MoveNext
        Next iCnt
    
        AdoRecord.Close
            
        '-- 참고치
        sSql = "SELECT PARTGBN, TESTITEMSEQ,(TEST.REFLOWM + ' ~ ' + TEST.REFHIGHM) AS TESTREF, PRN.MAXLEN, PRN.ALIGN, PRN.PRXPOS, PRN.PRYPOS " & _
                " FROM TESTITEM AS TEST, PRINTINFO AS PRN " & _
                " WHERE TEST.PARTGBN IN ('01','02','04') " & _
                "  AND PRN.PRINTGBN = '00' " & _
                "  AND PRN.REALGBN IN ('H01001' + TESTITEMSEQ + 'NNNN','H01002' + TESTITEMSEQ + 'NNNN','H01004' + TESTITEMSEQ + 'NNNN') " & _
                "  AND PRN.MASTERGBN ='6' " & _
                "  AND TEST.PARTCD IN ('C','H','U') " & _
                " Order by TEST.PARTGBN, TESTITEMSEQ "

        AdoRecord.CursorType = adOpenStatic
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
        
        If AdoRecord.RecordCount <> 0 Then
            If AdoRecord.RecordCount = 0 Then Exit Function
            For iCnt = 1 To AdoRecord.RecordCount
                tPrtRef(iCnt).sRealStr = AdoRecord("TESTREF") & ""
                tPrtRef(iCnt).sMaxLen = AdoRecord("MAXLEN") & ""
                tPrtRef(iCnt).sAlign = AdoRecord("ALIGN") & ""
                tPrtRef(iCnt).sXPos = AdoRecord("PRXPOS") & ""
                'tPrtRef(iCnt).sYPos = ypos(iCnt) & ""
                tPrtRef(iCnt).sYPos = p_Pos(CInt(AdoRecord("PARTGBN")), CInt(AdoRecord("TESTITEMSEQ"))) & ""
                
                AdoRecord.MoveNext
            Next iCnt
         End If
    End If
    
    AdoRecord.Close
    iCntPrt = iCntPrt + iCntPrt검체정보
    
    ' Comment 출력 셋팅
    sSql = "SELECT COMMENT " & _
            " FROM COMMENT " & _
            "WHERE LABDATE = '" & sLabDate & "' "
            If sSlipCd = "All" Then
                sSql = sSql & "AND PARTCD in ('C','H','U') " & _
                            "  AND PARTGBN in ('01','03','04') "
            Else
                sSql = sSql & "AND PARTCD = '" & Left(sSlipCd, 1) & "' " & _
                            "  AND PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' "
            End If
                
                        
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    sComment = ""
    If AdoRecord.RecordCount <> 0 Then
    ' 결과와 출력정보 취득
        For iCnt = 1 To AdoRecord.RecordCount
            sComment = sComment & AdoRecord("COMMENT") & " "
        Next iCnt
    End If
    AdoRecord.Close

    '-- 실제 출력
    opt_Nxt = 1: st_Item = 0
Rst:
    Printer.EndDoc
    Printer.FontName = "굴림체"
    
    '-- 폰트
    Printer.ScaleMode = vbMillimeters
    If Trim(tPrtFontSize(1).sFontSize) = "" Then
        Printer.FontSize = 10
    Else
        Printer.FontSize = tPrtFontSize(1).sFontSize
    End If
    '## 환자정보 #######################################
    '-- 환자정보
    For iCnt = 1 To iCntPrt환자정보
        Printer.CurrentX = Val(Trim(tPrtPatient(iCnt).sXPos))
        Printer.CurrentY = Val(Trim(tPrtPatient(iCnt).sYPos))
        If Trim(tPrtPatient(iCnt).sAlign) = "R" Then
            Printer.Print Spc(Val(Trim(tPrtPatient(iCnt).sMaxLen)) - (LenH(Trim(tPrtPatient(iCnt).sRealStr)))); Trim(tPrtPatient(iCnt).sRealStr)
        ElseIf Trim(tPrtPatient(iCnt).sAlign) = "L" Then
            Printer.Print Trim(tPrtPatient(iCnt).sRealStr)
        End If
    Next iCnt
    '## 환자정보 End ####################################
    
    '-- 검사명/결과/참고치 폰트
    If Trim(tPrtFontSize(2).sFontSize) = "" Then
        Printer.FontSize = 9
    Else
        Printer.FontSize = tPrtFontSize(2).sFontSize
    End If
    
    '## 검사정보 #######################################
    '-- 검사 정보
    For iCnt = 1 To iCntPrt
        If iCnt + st_Item > iCntPrt Then Exit For
        Printer.CurrentX = Val(Trim(tPrtItem(iCnt).sXPos))
        Printer.CurrentY = ypos(iCnt)
        If Trim(tPrtItem(iCnt).sAlign) = "R" Then
            Printer.Print Spc(Val(Trim(tPrtItem(iCnt).sMaxLen)) - (LenH(Trim(tPrtItem(iCnt).sRealStr)))); Trim(tPrtItem(iCnt).sRealStr);
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Spc(1); Trim(tPrtItem(iCnt).sFlagStr)
            End If
        ElseIf Trim(tPrtItem(iCnt).sAlign) = "L" Then
            'Printer.Print RTrim(tPrtItem(iCnt).sRealStr); Spc(Val(Trim(tPrtItem(iCnt).sMaxLen)) - (LenH(Trim(tPrtItem(iCnt).sRealStr))));
            Printer.Print tPrtItem(iCnt + st_Item).sRealStr; Spc(Val(Trim(tPrtItem(iCnt).sMaxLen)) - (LenH(Trim(tPrtItem(iCnt).sRealStr))));
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print tPrtItem(iCnt).sFlagStr
            End If
        End If
    Next iCnt
    
    '## 검사정보 End ####################################
    
    '## 결과정보 ########################################
    For iCnt = 1 To iCntPrt
        If iCnt + st_Item > iCntPrt Then Exit For
        Printer.CurrentX = Val(Trim(tPrtResult(iCnt + st_Item).sXPos))
        'Printer.CurrentY = ypos(iCnt)
        
        For yCnt = 1 To 32
            If tPrtResult(iCnt).sYPos <> "" Then
                If tPrtResult(iCnt).sYPos = ypos(yCnt) Then
                    'tPrtResult(iCnt).sYPos = ypos(yCnt) & ""
                    Printer.CurrentY = tPrtResult(iCnt).sYPos
                    Exit For
                End If
            Else
                Exit For
            End If
        Next

'        Printer.CurrentY = IIf(tPrtResult(iCnt).sYPos = "", "0", tPrtResult(iCnt).sYPos)
        
        If Trim(tPrtResult(iCnt + st_Item).sAlign) = "R" Then
            Printer.Print Spc(Val(Trim(tPrtResult(iCnt + st_Item).sMaxLen)) - (LenH(Trim(tPrtResult(iCnt + st_Item).sRealStr)))); Trim(tPrtResult(iCnt + st_Item).sRealStr);
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Spc(1); Trim(tPrtResult(iCnt + st_Item).sFlagStr)
            End If
        ElseIf Trim(tPrtResult(iCnt).sAlign) = "L" Then
            Printer.Print Trim(tPrtResult(iCnt).sRealStr); Spc(Val(Trim(tPrtResult(iCnt).sMaxLen)) - (LenH(Trim(tPrtResult(iCnt).sRealStr))));
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Trim(tPrtResult(iCnt).sFlagStr)
            End If
        End If
    Next iCnt
    '## 결과정보 End ####################################

    '## 참고치정보 ########################################
    For iCnt = 1 To 32 'iCntPrt
        If iCnt + st_Item > iCntPrt Then st_Item = 64: Exit For
        Printer.CurrentX = Val(Trim(tPrtRef(iCnt + st_Item).sXPos))
        Printer.CurrentY = ypos(iCnt)
        If Trim(tPrtRef(iCnt + st_Item).sAlign) = "R" Then
            If Trim(tPrtRef(iCnt + st_Item).sRealStr) <> "~" Then
                Printer.Print Spc(Val(Trim(tPrtRef(iCnt + st_Item).sMaxLen)) - (LenH(Trim(tPrtRef(iCnt + st_Item).sRealStr)))); Trim(tPrtRef(iCnt + st_Item).sRealStr);
            End If
            
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Spc(1); Trim(tPrtRef(iCnt + st_Item).sFlagStr)
            End If
        ElseIf Trim(tPrtRef(iCnt).sAlign) = "L" Then
            If Trim(tPrtRef(iCnt + st_Item).sRealStr) <> "~" Then
                Printer.Print Trim(tPrtRef(iCnt).sRealStr); Spc(Val(Trim(tPrtRef(iCnt).sMaxLen)) - (LenH(Trim(tPrtRef(iCnt).sRealStr))));
            End If
            If Trim(sOtherGbn) = "N" Then
                Printer.Print
            Else
                Printer.Print Trim(tPrtRef(iCnt).sFlagStr)
            End If
        End If
    Next iCnt
    '## 참고치정보 End ####################################
    
    If st_Item < 32 And iCntPrt > 32 Then
        opt_Nxt = opt_Nxt + 1
        st_Item = 32
    End If


    '-- 항원/항체
'    For iCnt = iCntPrt To iCntPrt + 2
'        Printer.CurrentX = Val(Trim(tPrtItem(1).sXPos))
'        Printer.CurrentY = ypos(iCnt)
''        Printer.Print Trim(tPrtItem(iCnt).sRealStr); Spc(Val(Trim(tPrtItem(0).sMaxLen)) - (LenH(Trim(tPrtItem(iCnt).sRealStr))));
'
'        If iCnt = iCntPrt Then
'             Printer.Print "혈청학검사"
'        ElseIf iCnt = iCntPrt + 1 Then
'             Printer.Print "항원"
'        Else
'             Printer.Print "항체"
'        End If
'
'    Next iCnt
    
    ' Comment를 출력
    If Trim(tPrtFontSize(3).sFontSize) = "" Then
        Printer.FontSize = 10
    Else
        Printer.FontSize = tPrtFontSize(3).sFontSize
    End If
        
    Printer.CurrentY = Val(sYPos)
        
    For iCnt = 1 To iMaxLineComment
        Printer.CurrentX = Val(sXPos)
        If sAlignComment = "R" Then
            Printer.Print Spc(Val(Trim(Left(Trim(sComment), iMaxLenComment))) - (LenH(Trim(Left(Trim(sComment), iMaxLenComment))))); Trim(Left(Trim(sComment), iMaxLenComment))
        ElseIf sAlignComment = "L" Then
            Printer.Print Left(sComment, iMaxLenComment)
        End If
        sComment = Mid(sComment, iMaxLenComment + 1)
    Next iCnt
    
    ' 보고차 출력
    If Trim(tPrtFontSize(1).sFontSize) = "" Then
        Printer.FontSize = 10
    Else
        Printer.FontSize = tPrtFontSize(1).sFontSize
    End If
    
    Printer.CurrentX = iUserNmXPos
    Printer.CurrentY = iUserNmYPos
    
    If sUserNmAlign = "R" Then
        Printer.Print Spc(Val(Trim(Left(Trim(sUserNm), iUserNmLen))) - (LenH(Trim(Left(Trim(sUserNm), iUserNmLen))))); Trim(Left(Trim(sUserNm), iUserNmLen))
    ElseIf sUserNmAlign = "L" Then
        Printer.Print Left(sUserNm, iUserNmLen)
    End If
    
    Printer.EndDoc
    
    If iCntPrt > 32 And st_Item = 32 Then
        GoTo Rst
    End If
    Print_AllLabResult = 1
    
Exit Function
ErrHandler:
    MsgBox "Print_Result(출력 DLL) - " & Err.Description
    Print_AllLabResult = -1
    
End Function


' P2) 출력을 관장하는 MAIN부분
Public Function Print_Result(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String) As Integer

    Dim sLabSeq     As String
    Dim sLabSeqN    As String
    Dim sPerLabSeq  As String
    Dim iLabSeqCnt  As Integer
    Dim iCnt        As Integer
    
' 출력의 기초자료를 취득함 ("00" 결과지용)
    If Get_Print_Info("00", sSlipCd) = 0 Then
        Print_Result = -1
        Exit Function
    End If
    
' 출력해야 할 모든 레코드 갯수를 파악
    ' 검체번호가 다르면 구간에 해당하는 검체번호 취득
    sLabSeqN = Get_LabResult_Info(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn)
    iLabSeqCnt = Val(GetByOne(sLabSeqN, sLabSeqN))

    For iCnt = 1 To iLabSeqCnt
    ' Loop를 이용 하여 한명의 환자의 결과를 출력
        sPerLabSeq = GetByOne(sLabSeqN, sLabSeqN)
        If Print_LabResult(sLabDate, sSlipCd, sPerLabSeq) > 0 Then
            If Edit_PrintFlag(sLabDate, sSlipCd, sPerLabSeq) <> "SUCCESS" Then
                Print_Result = -1
            End If
        Else
            Print_Result = -1
        End If
        'Printer.NewPage
    Next iCnt

    If iLabSeqCnt > 0 Then
        'Printer.EndDoc
        Print_Result = iLabSeqCnt
    Else
        Print_Result = 0
    End If
    
End Function

' P2) 출력을 관장하는 MAIN부분
Public Function Print_AllResult(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String) As Integer

    Dim sLabSeq     As String
    Dim sLabSeqN    As String
    Dim sPerLabSeq  As String
    Dim iLabSeqCnt  As Integer
    Dim iCnt        As Integer
    
' 출력의 기초자료를 취득함 ("00" 용)
    If Get_Print_Info("00", sSlipCd) = 0 Then
        Print_AllResult = -1
        Exit Function
    End If
    
' 출력해야 할 모든 레코드 갯수를 파악
    ' 검체번호가 다르면 구간에 해당하는 검체번호 취득
    sLabSeqN = Get_LabResult_Info(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn)
    iLabSeqCnt = Val(GetByOne(sLabSeqN, sLabSeqN))

    sPerLabSeq = ""
    If sSlipCd = "All" Then
        For iCnt = 1 To iLabSeqCnt
            sPerLabSeq = GetByOne(sLabSeqN, sLabSeqN)
        'Next
    
            'sPerLabSeq = GetByOne(sLabSeqN, sLabSeqN)
            If Print_AllLabResult(sLabDate, sSlipCd, sPerLabSeq) > 0 Then
                If Edit_PrintFlag(sLabDate, sSlipCd, sPerLabSeq) <> "SUCCESS" Then
                    Print_AllResult = -1
                End If
            Else
                Print_AllResult = -1
            End If
        Next
    Else
        For iCnt = 1 To iLabSeqCnt
        ' Loop를 이용 하여 한명의 환자의 결과를 출력
            sPerLabSeq = GetByOne(sLabSeqN, sLabSeqN)
            If Print_LabResult(sLabDate, sSlipCd, sPerLabSeq) > 0 Then
                If Edit_PrintFlag(sLabDate, sSlipCd, sPerLabSeq) <> "SUCCESS" Then
                    Print_AllResult = -1
                End If
            Else
                Print_AllResult = -1
            End If
            'Printer.NewPage
        Next iCnt
    End If
    
    If iLabSeqCnt > 0 Then
        'Printer.EndDoc
        Print_AllResult = iLabSeqCnt
    Else
        Print_AllResult = 0
    End If
    
End Function

' 작업대장 출력관련

' S3) 루틴작업 조회
Function Get_RoutineT(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String) As String

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim sRoutineT       As String
    Dim sSql            As String
    Dim sJubSuGbn       As String
    Dim sEmGbn          As String
    Dim sDeptCd         As String
    Dim sSpcNm          As String
    Dim sSex            As String
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    
    sRoutineT = ""
    
    On Error GoTo ErrHandler
    
    sSql = " SELECT DISTINCT RESULT.LABDATE+'-'+RESULT.LABSEQ AS LABNO, JUBSU.JUBSUGBN, JUBSU.EMERGBN, JUBSU.DEPTCD, JUBSU.WARD, RESULT.REGNO, RESULT.NAME, RESULT.SEX+'/'+RESULT.AGE AS SEXAGE, RESULT.ROUTINECD " & _
           "   FROM H_RESULT AS RESULT, H_JUBSU AS JUBSU " & _
           "  WHERE RESULT.LABDATE ='" & sLabDate & "' " & _
           "    AND RESULT.PARTGBN ='" & Mid(sSlipCd, 2, 2) & "' " & _
           "    AND RESULT.LABSEQ BETWEEN '" & sLabSeqS & "' AND '" & sLabSeqE & "' " & _
           "    AND JUBSU.LABDATE  = RESULT.LABDATE " & _
           "    AND JUBSU.PARTGBN  = RESULT.PARTGBN " & _
           "    AND JUBSU.LABSEQ   = RESULT.LABSEQ"
    If sPrtGbn = "0" Then
    ElseIf sPrtGbn = "1" Then
        sSql = sSql & "  ORDER BY RESULT.REGNO "
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 작업Rtn갯수 | LABNO(작업번호) | 접수구분 | 응급구분 | 진료과 | 병동
'                    | 등록번호 | 환자명 | 성별/나이 | 루틴코드 |
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        sRoutineT = ""
        
        For iCnt = 1 To RCnt
            sRoutineT = sRoutineT & AdoRecord("LABNO") & "|"
            
            Select Case Trim(AdoRecord("JUBSUGBN") & "")
            Case "0"
                sJubSuGbn = "외래"
            Case "1"
                sJubSuGbn = "입원"
            Case "2"
                sJubSuGbn = "수탁"
            Case Else
                sJubSuGbn = ""
            End Select
            sRoutineT = sRoutineT & sJubSuGbn & "|"
            
            Select Case Trim(AdoRecord("EMERGBN") & "")
            Case "0"
                sEmGbn = "(N)"
            Case "1"
                sEmGbn = "(E)"
            Case Else
                sEmGbn = ""
            End Select
            sRoutineT = sRoutineT & sEmGbn & "|"
            
            sDeptCd = Trim(AdoRecord("DEPTCD") & "")
            If sDeptCd <> "" Then
                sRoutineT = sRoutineT & _
                    Trim(tPrtDept(Asc(UCase$(Left$(sDeptCd, 1))) * 26 + Asc(UCase$(Right$(sDeptCd, 1)))).sDeptNm) & "|"
            Else
                sRoutineT = sRoutineT & "|"
            End If
            
            sRoutineT = sRoutineT & AdoRecord("WARD") & "|"
            sRoutineT = sRoutineT & AdoRecord("REGNO") & "|"
            sRoutineT = sRoutineT & AdoRecord("NAME") & "|"
            sSex = AdoRecord("SEXAGE") & ""
            
            If Left(sSex, 1) = "1" Or Left(sSex, 1) = "3" Then
                sSex = "남" & Mid(sSex, 2)
            ElseIf Left(sSex, 1) = "2" Or Left(sSex, 1) = "4" Then
                sSex = "여" & Mid(sSex, 2)
            Else
                sSex = ""
            End If
            
            sRoutineT = sRoutineT & sSex & "|"
            
            sSpcNm = Trim(AdoRecord("ROUTINECD") & "")
            If sSpcNm <> "" Then
                sRoutineT = sRoutineT & tPrtRtcd(Val(Mid(sSpcNm, 2))).sRtNm & "|"
            Else
                sRoutineT = sRoutineT & "|"
            End If
            
            AdoRecord.MoveNext
        Next iCnt
        sRoutineT = Str(RCnt) & "|" & sRoutineT
    End If
    AdoRecord.Close
    
    Get_RoutineT = sRoutineT
   
Exit Function
ErrHandler:
    MsgBox "GetItem(출력 DLL) - " & Err.Description
    AdoRecord.Close

End Function

' S4) 낱개검사 조회
Function Get_OrdInfoT(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String) As String
    
    Dim AdoRecord       As New ADODB.Recordset
    
    Dim sOrdInfoT       As String
    Dim sSql            As String
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    
    sOrdInfoT = ""
    
    On Error GoTo ErrHandler

    sSql = " SELECT  RESULT.LABDATE+'-'+RESULT.LABSEQ AS LABNO, ITEM.PRINTNM, ITEM.SUBMCD " & _
           "   FROM H_RESULT AS RESULT, TESTITEM AS ITEM " & _
           "  WHERE ITEM.PARTCD = '" & Left(sSlipCd, 1) & "' " & _
           "    AND ITEM.PARTGBN = RESULT.PARTGBN " & _
           "    AND ITEM.SPECIMENCD = RESULT.SPECIMENCD " & _
           "    AND ITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ " & _
           "    AND ITEM.SUBMCD = RESULT.SUBMCD " & _
           "    AND ITEM.SUBMCD IN ('NNNN', '00NN') " & _
           "    AND RESULT.LABDATE='" & sLabDate & "' " & _
           "    AND RESULT.PARTGBN='" & Mid(sSlipCd, 2, 2) & "' " & _
           "    AND RESULT.LABSEQ BETWEEN '" & sLabSeqS & "' AND '" & sLabSeqE & "' " & _
           "    AND (RESULT.ROUTINECD IS NULL " & _
           "      OR RESULT.ROUTINECD = '' ) "
    If sPrtGbn = "0" Then
        sSql = sSql & "  ORDER BY RESULT.LABDATE, RESULT.LABSEQ, ITEM.TESTITEMSEQ,ITEM.SUBMCD "
    ElseIf sPrtGbn = "1" Then
        sSql = sSql & "  ORDER BY RESULT.REGNO "
    End If

    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 작업낱개갯수 | LABNO(작업번호) | 검사항목명 |
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        sOrdInfoT = ""
        
        For iCnt = 1 To RCnt
            sOrdInfoT = sOrdInfoT & AdoRecord("LABNO") & "|"
            sOrdInfoT = sOrdInfoT & AdoRecord("PRINTNM") & "|"
            AdoRecord.MoveNext
        Next iCnt
        sOrdInfoT = Str(RCnt) & "|" & sOrdInfoT
    End If
    AdoRecord.Close
    
    Get_OrdInfoT = sOrdInfoT
   
Exit Function
ErrHandler:
    MsgBox "GetItem(출력 DLL) - " & Err.Description
    AdoRecord.Close

End Function

' S5) 슬립명 조회
Function Get_SlipNm(sSlipCd As String) As String

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim sSlipNm         As String
    Dim sSql            As String
    
    sSlipNm = ""
    
    On Error GoTo ErrHandler

    sSql = " SELECT SLIPNM " & _
           "   FROM PART " & _
           "  WHERE PARTCD = '" & Left(sSlipCd, 1) & "' " & _
           "    AND PARTGBN = '" & Mid(sSlipCd, 2, 2) & "' "
           
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 슬립명 |
        AdoRecord.MoveFirst
         
        sSlipNm = AdoRecord("SLIPNM")
    End If
    AdoRecord.Close
    
    Get_SlipNm = sSlipNm
   
Exit Function
ErrHandler:
    MsgBox "GetItem(출력 DLL) - " & Err.Description

End Function

' P4) 작업대장 출력 관리 MAIN부분
Function Print_WorkSheet(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String) As Integer

    Dim iCnt        As Integer
    Dim iLabSeqN    As Integer
    
' 출력의 기초자료를 취득함 ("01" 작업대장)
    If Get_Print_Info("01", sSlipCd) = 0 Then
        Print_WorkSheet = -1
        Exit Function
    End If
    
' 출력해야 할 모든 레코드 갯수를 파악 sPartGbn은 정렬구분
    'iLabSeqN = Print_Work작업대장(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn, 140)
    'KTY modify 19990305
    iLabSeqN = Print_Work작업대장(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn, 190)
    
    If iLabSeqN > 0 Then
        Printer.EndDoc
        Print_WorkSheet = iLabSeqN
    Else
        Print_WorkSheet = 0
    End If

End Function

' P5) 실제 작업대장 출력
Function Print_Work작업대장(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String, iLineLen As Integer) As Integer
   
    Dim iTotalCnt   As Integer
    Dim iOtotalcnt  As Integer
    Dim iRcnt       As Integer
    Dim iOCnt       As Integer
    Dim sPRoutine   As String
    Dim sPOrd       As String
    
    Dim sOldLabNo   As String
    Dim sOLabNo     As String
    Dim sOrdNm      As String
    
    Dim sLabNo      As String
    Dim sJubSuGbn   As String
    Dim sEmGbn      As String
    Dim sDeptCd     As String
    Dim sWard       As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sSexAge     As String
    Dim sRtCd       As String
           
    ' 슬립명을 취득
    sSlipNm = Get_SlipNm(sSlipCd)
' SQL문 2개 Routine와 실제 낱개 검사항목
    ' 신상과 루틴 Query ' 만약 루틴이 존재하지 않으면, Null로 신상만 출력
    sPRoutine = Get_RoutineT(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn)
    ' 낱개검사 Query
    sPOrd = Get_OrdInfoT(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn)
    
' LOOP을 하면서, 신상과 루틴이 있으면 출력하고, 없으면 낱개검사를 출력한다.
    If sPRoutine <> "" Then
        iTotalCnt = Val(GetByOne(sPRoutine, sPRoutine))
        iOtotalcnt = Val(GetByOne(sPOrd, sPOrd))
        'iPageLineMax = 80  ' 나누기를 해야 함으로 20Line
        iPageLineMax = 120  ' 나누기를 해야 함으로 30Line
        iColMax = 10
        iColNow = 0
        iNoCnt = 1
        iNowRow = 1
        sOldLabNo = ""
        
        For iRcnt = 1 To iTotalCnt
            sLabNo = GetByOne(sPRoutine, sPRoutine)
            sJubSuGbn = GetByOne(sPRoutine, sPRoutine)
            sEmGbn = GetByOne(sPRoutine, sPRoutine)
            sDeptCd = GetByOne(sPRoutine, sPRoutine)
            sWard = GetByOne(sPRoutine, sPRoutine)
            sRegno = GetByOne(sPRoutine, sPRoutine)
            sName = GetByOne(sPRoutine, sPRoutine)
            sSexAge = GetByOne(sPRoutine, sPRoutine)
            sRtCd = GetByOne(sPRoutine, sPRoutine)
        
            sDate = Left(sLabNo, 8)
        
            'Printer.FontSize = 9
            'KTY modify 19990305
            Printer.FontSize = 10
            If sOldLabNo <> sLabNo Then
                Call Print_CheckPage_작업대장(iPageLineMax, iNowRow - 1, iLineLen)
                If sOldLabNo = "" Then          ' 가장 처음 출력
                    'If iColNow = 0 Then        ' 신상 까지 출력
                    ' 인쇄
                        'Printer.FontSize = 9
                        'KTY modify 19990305
                        Printer.FontSize = 10
                        Printer.Print Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));    ' No Seq
                        iNoCnt = iNoCnt + 1
                        Printer.Print Spc(1); Trim(sLabNo);                             ' 작업번호
                        'Printer.Print Spc(1); Trim(sJubSuGbn);                          ' 접수구분
                        'KTY modify 19990305
                        Printer.Print Spc(1); Trim(sJubSuGbn); Spc(4 - LenH(Trim(sJubSuGbn)));   ' 접수구분
                        If Trim(sEmGbn) <> "" Then                                      ' 응급구분
                            Printer.Print Trim(sEmGbn);
                        Else
                            'Printer.Print Spc(6);
                            'KTY modify 19990305
                            Printer.Print Spc(4);
                        End If
                        'Printer.Print Spc(1); Trim(sDeptCd); "/"; Trim(sWard); Spc(30 - (LenH(sDeptCd) + LenH(sWard)));
                        'Printer.Print Spc(1); Trim(sRegno); Spc(15 - Len(Trim(sRegno)));
                        'Printer.Print Spc(1); Trim(sName); Spc(15 - LenH(Trim(sName)));
                        'KTY modify 19990305
                        
                        If LenH(sDeptCd) > 4 Then
                            sDeptCd = LeftH(sDeptCd, 4)
                        End If
                        If LenH(sWard) > 4 Then
                            sWard = LeftH(sWard, 4)
                        End If
                        If LenH(sDeptCd) + LenH(sWard) > 7 Then
                            sWard = LeftH(sWard, (7 - 1) - LenH(sDeptCd))
                        End If
                        
                        Printer.Print Spc(1); Trim(sDeptCd); "/"; Trim(sWard); Spc((7 - 1) - (LenH(sDeptCd) + LenH(sWard)));
                        
                        Printer.Print Spc(1); Trim(sRegno); Spc(8 - Len(Trim(sRegno)));
                        
                        If LenH(sName) > 8 Then
                            sName = LeftH(sName, 8)
                        End If
                        Printer.Print Spc(1); Trim(sName); Spc(8 - LenH(Trim(sName)));
                        
                        If sSexAge <> "" Then
                            Printer.Print Spc(1); Trim(sSexAge); Spc(9 - LenH(Trim(sSexAge)));
                        Else
                            Printer.Print Spc(10);
                        End If
                        If Trim(sRtCd) <> "" Then
                            Printer.Print Spc(1); LeftH(Trim(sRtCd), 10); Spc(10 - LenH(Trim(Left(sRtCd, 10))));
                            iColNow = iColNow + 1
                        End If
                    'End If
                Else
                ' 낱개 검사 항목 출력
                    If sPrtGbn = "0" Then
                        For iOCnt = 1 To iOtotalcnt
                            If sPOrd <> "" Then
                                sOLabNo = GetByOne(sPOrd, sPOrd)
                                sOrdNm = GetByOne(sPOrd, sPOrd)
                                If sOLabNo = sOldLabNo Then
                                ' 낱개 검사항목 출력
                                    If iColNow = iColMax Then
                                        iColNow = 0
                                        ' Line위치를 체크 (NewPage를 검사)
                                        Printer.Print
                                        Call Print_CheckPage_작업대장(iPageLineMax, iNowRow - 1, iLineLen)
                                        'Printer.FontSize = 9
                                        'Printer.Print Spc(1); Spc(103); LeftH(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(LeftH(sOrdNm, 10))));
                                        'KTY modify 19990305
                                        Printer.FontSize = 10
                                        Printer.Print Spc(1); Spc(63); LeftH(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(LeftH(sOrdNm, 10))));
                                    Else
                                        Printer.Print Spc(1); LeftH(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(LeftH(sOrdNm, 10))));
                                    End If
                                    iColNow = iColNow + 1
                                ElseIf sOLabNo > sOldLabNo Then
                                ' 원상복귀
                                    sPOrd = sOLabNo & "|" & sOrdNm & "|" & sPOrd
                                    Exit For
                                End If
                            End If
                        Next iOCnt
                    ElseIf sPrtGbn = "1" Then
                        For iOCnt = 1 To iOtotalcnt
                            If sPOrd <> "" Then
                                sOLabNo = GetByOne(sPOrd, sPOrd)
                                sOrdNm = GetByOne(sPOrd, sPOrd)
                                If sOLabNo = sOldLabNo Then
                                ' 낱개 검사항목 출력
                                    If iColNow = iColMax Then
                                        iColNow = 0
                                        ' Line위치를 체크 (NewPage를 검사)
                                        Printer.Print
                                        Call Print_CheckPage_작업대장(iPageLineMax, iNowRow - 1, iLineLen)
                                        'Printer.FontSize = 9
                                        'Printer.Print Spc(1); Spc(103); LeftH(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(LeftH(sOrdNm, 10))));
                                        'KTY modify 19990305
                                        Printer.FontSize = 10
                                        Printer.Print Spc(1); Spc(63); LeftH(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(LeftH(sOrdNm, 10))));
                                    Else
                                        Printer.Print Spc(1); LeftH(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(LeftH(sOrdNm, 10))));
                                    End If
                                    iColNow = iColNow + 1
                                ElseIf sOLabNo <> sOldLabNo Then
                                ' 원상복귀
                                    sPOrd = sOLabNo & "|" & sOrdNm & "|" & sPOrd
                                    Exit For
                                End If
                            End If
                        Next iOCnt
                    End If
                ' 다음 신상 출력
                    iColNow = 0
                    Printer.Print
                    'KTY Add 19990305
                    Printer.FontSize = 10
                    Printer.Print String(iLineLen, Chr(45))
                    Call Print_CheckPage_작업대장(iPageLineMax, iNowRow - 1, iLineLen)
                    'Printer.FontSize = 9
                    'KTY modify 19990305
                    Printer.FontSize = 10
                    Printer.Print Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));    ' No Seq
                    iNoCnt = iNoCnt + 1
                    Printer.Print Spc(1); Trim(sLabNo);                             ' 작업번호
                    'Printer.Print Spc(1); Trim(sJubSuGbn);                          ' 접수구분
                    'KTY modify 19990305
                    Printer.Print Spc(1); Trim(sJubSuGbn); Spc(4 - LenH(Trim(sJubSuGbn))); ' 접수구분
                    
                    If Trim(sEmGbn) <> "" Then                                      ' 응급구분
                        Printer.Print Trim(sEmGbn);
                    Else
                        'Printer.Print Spc(6);
                        'KTY modify 19990305
                        Printer.Print Spc(4);
                    End If
                    'Printer.Print Spc(1); Trim(sDeptCd); "/"; Trim(sWard); Spc(30 - (LenH(sDeptCd) + LenH(sWard)));
                    'Printer.Print Spc(1); Trim(sRegno); Spc(15 - Len(Trim(sRegno)));
                    'Printer.Print Spc(1); Trim(sName); Spc(15 - LenH(Trim(sName)));
                    'KTY modify 19990305
                    
                    If LenH(sDeptCd) > 4 Then
                        sDeptCd = LeftH(sDeptCd, 4)
                    End If
                    If LenH(sWard) > 4 Then
                        sWard = LeftH(sWard, 4)
                    End If
                    If LenH(sDeptCd) + LenH(sWard) > 7 Then
                        sWard = LeftH(sWard, (7 - 1) - LenH(sDeptCd))
                    End If
                    
                    Printer.Print Spc(1); Trim(sDeptCd); "/"; Trim(sWard); Spc((7 - 1) - (LenH(sDeptCd) + LenH(sWard)));
                    
                    Printer.Print Spc(1); Trim(sRegno); Spc(8 - Len(Trim(sRegno)));
                    
                    If LenH(sName) > 8 Then
                        sName = LeftH(sName, 8)
                    End If
                    Printer.Print Spc(1); Trim(sName); Spc(8 - LenH(Trim(sName)));
                    
                    If sSexAge <> "" Then
                        Printer.Print Spc(1); Trim(sSexAge); Spc(9 - LenH(Trim(sSexAge)));
                    Else
                        Printer.Print Spc(10);
                    End If
                    If Trim(sRtCd) <> "" Then
                        Printer.Print Spc(1); LeftH(Trim(sRtCd), 10); Spc(10 - LenH(Trim(Left(sRtCd, 10))));
                        iColNow = iColNow + 1
                    End If
                End If
                sOldLabNo = sLabNo
            ' 낱개 출력 종료시 루틴 새로이 출력
            Else
            ' 루틴 검사 출력
                If Trim(sRtCd) <> "" Then
                    If iColNow = iColMax Then
                        iColNow = 0
                        Printer.Print
                        Call Print_CheckPage_작업대장(iPageLineMax, iNowRow - 1, iLineLen)
                        ' 인쇄
                        'Printer.FontSize = 9
                        'Printer.Print Spc(1); Spc(103); LeftH(Trim(sRtCd), 10); Spc(10 - LenH(Trim(Left(sRtCd, 10))));
                        'KTY modify 19990305
                        Printer.FontSize = 10
                        Printer.Print Spc(1); Spc(63); LeftH(Trim(sRtCd), 10); Spc(10 - LenH(Trim(Left(sRtCd, 10))));
                        iColNow = iColNow + 1
                    Else
                        Printer.Print Spc(1); LeftH(Trim(sRtCd), 10); Spc(10 - LenH(Trim(Left(sRtCd, 10))));
                        iColNow = iColNow + 1
                    End If
                End If
            End If
        Next iRcnt
        ' 최종 낱개 항목 있는지 판단
        ' 낱개 검사 항목 출력
        For iOCnt = 1 To iOtotalcnt
            If sPOrd <> "" Then
                sOLabNo = GetByOne(sPOrd, sPOrd)
                sOrdNm = GetByOne(sPOrd, sPOrd)
                If sOLabNo = sLabNo Then
                ' 낱개 검사항목 출력
                    If iColNow = iColMax Then
                        iColNow = 0
                        ' Line위치를 체크 (NewPage를 검사)
                        Printer.Print
                        Call Print_CheckPage_작업대장(iPageLineMax, iNowRow - 1, iLineLen)
                        'Printer.Print Spc(1); Spc(103); Left(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(Left(sOrdNm, 10))));
                        'KTY modify 19990305
                        Printer.Print Spc(1); Spc(63); Left(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(Left(sOrdNm, 10))));
                    Else
                        Printer.Print Spc(1); Left(Trim(sOrdNm), 10); Spc(10 - LenH(Trim(Left(sOrdNm, 10))));
                    End If
                    iColNow = iColNow + 1
                Else
                ' 원상복귀
                    sPOrd = sOLabNo & "|" & sOrdNm & "|"
                End If
            End If
        Next iOCnt
        Call Print_Tail(iLineLen)
    End If
    Print_Work작업대장 = iNoCnt - 1
    
End Function

' P6) 한 페이지 당 HEAD출력(작업대장)
Sub Print_Head_작업대장(iLineLen As Integer)

    Dim sLabNo  As String
    Dim P       As Object

    For Each P In Printers
        If P.DeviceName Like "*대장*" Then
            Set Printer = P
            Exit For
        End If
    Next P
    
    sLabNo = sDate
    sLabNo = Left(sLabNo, 4) & "-" & Mid(sLabNo, 5, 2) & "-" & Mid(sLabNo, 7, 2)
    
    Printer.FontName = "굴림체"
' 메인 TITLE출력
    Printer.FontSize = 16
    'Printer.Print Spc(40); "SLIP별 작업대장"
    'KTY modify 19990305
    Printer.Print Spc(50); "SLIP별 작업대장"
' 상위 HEAD
    Printer.FontSize = 11
    'Printer.Print "접수일자 : "; sLabNo; Spc(85 - Len(Trim(sLabNo))); "출력일자 : "; Format(Now, "YYYY-MM-DD")
    'Printer.Print "SLIP  명 : "; sSlipNm; Spc(85 - LenH(Trim(sSlipNm))); "PAGE     : "; Printer.Page
    'KTY modify 199990305
    Printer.Print "접수일자 : "; sLabNo; Spc(135 - Len(Trim(sLabNo))); "출력일자 : "; Format(Now, "YYYY-MM-DD")
    Printer.Print "SLIP  명 : "; sSlipNm; Spc(135 - LenH(Trim(sSlipNm))); "PAGE     : "; Printer.Page
' LINE 출력
    Printer.FontSize = 10
    Printer.Print String(iLineLen, Chr(6))
' 하위 HEAD
    'Printer.FontSize = 9
    'KTY modify 19990305
    Printer.FontSize = 10
    'Printer.Print "No."; Spc(1); "작업번호"; Spc(7); "접수구분"; Spc(3); "과/병동"; Spc(25); "등록번호"; Spc(8); "이름"; Spc(12); "성별/나이"; Spc(1); "검 사 항 목"
    'KTY modify 19990305
    Printer.Print "No."; Spc(1); "작업번호"; Spc(7); "접수구분"; Spc(1); "과/병동"; Spc(1); "등록번호"; Spc(1); "이름"; Spc(5); "성별/나이"; Spc(1); "검 사 항 목"
' LINE 출력
    Printer.FontSize = 10
    Printer.Print String(iLineLen, Chr(6))

End Sub

' P7) 한 페이지 당 TAIL출력(공통)
Sub Print_Tail(iLineLen As Integer)

' LINE 출력
    Printer.Print
    Printer.FontSize = 10
    Printer.Print String(iLineLen, Chr(6))
' TAIL 내용 출력
    Printer.Print fCurAppTitle
    
End Sub

' P8) 작업대장 PageSkip Check
Sub Print_CheckPage_작업대장(iPageLineMax As Integer, iLNowRow As Integer, iLineLen As Integer)
    
    If (iLNowRow <> 0) And (iLNowRow Mod (iPageLineMax / 2) = 0) Then
        Call Print_Tail(iLineLen)      ' TAIL인쇄후 HEAD인쇄
        Printer.NewPage
        Call Print_Head_작업대장(iLineLen)
    ElseIf iLNowRow = 0 Then
        Call Print_Head_작업대장(iLineLen)   ' HEAD만 인쇄
    End If
    iNowRow = iLNowRow + 2              ' 현재 Line번호 증가
End Sub

' 결과대장

' S6) 신상과 검사결과를 취득
Function Get_ResultT(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn) As String

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim sResultT        As String
    Dim sPerResultT     As String
    Dim sSql            As String
    Dim RCnt            As Integer
    Dim PCnt            As Integer
    Dim iCnt            As Integer
    
    Dim sJubSuGbn       As String
    Dim sEmGbn          As String
    Dim sDeptCd         As String
    Dim sSex            As String
    Dim sLabNo          As String
    Dim sOldLabNo       As String
    
    sResultT = ""
    
    On Error GoTo ErrHandler

    sSql = " SELECT  RESULT.LABDATE+'-'+RESULT.LABSEQ AS LABNO, JUBSU.JUBSUGBN, JUBSU.EMERGBN, JUBSU.DEPTCD, JUBSU.WARD, " & _
           "         RESULT.REGNO, RESULT.NAME, RESULT.SEX + '/' + RESULT.AGE AS SEXAGE, ITEM.PRINTNM, ITEM.SUBMCD, RESULT.RESULT " & _
           "   FROM H_RESULT AS RESULT, H_JUBSU AS JUBSU, TESTITEM AS ITEM "
    
        If Trim(sSlipCd) = "All" Then
            sSql = sSql & "  WHERE ITEM.PARTCD in ('C','H','U') "
        
            sSql = sSql & "  AND ITEM.PARTGBN     = RESULT.PARTGBN " & _
                          "  AND RESULT.PARTGBN   = JUBSU.PARTGBN " & _
                          "  AND ITEM.SPECIMENCD  = RESULT.SPECIMENCD " & _
                          "  AND ITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ " & _
                          "  AND JUBSU.LABDATE    = RESULT.LABDATE " & _
                          "  AND JUBSU.LABSEQ     = RESULT.LABSEQ " & _
                          "  AND RESULT.LABDATE   = '" & sLabDate & "' "

            sSql = sSql & "  AND RESULT.PARTGBN in ('01','02','04') "
        
        Else
            sSql = sSql & "  WHERE ITEM.PARTCD    = '" & Left(sSlipCd, 1) & "' "
            
            sSql = sSql & "  AND ITEM.PARTGBN     = RESULT.PARTGBN " & _
                          "  AND RESULT.PARTGBN   = JUBSU.PARTGBN " & _
                          "  AND ITEM.SPECIMENCD  = RESULT.SPECIMENCD " & _
                          "  AND ITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ " & _
                          "  AND JUBSU.LABDATE    = RESULT.LABDATE " & _
                          "  AND JUBSU.LABSEQ     = RESULT.LABSEQ " & _
                          "  AND RESULT.LABDATE   = '" & sLabDate & "' "
        
            sSql = sSql & "    AND RESULT.PARTGBN   = '" & Mid(sSlipCd, 2, 2) & "' "
        
        End If
                
        sSql = sSql & "    AND RESULT.LABSEQ BETWEEN '" & sLabSeqS & "' AND '" & sLabSeqE & "' "
           
    If sPrtGbn = "0" Then
        'sSql = sSql & "  ORDER BY ITEM.PARTCD, RESULT.PARTGBN, RESULT.LABDATE, RESULT.LABSEQ, ITEM.TESTITEMSEQ,ITEM.SUBMCD "
        sSql = sSql & "  ORDER BY RESULT.LABSEQ , RESULT.PARTGBN, ITEM.TESTITEMSEQ  "
    ElseIf sPrtGbn = "1" Then
        sSql = sSql & "  ORDER BY ITEM.PARTCD, RESULT.PARTGBN, RESULT.REGNO, RESULT.LABDATE, RESULT.LABSEQ, ITEM.TESTITEMSEQ, ITEM.SUBMCD "
    End If

    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 작업총갯수 | 검체당 검사항목갯수| LABNO(작업번호) | 접수구분 |
'        응급구분 | 과 | 병동 | 등록번호 | 이름 | 성별/나이 | {검사항목명 | 검사결과 | }n
         
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        sResultT = ""
        sPerResultT = ""
        sOldLabNo = ""
        PCnt = 0
    
        For iCnt = 1 To RCnt
            
            sLabNo = AdoRecord("LABNO") & ""
            
            If (sLabNo <> sOldLabNo) And (sOldLabNo = "") Then
                sPerResultT = sPerResultT & sLabNo & "|"
                
                Select Case Trim(AdoRecord("JUBSUGBN") & "")
                Case "0"
                    sJubSuGbn = "외래"
                Case "1"
                    sJubSuGbn = "입원"
                Case "2"
                    sJubSuGbn = "수탁"
                Case Else
                    sJubSuGbn = ""
                End Select
                sPerResultT = sPerResultT & sJubSuGbn & "|"
            
                Select Case Trim(AdoRecord("EMERGBN") & "")
                Case "0"
                    sEmGbn = "(N)"
                Case "1"
                    sEmGbn = "(E)"
                Case Else
                    sEmGbn = ""
                End Select
                sPerResultT = sPerResultT & sEmGbn & "|"
            
                sDeptCd = Trim(AdoRecord("DEPTCD") & "")
                If sDeptCd <> "" Then
                    sPerResultT = sPerResultT & _
                        Trim(tPrtDept(Asc(UCase$(Left$(sDeptCd, 1))) * 26 + Asc(UCase$(Right$(sDeptCd, 1)))).sDeptNm) & "|"
                Else
                    sPerResultT = sPerResultT & "|"
                End If
            
                sPerResultT = sPerResultT & AdoRecord("WARD") & "|"
                sPerResultT = sPerResultT & AdoRecord("REGNO") & "|"
                sPerResultT = sPerResultT & AdoRecord("NAME") & "|"
            
                sSex = AdoRecord("SEXAGE") & ""
                If Left(sSex, 1) = "1" Or Left(sSex, 1) = "3" Then
                    sSex = "남" & Mid(sSex, 2)
                ElseIf Left(sSex, 1) = "2" Or Left(sSex, 1) = "4" Then
                    sSex = "여" & Mid(sSex, 2)
                Else
                    sSex = ""
                End If
                sPerResultT = sPerResultT & sSex & "|"
                
                sPerResultT = sPerResultT & AdoRecord("PRINTNM") & "|"
                sPerResultT = sPerResultT & AdoRecord("RESULT") & "|"
                PCnt = PCnt + 1
                sOldLabNo = sLabNo
                AdoRecord.MoveNext
            ElseIf (sOldLabNo <> sLabNo) And (sOldLabNo <> "") Then
                sResultT = sResultT & Trim(Str(PCnt)) & "|" & sPerResultT
                PCnt = 0
                sPerResultT = ""
                sOldLabNo = ""
                iCnt = iCnt - 1
            Else
                sPerResultT = sPerResultT & AdoRecord("PRINTNM") & "|"
                sPerResultT = sPerResultT & AdoRecord("RESULT") & "|"
                PCnt = PCnt + 1
                AdoRecord.MoveNext
            End If
        Next iCnt
        sResultT = sResultT & Trim(Str(PCnt)) & "|" & sPerResultT
        sResultT = Trim(Str(RCnt)) & "|" & sResultT
    End If
    AdoRecord.Close
    
    Get_ResultT = sResultT
   
Exit Function
ErrHandler:
    MsgBox "Get_ResultT(출력 DLL) - " & Err.Description
    AdoRecord.Close

End Function

' P9) 결과대장 출력 관리 MAIN부분
Function Print_ResultSheet(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String, sPrtKind As String) As Integer

    Dim iCnt        As Integer
    Dim iLabSeqN    As Integer

' 출력의 기초자료를 취득함 ("02" 결과대장)
    If Get_Print_Info("02", sSlipCd) = 0 Then
        Print_ResultSheet = -1
        Exit Function
    End If

' 출력해야 할 모든 레코드 갯수를 파악 sPartGbn은 정렬구분
    'iLabSeqN = Print_Work결과대장(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn, sPrtKind, 150)
    'KTY modify 19990305
    iLabSeqN = Print_Work결과대장(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn, sPrtKind, 190)
            
    If iLabSeqN > 0 Then
        Printer.EndDoc
        Print_ResultSheet = iLabSeqN
    Else
        Print_ResultSheet = 0
    End If

End Function

' P10) 실제 결과대장 출력
Function Print_Work결과대장(sLabDate As String, sSlipCd As String, sLabSeqS As String, sLabSeqE As String, sPrtGbn As String, sPrtKind As String, iLineLen As Integer) As Integer

    Dim iTotalCnt   As Integer
    Dim iOtotalcnt  As Integer
    Dim iRcnt       As Integer
    Dim iOCnt       As Integer
    Dim sResult     As String
    
    Dim sOldLabNo   As String
    Dim sOLabNo     As String
    Dim sOrdNm      As String
    
    Dim sLabNo      As String
    Dim sJubSuGbn   As String
    Dim sEmGbn      As String
    Dim sDeptCd     As String
    Dim sWard       As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sSexAge     As String
    Dim sPrintNm    As String
    Dim sResultT    As String
    
    Dim lResult    As String
    Dim jr         As Integer
    Dim kr         As Integer
    
    ' 슬립명을 취득
    '-- 2001.08.17(osw)
    If Trim(sSlipCd) = "All" Then
        sSlipNm = "전체항목"
    Else
        sSlipNm = Get_SlipNm(sSlipCd)
    End If
    
' SQL문 검사항목과 결과 취득
    sResultT = Get_ResultT(sLabDate, sSlipCd, sLabSeqS, sLabSeqE, sPrtGbn)
    
' LOOP을 하면서, 신상과 루틴이 있으면 출력하고, 없으면 낱개검사를 출력한다.
    If sResultT <> "" Then
        iTotalCnt = Val(GetByOne(sResultT, sResultT))   ' 전체 갯수
        'iPageLineMax = 80  ' 나누기를 해야 함으로 40Line
        'iColMax = 6
        'KTY modify 19990305
        iPageLineMax = 74  ' 나누기를 해야 함으로 40Line
        iColMax = 16
        iNoCnt = 1
        iNowRow = 1
        sOldLabNo = ""

        ' if sPrtGbn = "0" then ' 일반 결과대장출력
        ' if sPrtGbn = "1" then ' 혈액학 스타일의 결과대장출력
        For iRcnt = 1 To iTotalCnt
            iColNow = 0                                     ' 검체가 바뀔때 초기화
            iOtotalcnt = Val(GetByOne(sResultT, sResultT))  ' 검체당 검사갯수
            sLabNo = GetByOne(sResultT, sResultT)
            sJubSuGbn = GetByOne(sResultT, sResultT)
            sEmGbn = GetByOne(sResultT, sResultT)
            sDeptCd = GetByOne(sResultT, sResultT)
            sWard = GetByOne(sResultT, sResultT)
            sRegno = GetByOne(sResultT, sResultT)
            sName = GetByOne(sResultT, sResultT)
            sSexAge = GetByOne(sResultT, sResultT)
            
            sDate = Left(sLabNo, 8)
            
            'Printer.FontSize = 9
            'KTY 수정 19990305
            Printer.FontSize = 10
            
            If sOldLabNo <> sLabNo Then
                'Call Print_CheckPage_결과대장(iPageLineMax, iNowRow - 1, 150)
                'KTY modify 19990305
                Call Print_CheckPage_결과대장(iPageLineMax, iNowRow - 1, iLineLen)
                ' 인쇄
                
                'Printer.FontSize = 9
                'KTY 수정 19990305
                Printer.FontSize = 10
                Printer.Print Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));    ' No Seq
                iNoCnt = iNoCnt + 1
                Printer.Print Spc(1); Trim(sLabNo);                             ' 작업번호
                'Printer.Print Spc(1); Trim(sJubSuGbn);                          ' 접수구분
                'KTY modify 19990305
                Printer.Print Spc(1); Trim(sJubSuGbn); Spc(4 - LenH(Trim(sJubSuGbn)));   ' 접수구분
                If Trim(sEmGbn) <> "" Then                                      ' 응급구분
                    Printer.Print Trim(sEmGbn); Spc(4 - LenH(Trim(sEmGbn)));
                Else
                    Printer.Print Spc(4);
                End If
                
                'Printer.Print Spc(1); Trim(sDeptCd); "/"; Trim(sWard); Spc(30 - (LenH(sDeptCd) + LenH(sWard)));
                'Printer.Print Spc(1); Trim(sRegno); Spc(15 - Len(Trim(sRegno)));
                'Printer.Print Spc(1); Trim(sName); Spc(15 - LenH(Trim(sName)));
                'KTY 수정 19990305
                
                If LenH(sDeptCd) > 4 Then
                    sDeptCd = LeftH(sDeptCd, 4)
                End If
                If LenH(sWard) > 4 Then
                    sWard = LeftH(sWard, 4)
                End If
                If LenH(sDeptCd) + LenH(sWard) > 7 Then
                    sWard = LeftH(sWard, (7 - 1) - LenH(sDeptCd))
                End If
                
                Printer.Print Spc(1); Trim(sDeptCd); "/"; Trim(sWard); Spc((7 - 1) - (LenH(sDeptCd) + LenH(sWard)));
                
                Printer.Print Spc(1); Trim(sRegno); Spc(8 - Len(Trim(sRegno)));
                
                If LenH(sName) > 8 Then
                    sName = LeftH(sName, 8)
                End If
                Printer.Print Spc(1); Trim(sName); Spc(8 - LenH(Trim(sName)));
                
                If sSexAge <> "" Then
                    Printer.Print Spc(1); Trim(sSexAge); Spc(9 - LenH(Trim(sSexAge)));
                Else
                    Printer.Print Spc(10);
                End If
                
                For iOCnt = 1 To iOtotalcnt
                    iTotalCnt = iTotalCnt + 1
                    sPrintNm = GetByOne(sResultT, sResultT)
                    sResult = GetByOne(sResultT, sResultT)
                        
                    ' 낱개 검사항목 출력
                        If iColNow = iColMax Then
                            iColNow = 0
                            ' Line위치를 체크 (NewPage를 검사)
                            Printer.Print
                            Call Print_CheckPage_결과대장(iPageLineMax, iNowRow - 1, iLineLen)
                            'Printer.FontSize = 9
                            'KTY modify 19990305
                            Printer.FontSize = 10
                            '다음줄처리
                            'Printer.Print Spc(103); Spc(1); LeftH(Trim(sPrintNm), 10); Spc(10 - LenH(Trim(Left(sPrintNm, 10))));
                            'Printer.Print Spc(1); LeftH(Trim(sResult), 10); Spc(10 - LenH(Trim(Left(sResult, 10))));
                            'KTY modify 19990305
                            
                            Printer.Print Spc(63); Spc(1); LeftH(Trim(sPrintNm), 8); Spc(8 - LenH(Trim(Left(sPrintNm, 8))));
                            'Printer.Print Spc(1); LeftH(Trim(sResult), 6); Spc(6 - LenH(Trim(Left(sResult, 6))));
                        
                            If Len(Trim(sResult)) <= 6 Then
                                Printer.Print Spc(1); LeftH(Trim(sResult), 6); Spc(6 - LenH(Trim(Left(sResult, 6))));
                            Else
                                kr = 0
                                For jr = 2 To Len(Trim(sResult))
                                    If InStr(Left(Trim(sResult), jr), " ") <> 0 Then
                                        If kr <> 0 Then
                                           lResult = Left(Trim(sResult), kr - 1) & Mid(sResult, jr + 1, 6)
                                        Else
                                           kr = jr
                                           lResult = Left(Trim(sResult), jr - 1) & Mid(sResult, jr + 1, 6)
                                        End If
                                        If Len(Trim(lResult)) <= 6 Then Exit For
                                    End If
                                    
                                Next
                                
                                Printer.Print Spc(1); LeftH(Trim(lResult), 6); Spc(6 - LenH(Trim(Left(lResult, 6))));
                                lResult = ""
                            End If
                        
                        
                        Else
                            'Printer.Print Spc(1); LeftH(Trim(sPrintNm), 10); Spc(10 - LenH(Trim(Left(sPrintNm, 10))));
                            'Printer.Print Spc(1); LeftH(Trim(sResult), 10); Spc(10 - LenH(Trim(Left(sResult, 10))));
                            'KTY modify 19990305
'                            If sPrintNm = "LEU" Then Stop
                            Printer.Print Spc(1); LeftH(Trim(sPrintNm), 8); Spc(8 - LenH(Trim(Left(sPrintNm, 8))));
                            
                            If Len(Trim(sResult)) <= 6 Then
                                Printer.Print Spc(1); LeftH(Trim(sResult), 6); Spc(6 - LenH(Trim(Left(sResult, 6))));
                            Else
                                kr = 0
                                For jr = 2 To Len(Trim(sResult))
                                    If InStr(Left(Trim(sResult), jr), " ") <> 0 Then
                                        If kr <> 0 Then
                                           lResult = Left(Trim(sResult), kr - 1) & Mid(sResult, jr + 1, 6)
                                        Else
                                           kr = jr
                                           lResult = Left(Trim(sResult), jr - 1) & Mid(sResult, jr + 1, 6)
                                        End If
                                        If Len(Trim(lResult)) <= 6 Then Exit For
                                    End If
                                    
                                Next
                                
                                Printer.Print Spc(1); LeftH(Trim(lResult), 6); Spc(6 - LenH(Trim(Left(lResult, 6))));
                                lResult = ""
                            End If
                        End If
                        iColNow = iColNow + 2
                Next iOCnt
                
                Printer.Print
                'KTY Add 19990305
                Printer.FontSize = 10
                Printer.Print String(iLineLen, Chr(45))
            End If
        Next iRcnt
        Call Print_Tail(iLineLen)
    End If
    Print_Work결과대장 = iNoCnt - 1

End Function

' P6) 한 페이지 당 HEAD출력(결과대장)
Sub Print_Head_결과대장(iLineLen As Integer)

    Dim sLabNo  As String
    Dim iCnt    As Integer

    Dim P       As Object

    For Each P In Printers
        If P.DeviceName Like "*대장*" Then
            Set Printer = P
            Exit For
        End If
    Next P

    sLabNo = sDate
    sLabNo = Left(sLabNo, 4) & "-" & Mid(sLabNo, 5, 2) & "-" & Mid(sLabNo, 7, 2)
    
    Printer.FontName = "굴림체"
' 메인 TITLE출력
    Printer.FontSize = 16
    'Printer.Print Spc(40); "SLIP별 결과대장"
    'KTY 수정 19990305
    Printer.Print Spc(50); "SLIP별 결과대장"
' 상위 HEAD
    Printer.FontSize = 11
    'Printer.Print "접수일자 : "; sLabNo; Spc(85 - Len(Trim(sLabNo))); "출력일자 : "; Format(Now, "YYYY-MM-DD")
    'Printer.Print "SLIP  명 : "; sSlipNm; Spc(85 - LenH(Trim(sSlipNm))); "PAGE     : "; Printer.Page
    'KTY 수정 19990305
    Printer.Print "접수일자 : "; sLabNo; Spc(135 - Len(Trim(sLabNo))); "출력일자 : "; Format(Now, "YYYY-MM-DD")
    Printer.Print "SLIP  명 : "; sSlipNm; Spc(135 - LenH(Trim(sSlipNm))); "PAGE     : "; Printer.Page
' LINE 출력
    Printer.FontSize = 10
    Printer.Print String(iLineLen, Chr(6))
' 하위 HEAD
    'Printer.FontSize = 9
    'KTY 수정 19990305
    Printer.FontSize = 10
    'Printer.Print "No."; Spc(1); "작업번호"; Spc(7); "접수구분"; Spc(3); "과/병동"; Spc(25); "등록번호"; Spc(8); "이름"; Spc(12); "성별/나이"; Spc(1);
    'KTY 수정 19990305
    Printer.Print "No."; Spc(1); "작업번호"; Spc(7); "접수구분"; Spc(1); "과/병동"; Spc(1); "등록번호"; Spc(1); "이름"; Spc(5); "성별/나이"; Spc(1);
    For iCnt = 1 To iColMax / 2
        'Printer.Print "검사항목"; Spc(3); "결과"; Spc(7);
        'KTY 수정
        Printer.Print "검사항목"; Spc(1); "결과"; Spc(3);
    Next iCnt
    Printer.Print
' LINE 출력
    Printer.FontSize = 10
    Printer.Print String(iLineLen, Chr(6))

End Sub

' P12) 결과대장 PageSkip Check
Sub Print_CheckPage_결과대장(iPageLineMax As Integer, iLNowRow As Integer, iLineLen As Integer)
    
    If (iLNowRow <> 0) And (iLNowRow Mod (iPageLineMax / 2) = 0) Then
        Call Print_Tail(iLineLen)       ' TAIL인쇄후 HEAD인쇄
        Printer.NewPage
        
        '-- 2001.08.14(osw)
        Printer.PaperSize = 12
        Printer.Orientation = 2
        
        'Call Print_Head_결과대장(150)
        'KTY 수정 19990305
        Call Print_Head_결과대장(iLineLen)
    ElseIf iLNowRow = 0 Then
        'Call Print_Head_결과대장(150)   ' HEAD만 인쇄
        'KTY 수정 19990305
        
        '-- 2001.08.14(osw)
        Printer.PaperSize = 12
        Printer.Orientation = 2
        
        Call Print_Head_결과대장(iLineLen)   ' HEAD만 인쇄
    End If
    iNowRow = iLNowRow + 2              ' 현재 Line번호 증가

End Sub

