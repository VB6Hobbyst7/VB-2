VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCT0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' 화면 표시
' S1) Function Get_ResultCnt(sPartCd As String, sSearchUnit As String, sLabStart As String, sLabEnd As String, sPartGbn As String, sSearchGbn As String) As String
' 출력
' S2) Function Get_Print_Info(PrtKind As String) As Integer
' P1) Function Print_ResultCnt(sPartCd As String, sSearchUnit As String, sLabStart As String, sLabEnd As String, sPartGbn As String, sSearchGbn As String) As Integer
' P2) Sub Print_Head_검사건수통계(iLineLen As Integer)
' P3) Sub Print_Tail(iLineLen As Integer)
' P4) Sub Print_CheckPage_검사건수통계(iPageLineMax As Integer, iLNowRow As Integer, iLineLen As Integer)

Type sDayC
    sDayTitle       As String
    sDayCnt         As String
    sDaySum         As String
End Type

Type sMonthC
    sMonthTitle     As String
    sMonthCnt       As String
    sMonthSum       As String
End Type

Type sYearC
    sYearTitle      As String
    sYearCnt        As String
    sYearSum        As String
End Type

' 출력 전체
Dim sFontNm             As String       ' 출력시 FontName
Dim sDirectory          As String       ' 출력시 인쇄방향
Dim sFontSize           As String       ' 출력시 FontSize

' 년월일 검사건수 통계
Dim iTotalCnt           As Integer      ' 출력건수
Dim iPageLineMax        As Integer      ' 출력시 한 페이지의 최대 라인
Dim iNoCnt              As Integer      ' 출력시 Seq번호를 처리하는 변수
Dim iNowRow             As Integer      ' 현재 출력 Line수
Dim iColMax             As Integer      ' ITEM건수 인쇄시 최대 갯수
Dim iColNow             As Integer      ' 검사항목 현재 인쇄 Col위치
Dim iMaxCnt             As Integer      ' 총 인쇄해야 할 날짜 갯수
Dim sTitleFont          As String       ' TITLE의 FONT SIZE
Dim sPartNm             As String       ' Part명
Dim sSlipNm             As String       ' Slip명
Dim sDate               As String       ' 검색구간
Dim sAddTitle           As String       ' TITLE에 추가 구문

Dim sDayCnt(31)         As sDayC        ' 일자별 조회시 구조체
Dim sMonthCnt(12)       As sMonthC      ' 월별 조회시 구조체
Dim sYearCnt(10)        As sYearC       ' 년별 조회시 구조체
Dim P                   As Object

Private Const ciHeaderFntSize = 10

' S1) 일별검사항목 통계처리
Public Function Get_ResultCnt(ByVal sPartCd As String, ByVal sSearchUnit As String, ByVal sLabStart As String, ByVal sLabEnd As String, ByVal sPartGbn As String, ByVal sSearchGbn As String) As String
    
    Dim AdoRecord   As New ADODB.Recordset
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    Dim sSql            As String
    Dim sLabDate        As String
    Dim sSLabTime       As String
    Dim sELabTime       As String
    Dim sResultCnt      As String
    Dim sTestNmCfg      As String
    Dim sSearchGbn2     As String
    Dim sTmp            As String
    Dim iJUBSU          As Integer
    
    sResultCnt = ""
    sTestNmCfg = "P"
    
    If Len(Trim(sLabStart)) > 8 Then
        sSLabTime = Mid(sLabStart, 9, 2) & "00"
        sLabStart = Left(sLabStart, 8)
    Else
        sLabStart = Left(sLabStart, 8)
    End If
    
    If Len(Trim(sLabEnd)) > 8 Then
        sELabTime = Mid(sLabEnd, 9, 2) & "59"
        sLabEnd = Left(sLabEnd, 8)
    Else
        sLabEnd = Left(sLabEnd, 8)
    End If
    
    Select Case sSearchUnit
    Case "D"
        sLabDate = "RESULT.LABDATE"
    Case "M"
        sLabDate = "LEFT(RESULT.LABDATE,6)"
        sLabStart = Left(sLabStart, 6) & "01"
        sLabEnd = Left(sLabEnd, 6) & "31"
    Case "Y"
        sLabDate = "LEFT(RESULT.LABDATE,4)"
        sLabStart = Left(sLabStart, 4) & "0101"
        sLabEnd = Left(sLabEnd, 4) & "1231"
    End Select
    
    sTmp = GetByOne(sSearchGbn, sSearchGbn)
    sSearchGbn2 = GetByOne(sSearchGbn, sSearchGbn)
    sSearchGbn = sTmp
    
    iJUBSU = 0
    
    On Error GoTo ErrHandler
' 전체 통계
    If sSearchGbn = "N" Then
        If sTestNmCfg = "T" Then
            sSql = " SELECT " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.TESTITEMNM, COUNT(*) AS RCOUNT " & _
                   "   FROM " & sPartCd & "_RESULT AS RESULT, TESTITEM "
        ElseIf sTestNmCfg = "P" Then
            sSql = " SELECT " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.PRINTNM, COUNT(*) AS RCOUNT " & _
                   "   FROM " & sPartCd & "_RESULT AS RESULT, TESTITEM "
        End If
        
        If sSLabTime <> "" And sELabTime <> "" Then
            sSql = sSql & ", " & sPartCd & "_JUBSU AS JUBSU"
            
            iJUBSU = 1
        Else
            If sSearchGbn2 <> "T" Then
                sSql = sSql & ", " & sPartCd & "_JUBSU AS JUBSU"
                
                iJUBSU = 1
            End If
        End If
    
        sSql = sSql & "  WHERE RESULT.LABDATE BETWEEN '" & sLabStart & "'  AND '" & sLabEnd & "' "
    
        If sPartGbn <> "" Then
            sSql = sSql & " AND RESULT.PARTGBN   = '" & sPartGbn & "' "
        End If
        
        '전체
        If sSearchGbn2 = "T" Then
        '외래
        ElseIf sSearchGbn2 = "O" Then
            sSql = sSql & " AND (JUBSU.JUBSUGBN = '0' OR JUBSU.JUBSUGBN = '' OR JUBSU.JUBSUGBN IS NULL) "
        '입원
        ElseIf sSearchGbn2 = "I" Then
            sSql = sSql & " AND JUBSU.JUBSUGBN = '1' "
        '수탁
        ElseIf sSearchGbn2 = "S" Then
            sSql = sSql & " AND JUBSU.JUBSUGBN = '2' "
        End If
        
        If sSLabTime <> "" And sELabTime <> "" Then
            sSql = sSql & " AND JUBSU.LABTIME BETWEEN '" & sSLabTime & "'  AND '" & sELabTime & "' "
        End If
        
        If iJUBSU = 1 Then
            sSql = sSql & " AND JUBSU.LABDATE        = RESULT.LABDATE" & _
                          " AND JUBSU.PARTGBN        = RESULT.PARTGBN " & _
                          " AND JUBSU.LABSEQ         = RESULT.LABSEQ "
        End If
        
        sSql = sSql & " AND TESTITEM.PARTCD      = '" & sPartCd & "' " & _
                      " AND TESTITEM.PARTGBN     = RESULT.PARTGBN " & _
                      " AND TESTITEM.SPECIMENCD  = RESULT.SPECIMENCD " & _
                      " AND TESTITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ " & _
                      " AND TESTITEM.SUBMCD      = RESULT.SUBMCD " & _
                      " AND RESULT.RESULT <> '' " & _
                      " AND RESULT.RESULT IS NOT NULL "

        If sTestNmCfg = "T" Then
            sSql = sSql & "GROUP BY " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.TESTITEMNM "
        ElseIf sTestNmCfg = "P" Then
            sSql = sSql & "GROUP BY " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.PRINTNM "
        End If
        
        sSql = sSql & "ORDER BY RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD "
               
    ElseIf sSearchGbn = "A" Then
        If sTestNmCfg = "T" Then
            sSql = " SELECT " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.TESTITEMNM, COUNT(*) AS RCOUNT " & _
                   "   FROM " & sPartCd & "_ABNORMAL AS RESULT, TESTITEM "
        ElseIf sTestNmCfg = "P" Then
            sSql = " SELECT " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.PRINTNM, COUNT(*) AS RCOUNT " & _
                   "   FROM " & sPartCd & "_ABNORMAL AS RESULT, TESTITEM "
        End If
        
        If sSLabTime <> "" And sELabTime <> "" Then
            sSql = sSql & ", " & sPartCd & "_JUBSU AS JUBSU"
            
            iJUBSU = 1
        Else
            If sSearchGbn2 <> "T" Then
                sSql = sSql & ", " & sPartCd & "_JUBSU AS JUBSU"
                
                iJUBSU = 1
            End If
        End If
        
        sSql = sSql & "  WHERE RESULT.LABDATE BETWEEN '" & sLabStart & "'  AND '" & sLabEnd & "' "
    
        If sPartGbn <> "" Then
            sSql = sSql & " AND RESULT.PARTGBN   = '" & sPartGbn & "' "
        End If
    
        '전체
        If sSearchGbn2 = "T" Then
        '외래
        ElseIf sSearchGbn2 = "O" Then
            sSql = sSql & " AND (JUBSU.JUBSUGBN = '0' OR JUBSU.JUBSUGBN = '' OR JUBSU.JUBSUGBN IS NULL) "
        '입원
        ElseIf sSearchGbn2 = "I" Then
            sSql = sSql & " AND JUBSU.JUBSUGBN = '1' "
        '수탁
        ElseIf sSearchGbn2 = "S" Then
            sSql = sSql & " AND JUBSU.JUBSUGBN = '2' "
        End If
        
        If sSLabTime <> "" And sELabTime <> "" Then
            sSql = sSql & " AND JUBSU.LABTIME BETWEEN '" & sSLabTime & "'  AND '" & sELabTime & "' "
        End If
        
        If iJUBSU = 1 Then
            sSql = sSql & " AND JUBSU.LABDATE        = RESULT.LABDATE" & _
                          " AND JUBSU.PARTGBN        = RESULT.PARTGBN " & _
                          " AND JUBSU.LABSEQ         = RESULT.LABSEQ "
        End If
        
        sSql = sSql & " AND TESTITEM.PARTCD      = '" & sPartCd & "' " & _
                      " AND TESTITEM.PARTGBN     = RESULT.PARTGBN " & _
                      " AND TESTITEM.SPECIMENCD  = RESULT.SPECIMENCD " & _
                      " AND TESTITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ " & _
                      " AND TESTITEM.SUBMCD      = RESULT.SUBMCD " & _
                      " AND RESULT.RESULT <> '' " & _
                      " AND RESULT.RESULT IS NOT NULL "
        
        If sTestNmCfg = "T" Then
            sSql = sSql & "GROUP BY " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.TESTITEMNM "
        ElseIf sTestNmCfg = "P" Then
            sSql = sSql & "GROUP BY " & sLabDate & ", RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD, TESTITEM.PRINTNM "
        End If
        
        sSql = sSql & "ORDER BY RESULT.PARTGBN, RESULT.SPECIMENCD, RESULT.TESTITEMSEQ, RESULT.SUBMCD "
        
    End If
' 이상자 통계

    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet {검사항목약명 | Delta값 | Delta 구분 |}n
        RCnt = AdoRecord.RecordCount
        sResultCnt = AdoRecord.GetString(adClipString, RCnt, "|", "|", "0")
    Else
        sResultCnt = "0|"
    End If
    
    sResultCnt = Trim(Str(RCnt)) & "|" & sResultCnt
    
    AdoRecord.Close
    
    Get_ResultCnt = sResultCnt

Exit Function


ErrHandler:
    MsgBox "GetItem(통계 DLL) - " & Err.Description
    
End Function

' P1) 통계 출력
Public Function Print_ResultCnt(ByVal sPartCd As String, ByVal sSearchUnit As String, ByVal sLabStart As String, ByVal sLabEnd As String, ByVal sPartGbn As String, ByVal sSearchGbn As String, ByVal sTotList As String, ByVal sSum As String) As Integer

    Dim dsLabStart  As Date
    Dim dsLabEnd    As Date
    
    Dim iPrtCnt     As Integer
    
    Dim RCnt        As Integer
    
    Dim sPrtDate    As Long
    Dim iCnt        As Integer
    Dim iLabSeqN    As Integer
    Dim sResultCnt  As String
    
    Dim sdummy      As String
    Dim sPrintNm    As String
    Dim sOldPrintNm As String
    Dim sRCount     As String
    
    Dim sSearchGbn2 As String
    Dim sTmp        As String
    Dim DCJ0101 As DCJ0101
    
    '출력의 기초자료를 취득함 ("21" 년월일검사건수통계)
    If Get_Print_Info("21") = 0 Then
        Print_ResultCnt = -1
        Exit Function
    End If
        
    'Part명 취득
    For iCnt = 1 To giPartCnt
        If gPartTable(iCnt).sPartInit = sPartCd Then
            sSlipNm = gPartTable(iCnt).sPartName
            Exit For
        End If
    Next
    
    '슬립명을 취득
    If Trim(sPartGbn) <> "" Then
        Set DCJ0101 = New DCJ0101
        sSlipNm = sSlipNm & "(" & Trim(DCJ0101.Get_SlipNm(sPartCd + sPartGbn)) & ")"
        Set DCJ0101 = Nothing
    End If
    
    ' 추가 타이틀 구문
    Select Case sSearchUnit
    Case "D"
        sAddTitle = "일"
        ' 조회구간
        If Left(sLabStart, 8) = Left(sLabEnd, 8) Then
            sDate = Left(sLabStart, 4) & "/" & Mid(sLabStart, 5, 2) & "/" & Mid(sLabStart, 7, 2)
        Else
            sDate = Left(sLabStart, 4) & "/" & Mid(sLabStart, 5, 2) & "/" & Mid(sLabStart, 7, 2) & _
                    " ~ " & Left(sLabEnd, 4) & "/" & Mid(sLabEnd, 5, 2) & "/" & Mid(sLabEnd, 7, 2)
        End If
        
        If Len(sLabStart) > 8 Then
            If Left(sLabStart, 8) = Left(sLabEnd, 8) Then
                sDate = Left(sLabStart, 4) & "/" & Mid(sLabStart, 5, 2) & "/" & Mid(sLabStart, 7, 2) & "  " & _
                        Mid(sLabStart, 9, 2) & ":" & "00" & _
                        "~" & Mid(sLabEnd, 9, 2) & ":" & "59"
            Else
                sDate = Left(sLabStart, 4) & "/" & Mid(sLabStart, 5, 2) & "/" & Mid(sLabStart, 7, 2) & _
                        " ~ " & Left(sLabEnd, 4) & "/" & Mid(sLabEnd, 5, 2) & "/" & Mid(sLabEnd, 7, 2) & "  " & _
                        Mid(sLabStart, 9, 2) & ":" & "00" & _
                        "~" & Mid(sLabEnd, 9, 2) & ":" & "59"
            End If
        End If
    Case "M"
        sAddTitle = "월"
        ' 조회구간
        If Left(sLabStart, 6) = Left(sLabEnd, 6) Then
            sDate = Left(sLabStart, 4) & "/" & Mid(sLabStart, 5, 2)
        Else
            sDate = Left(sLabStart, 4) & "/" & Mid(sLabStart, 5, 2) & " ~ " & Left(sLabEnd, 4) & "/" & Mid(sLabEnd, 5, 2)
        End If
    Case "Y"
        sAddTitle = "년"
        ' 조회구간
        If Left(sLabStart, 4) = Left(sLabEnd, 4) Then
            sDate = Left(sLabStart, 4)
        Else
            sDate = Left(sLabStart, 4) & " ~ " & Left(sLabEnd, 4)
        End If
    End Select
    
    sTmp = GetByOne(sSearchGbn, sSearchGbn)
    sSearchGbn2 = GetByOne(sSearchGbn, sSearchGbn)
    sSearchGbn = sTmp
    
    If sSearchGbn = "N" Then
        sAddTitle = sAddTitle & "/모두"
    ElseIf sSearchGbn = "A" Then
        sAddTitle = sAddTitle & "/이상자"
    End If
    
    If sSearchGbn2 = "T" Then
        sAddTitle = sAddTitle & " - 전체"
    ElseIf sSearchGbn2 = "O" Then
        sAddTitle = sAddTitle & " - 외래"
    ElseIf sSearchGbn2 = "I" Then
        sAddTitle = sAddTitle & " - 입원"
    ElseIf sSearchGbn2 = "S" Then
        sAddTitle = sAddTitle & " - 수탁"
    End If
    
    iPrtCnt = 0
    
    '출력해야 할 모든 레코드 갯수를 파악 sPartGbn은 정렬구분
    sResultCnt = sTotList
    
    If sResultCnt = "" Then
        Print_ResultCnt = 0
        Exit Function
    End If
    
    '출력할 Printer 설정
    For Each P In Printers
        If P.DeviceName Like "*통계*" Then
            Set Printer = P
        End If
    Next
    
    '출력
    Select Case sSearchUnit
    Case "D"
        dsLabStart = CDate(Left(sLabStart, 4) & "-" & Mid(sLabStart, 5, 2) & "-" & Mid(sLabStart, 7, 2))
        dsLabEnd = CDate(Left(sLabEnd, 4) & "-" & Mid(sLabEnd, 5, 2) & "-" & Mid(sLabEnd, 7, 2))
        
        For iCnt = 0 To 30
            If Format(dsLabEnd, "YYYYMMDD") < Format(DateAdd("d", iCnt, dsLabStart), "YYYYMMDD") Then
                ' 출력숨김
                sDayCnt(iCnt).sDayTitle = ""
            Else
                '출력
                sDayCnt(iCnt).sDayTitle = Replace(CStr(DateAdd("d", iCnt, dsLabStart)), "-", "/")
                iPrtCnt = iPrtCnt + 1
            End If
        Next iCnt
        
        If sLabStart = sLabEnd Then
            iPrtCnt = 1
            sDayCnt(0).sDayTitle = Replace(CStr(dsLabStart), "-", "/")
        End If
        
        iMaxCnt = iPrtCnt
        
        iTotalCnt = Val(GetByOne(sResultCnt, sResultCnt))   ' 전체 갯수
        iNoCnt = 1
        iNowRow = 1
        sOldPrintNm = ""
        
        '메모리 CLEAR
        For RCnt = 0 To iMaxCnt - 1
            sDayCnt(RCnt).sDayCnt = "0"
            sDayCnt(RCnt).sDaySum = GetByOne(sSum, sSum)
        Next RCnt
        sDayCnt(31).sDayCnt = "0"
        sDayCnt(31).sDaySum = GetByOne(sSum, sSum)
        
        ' 실제 출력
        For iCnt = 1 To iTotalCnt
            iColNow = 0
            sPrtDate = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sPrintNm = GetByOne(sResultCnt, sResultCnt)
            sRCount = GetByOne(sResultCnt, sResultCnt)

            '검사항목이 바뀜
            If sOldPrintNm <> sPrintNm And Trim(sOldPrintNm) <> "" Then
                '헤더 또는 라인번호 증가 또는 테일 출력
                Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                
                '검사항목 및 건수 출력
                Printer.FontSize = sFontSize
                Printer.Print Spc(11); Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));   ' No Seq
                iNoCnt = iNoCnt + 1
                Printer.Print Spc(1); LeftH(sOldPrintNm, 10); Spc(10 - LenH(Trim(sOldPrintNm))); Spc(4); ' 일통계만 다름
                
                For RCnt = 0 To iMaxCnt - 1
                    If (RCnt Mod iColMax = 0) And (RCnt > 0) Then
                        Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                        Printer.FontSize = sFontSize
                        '이전 라인 내용 찍고
                        Printer.Print
                                                
                        '새 라인 내용 구성
                        Printer.Print Spc(11); Spc(18);
                    End If
                    sDayCnt(RCnt).sDayCnt = Format$(Val(sDayCnt(RCnt).sDayCnt), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sDayCnt(RCnt).sDayCnt))); sDayCnt(RCnt).sDayCnt;
                Next RCnt
                
                '소계 출력
                sDayCnt(31).sDayCnt = Format$(Val(sDayCnt(31).sDayCnt), "##,###,##0")
                
                '라인을 꽉 채웠을 때
                If iMaxCnt Mod (iColMax) = 0 Then
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sDayCnt(31).sDayCnt))); sDayCnt(31).sDayCnt
                '라인을 덜 채웠을 때
                Else
                    Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sDayCnt(31).sDayCnt))); sDayCnt(31).sDayCnt
                End If

                '라인 출력
                Printer.Print Spc(11); String(18 + (11 * (iColMax + 1)), Chr(45))
                
                '새로운 검사항목에 대한 작업
                '메모리 CLEAR
                For RCnt = 0 To iMaxCnt
                    sDayCnt(RCnt).sDayCnt = "0"
                Next RCnt
                sDayCnt(31).sDayCnt = "0"
            End If
                
            For RCnt = 0 To iMaxCnt - 1
                If Left(sPrtDate, 4) & "/" & Mid(sPrtDate, 5, 2) & "/" & Mid(sPrtDate, 7, 2) = Left(sDayCnt(RCnt).sDayTitle, 10) Then
                    sDayCnt(31).sDayCnt = Trim(Format$(Str(Val(sDayCnt(31).sDayCnt) + Val(sRCount)), "##,###,##0"))
                    sDayCnt(RCnt).sDayCnt = Format$(Val(sRCount), "##,###,##0")
                    Exit For
                End If
            Next RCnt
            
            sOldPrintNm = sPrintNm
        Next iCnt
        
' 마지막 레코드 인쇄
        '라인 번호 증가
        Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
        
        '마지막 검사항목 및 건수 출력
        Printer.FontSize = sFontSize
        Printer.Print Spc(11); Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));   ' No Seq
        iNoCnt = iNoCnt + 1
        Printer.Print Spc(1); LeftH(sOldPrintNm, 10); Spc(10 - LenH(Trim(sOldPrintNm))); Spc(4); ' 일통계만 다름
        
        For RCnt = 0 To iMaxCnt - 1
            If (RCnt Mod iColMax = 0) And (RCnt > 0) Then
                Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                Printer.FontSize = sFontSize
                '이전 라인 내용 찍고
                Printer.Print
                                        
                '새 라인 내용 구성
                Printer.Print Spc(11); Spc(18);
            End If
            sDayCnt(RCnt).sDayCnt = Format$(Val(sDayCnt(RCnt).sDayCnt), "##,###,##0")
            Printer.Print Spc(1); Spc(10 - Len(Trim(sDayCnt(RCnt).sDayCnt))); sDayCnt(RCnt).sDayCnt;
        Next RCnt
        
        '소계 출력
        sDayCnt(31).sDayCnt = Format$(Val(sDayCnt(31).sDayCnt), "##,###,##0")
        
        '라인을 꽉 채웠을 때
        If iMaxCnt Mod (iColMax) = 0 Then
            Printer.Print Spc(1); Spc(10 - Len(Trim(sDayCnt(31).sDayCnt))); sDayCnt(31).sDayCnt
        '라인을 덜 채웠을 때
        Else
            Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sDayCnt(31).sDayCnt))); sDayCnt(31).sDayCnt
        End If
        
        '라인 출력
        Printer.Print Spc(11); String(18 + (11 * (iColMax + 1)), Chr(45))
        
        Call Print_Tail(18 + (11 * (iColMax + 1)))
    
    Case "M"
        dsLabStart = CDate(Left(sLabStart, 4) & "-" & Mid(sLabStart, 5, 2) & "-" & Mid(sLabStart, 7, 2))
        dsLabEnd = CDate(Left(sLabEnd, 4) & "-" & Mid(sLabEnd, 5, 2) & "-28")
        
        For iCnt = 0 To 11
            If Format(dsLabEnd, "YYYYMM") < Format(DateAdd("m", iCnt, dsLabStart), "YYYYMM") Then
                ' 출력숨김
                sMonthCnt(iCnt).sMonthTitle = ""
            Else
                '출력
                sMonthCnt(iCnt).sMonthTitle = Replace(Left(CStr(DateAdd("m", iCnt, dsLabStart)), 7), "-", "/")
                iPrtCnt = iPrtCnt + 1
            End If
        Next iCnt
        
        If Left(sLabStart, 6) = Left(sLabEnd, 6) Then
            iPrtCnt = 1
            sMonthCnt(0).sMonthTitle = Replace(Left(CStr(dsLabStart), 7), "-", "/")
        End If
        
        iMaxCnt = iPrtCnt
        
        iTotalCnt = Val(GetByOne(sResultCnt, sResultCnt))   ' 전체 갯수
        iNoCnt = 1
        iNowRow = 1
        sOldPrintNm = ""
        
        '메모리 CLEAR
        For RCnt = 0 To iMaxCnt - 1
            sMonthCnt(RCnt).sMonthCnt = "0"
            sMonthCnt(RCnt).sMonthSum = GetByOne(sSum, sSum)
        Next RCnt
        sMonthCnt(12).sMonthCnt = "0"
        sMonthCnt(12).sMonthSum = GetByOne(sSum, sSum)
        
        ' 실제 출력
        For iCnt = 1 To iTotalCnt
            iColNow = 0
            sPrtDate = Left(GetByOne(sResultCnt, sResultCnt), 6)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sPrintNm = GetByOne(sResultCnt, sResultCnt)
            sRCount = GetByOne(sResultCnt, sResultCnt)
            
            '검사항목이 바뀜
            If sOldPrintNm <> sPrintNm And Trim(sOldPrintNm) <> "" Then
                '헤더 또는 라인번호 증가 또는 테일 출력
                Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                
                '검사항목 및 건수 출력
                Printer.FontSize = sFontSize
                Printer.Print Spc(11); Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));   ' No Seq
                iNoCnt = iNoCnt + 1
                Printer.Print Spc(1); LeftH(sOldPrintNm, 10); Spc(10 - LenH(Trim(sOldPrintNm))); Spc(4);
                
                For RCnt = 0 To iMaxCnt - 1
                    If (RCnt Mod iColMax = 0) And (RCnt > 0) Then
                        Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                        Printer.FontSize = sFontSize
                        '이전 라인 내용 찍고
                        Printer.Print
                                                
                        '새 라인 내용 구성
                        Printer.Print Spc(11); Spc(18);
                    End If
                    sMonthCnt(RCnt).sMonthCnt = Format$(Val(sMonthCnt(RCnt).sMonthCnt), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sMonthCnt(RCnt).sMonthCnt))); sMonthCnt(RCnt).sMonthCnt;
                Next RCnt
                
                '소계 출력
                sMonthCnt(12).sMonthCnt = Format$(Val(sMonthCnt(12).sMonthCnt), "##,###,##0")
                
                '라인을 꽉 채웠을 때
                If iMaxCnt Mod (iColMax) = 0 Then
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sMonthCnt(12).sMonthCnt))); sMonthCnt(12).sMonthCnt
                '라인을 덜 채웠을 때
                Else
                    Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sMonthCnt(12).sMonthCnt))); sMonthCnt(12).sMonthCnt
                End If
                
                '라인 출력
                Printer.Print Spc(11); String(18 + (11 * (iColMax + 1)), Chr(45))
                
                '새로운 검사항목에 대한 작업
                '메모리 CLEAR
                For RCnt = 0 To iMaxCnt
                    sMonthCnt(RCnt).sMonthCnt = "0"
                Next RCnt
                sMonthCnt(12).sMonthCnt = "0"
            End If
                
            For RCnt = 0 To iMaxCnt - 1
                If Left(sPrtDate, 4) & "/" & Mid(sPrtDate, 5, 2) = Left(sMonthCnt(RCnt).sMonthTitle, 7) Then
                    sMonthCnt(12).sMonthCnt = Trim(Format$(Str(Val(sMonthCnt(12).sMonthCnt) + Val(sRCount)), "##,###,##0"))
                    sMonthCnt(RCnt).sMonthCnt = Format$(Val(sRCount), "##,###,##0")
                    Exit For
                End If
            Next RCnt
            
            sOldPrintNm = sPrintNm
        Next iCnt
        
' 마지막 레코드 인쇄
        '라인 번호 증가
        Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
        
        '마지막 검사항목 및 건수 출력
        Printer.FontSize = sFontSize
        Printer.Print Spc(11); Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));   ' No Seq
        iNoCnt = iNoCnt + 1
        Printer.Print Spc(1); LeftH(sOldPrintNm, 10); Spc(10 - LenH(Trim(sOldPrintNm))); Spc(4); ' 일통계만 다름
        
        For RCnt = 0 To iMaxCnt - 1
            If (RCnt Mod iColMax = 0) And (RCnt > 0) Then
                Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                Printer.FontSize = sFontSize
                '이전 라인 내용 찍고
                Printer.Print
                                        
                '새 라인 내용 구성
                Printer.Print Spc(11); Spc(18);
            End If
            sMonthCnt(RCnt).sMonthCnt = Format$(Val(sMonthCnt(RCnt).sMonthCnt), "##,###,##0")
            Printer.Print Spc(1); Spc(10 - Len(Trim(sMonthCnt(RCnt).sMonthCnt))); sMonthCnt(RCnt).sMonthCnt;
        Next RCnt
        
        '소계 출력
        sMonthCnt(12).sMonthCnt = Format$(Val(sMonthCnt(12).sMonthCnt), "##,###,##0")
        
        '라인을 꽉 채웠을 때
        If iMaxCnt Mod (iColMax) = 0 Then
            Printer.Print Spc(1); Spc(10 - Len(Trim(sMonthCnt(12).sMonthCnt))); sMonthCnt(12).sMonthCnt
        '라인을 덜 채웠을 때
        Else
            Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sMonthCnt(12).sMonthCnt))); sMonthCnt(12).sMonthCnt
        End If
        
        '라인 출력
        Printer.Print Spc(11); String(18 + (11 * (iColMax + 1)), Chr(45))
        
        Call Print_Tail(18 + (11 * (iColMax + 1)))
        
    Case "Y"
        dsLabStart = CDate(Left(sLabStart, 4) & "-" & "01" & "-" & "01")
        dsLabEnd = CDate(Left(sLabEnd, 4) & "-" & "12" & "-" & "31")
        
        For iCnt = 0 To 9
            If Format(dsLabEnd, "YYYY") < Format(DateAdd("yyyy", iCnt, dsLabStart), "YYYY") Then
                ' 출력숨김
                sYearCnt(iCnt).sYearTitle = ""
            Else
                '출력
                sYearCnt(iCnt).sYearTitle = Format(DateAdd("yyyy", iCnt, dsLabStart), "YYYY")
                iPrtCnt = iPrtCnt + 1
            End If
        Next iCnt
        
        If Left(sLabStart, 4) = Left(sLabEnd, 4) Then
            iPrtCnt = 1
            sYearCnt(0).sYearTitle = Format(dsLabStart, "YYYY")
        End If
        
        iMaxCnt = iPrtCnt
        
        iTotalCnt = Val(GetByOne(sResultCnt, sResultCnt))   ' 전체 갯수
        iNoCnt = 1
        iNowRow = 1
        sOldPrintNm = ""
        
        '메모리 CLEAR
        For RCnt = 0 To iMaxCnt - 1
            sYearCnt(RCnt).sYearCnt = "0"
            sYearCnt(RCnt).sYearSum = GetByOne(sSum, sSum)
        Next RCnt
        sYearCnt(10).sYearCnt = "0"
        sYearCnt(10).sYearSum = GetByOne(sSum, sSum)
        
        ' 실제 출력
        For iCnt = 1 To iTotalCnt
            iColNow = 0
            sPrtDate = Left(GetByOne(sResultCnt, sResultCnt), 6)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sdummy = GetByOne(sResultCnt, sResultCnt)
            sPrintNm = GetByOne(sResultCnt, sResultCnt)
            sRCount = GetByOne(sResultCnt, sResultCnt)
            
            '검사항목이 바뀜
            If sOldPrintNm <> sPrintNm And Trim(sOldPrintNm) <> "" Then
                '헤더 또는 라인번호 증가 또는 테일 출력
                Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                
                '검사항목 및 건수 출력
                Printer.FontSize = sFontSize
                Printer.Print Spc(11); Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));   ' No Seq
                iNoCnt = iNoCnt + 1
                Printer.Print Spc(1); LeftH(sOldPrintNm, 10); Spc(10 - LenH(Trim(sOldPrintNm))); Spc(4);
                
                For RCnt = 0 To iMaxCnt - 1
                    If (RCnt Mod iColMax = 0) And (RCnt > 0) Then
                        Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                        Printer.FontSize = sFontSize
                        '이전 라인 내용 찍고
                        Printer.Print
                                                
                        '새 라인 내용 구성
                        Printer.Print Spc(11); Spc(18);
                    End If
                    sYearCnt(RCnt).sYearCnt = Format$(Val(sYearCnt(RCnt).sYearCnt), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sYearCnt(RCnt).sYearCnt))); sYearCnt(RCnt).sYearCnt;
                Next RCnt
                
                '소계 출력
                sYearCnt(10).sYearCnt = Format$(Val(sYearCnt(10).sYearCnt), "##,###,##0")
                
                '라인을 꽉 채웠을 때
                If iMaxCnt Mod (iColMax) = 0 Then
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sYearCnt(10).sYearCnt))); sYearCnt(10).sYearCnt
                '라인을 덜 채웠을 때
                Else
                    Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sYearCnt(10).sYearCnt))); sYearCnt(10).sYearCnt
                End If
                
                '라인 출력
                Printer.Print Spc(11); String(18 + (11 * (iColMax + 1)), Chr(45))
                
                '새로운 검사항목에 대한 작업
                '메모리 CLEAR
                For RCnt = 0 To iMaxCnt
                    sYearCnt(RCnt).sYearCnt = "0"
                Next RCnt
                sYearCnt(10).sYearCnt = "0"
            End If
                
            For RCnt = 0 To iMaxCnt - 1
                If Left(sPrtDate, 4) = Left(sYearCnt(RCnt).sYearTitle, 4) Then
                    sYearCnt(10).sYearCnt = Trim(Format$(Str(Val(sYearCnt(10).sYearCnt) + Val(sRCount)), "##,###,##0"))
                    sYearCnt(RCnt).sYearCnt = Format$(Val(sRCount), "##,###,##0")
                    Exit For
                End If
            Next RCnt
            
            sOldPrintNm = sPrintNm
        Next iCnt
        
' 마지막 레코드 인쇄
        '라인 번호 증가
        Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
        
        '마지막 검사항목 및 건수 출력
        Printer.FontSize = sFontSize
        Printer.Print Spc(11); Spc(3 - Len(Trim(Str(iNoCnt)))); Trim(Str(iNoCnt));   ' No Seq
        iNoCnt = iNoCnt + 1
        Printer.Print Spc(1); LeftH(sOldPrintNm, 10); Spc(10 - LenH(Trim(sOldPrintNm))); Spc(4); ' 일통계만 다름
        
        For RCnt = 0 To iMaxCnt - 1
            If (RCnt Mod iColMax = 0) And (RCnt > 0) Then
                Call Print_CheckPage_검사건수통계(iPageLineMax, iNowRow - 1, 18 + (11 * (iColMax + 1)))
                Printer.FontSize = sFontSize
                '이전 라인 내용 찍고
                Printer.Print
                                        
                '새 라인 내용 구성
                Printer.Print Spc(11); Spc(18);
            End If
            sYearCnt(RCnt).sYearCnt = Format$(Val(sYearCnt(RCnt).sYearCnt), "##,###,##0")
            Printer.Print Spc(1); Spc(10 - Len(Trim(sYearCnt(RCnt).sYearCnt))); sYearCnt(RCnt).sYearCnt;
        Next RCnt
        
        '소계 출력
        sYearCnt(10).sYearCnt = Format$(Val(sYearCnt(10).sYearCnt), "##,###,##0")
        
        '라인을 꽉 채웠을 때
        If iMaxCnt Mod (iColMax) = 0 Then
            Printer.Print Spc(1); Spc(10 - Len(Trim(sYearCnt(10).sYearCnt))); sYearCnt(10).sYearCnt
        '라인을 덜 채웠을 때
        Else
            Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sYearCnt(10).sYearCnt))); sYearCnt(10).sYearCnt
        End If
        
        '라인 출력
        Printer.Print Spc(11); String(18 + (11 * (iColMax + 1)), Chr(45))
        
        Call Print_Tail(18 + (11 * (iColMax + 1)))
        
    End Select
    
    ' 최종 인쇄 확인
    If iTotalCnt > 0 Then
        Printer.EndDoc
        Print_ResultCnt = iTotalCnt
    Else
        Print_ResultCnt = 0
    End If


End Function

' S2) 출력위치와 크기를 미리 정의해 놓은 곳에서 읽어옴
Public Function Get_Print_Info(ByVal PrtKind As String) As Integer

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim sSql            As String
    
    Get_Print_Info = 1
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT MASTERGBN, REALGBN, MAXLEN, ALIGN, PRXPOS, PRYPOS " & _
                " FROM PRINTINFO " & _
                "WHERE PRINTGBN = '" & PrtKind & "' "

        
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;", adOpenForwardOnly
    
    Select Case PrtKind
    Case "21"   ' 년월일 검사건수 통계
        ' 출력정보 MASTER
        
        If AdoRecord.EOF = False Then
        ' 출력 자료 관리
            
            While AdoRecord.EOF = False
                If AdoRecord("MASTERGBN") & "" = "0" Then
                    If Left(AdoRecord("REALGBN") & "", 1) = "0" Then
                        sFontNm = Mid(AdoRecord("REALGBN") & "", 2) ' Font명
                        sTitleFont = AdoRecord("MAXLEN") & ""       ' Title폰트 SIZE
                        sDirectory = AdoRecord("PRXPOS") & ""
                    ElseIf AdoRecord("REALGBN") & "" = "1" Then
                        sFontSize = AdoRecord("MAXLEN") & ""        ' 인쇄시 Font크기
                    End If
                ElseIf AdoRecord("MASTERGBN") & "" = "1" Then
                    If AdoRecord("REALGBN") & "" = "HEIGHT" Then
                        iPageLineMax = Val(AdoRecord("MAXLEN") & "")
                        If iPageLineMax = 0 Then ' Default
                            iPageLineMax = 80
                        End If
                    ElseIf AdoRecord("REALGBN") & "" = "WIDTH" Then
                        iColMax = Val(AdoRecord("MAXLEN") & "")
                    End If
                End If
                AdoRecord.MoveNext
            Wend
        Else
            Get_Print_Info = 0
        End If
        AdoRecord.Close
    End Select
    
Exit Function
ErrHandler:
    MsgBox "GetItem(출력 DLL) - " & Err.Description
    Get_Print_Info = -1

End Function

' P2) 한 페이지 당 HEAD출력(검사건수통계)
Public Sub Print_Head_검사건수통계(ByVal iLineLen As Integer)
    Dim iCnt        As Integer
    
    Printer.FontName = sFontNm
    
' 메인 TITLE출력
    Printer.FontSize = sTitleFont
    Printer.Print Spc(3); Spc(1 + (iColMax * 3)); "검사건수 통계(" & sAddTitle & ")"
    Printer.Print
    
' 상위 HEAD
    Printer.FontSize = ciHeaderFntSize
    Printer.Print Spc(11); "조회구간 : "; sDate; Spc(7 + (9 * iColMax) - Len(Trim(sDate))); "출력일자 : "; Replace(Format(Now, "YYYY-MM-DD"), "-", "/")
    Printer.Print Spc(11); "SLIP  명 : "; sSlipNm; Spc(7 + (9 * iColMax) - LenH(Trim(sSlipNm))); "PAGE     :"; Printer.Page

' LINE 출력
    Printer.FontSize = ciHeaderFntSize
    Printer.Print Spc(11); String(iLineLen, Chr(6))

' 하위 HEAD
    Printer.FontSize = sFontSize
    
    Select Case LeftH(sAddTitle, 2)
        Case "일"
            Printer.Print Spc(11); "No."; Spc(1); "검사항목"; Spc(6);
        Case "월"
            Printer.Print Spc(11); "No."; Spc(1); "검사항목"; Spc(9);
        Case "년"
            Printer.Print Spc(11); "No."; Spc(1); "검사항목"; Spc(12);
    End Select
    
    For iCnt = 0 To iMaxCnt - 1
        If (iCnt Mod (iColMax) = 0) And (iCnt > 0) Then
            Printer.Print
            
            Select Case LeftH(sAddTitle, 2)
                Case "일"
                    Printer.Print Spc(18);
                Case "월"
                    Printer.Print Spc(21);
                Case "년"
                    Printer.Print Spc(24);
            End Select
        End If
            Select Case LeftH(sAddTitle, 2)
                Case "일"
                    Printer.Print Spc(1); sDayCnt(iCnt).sDayTitle; Spc(10 - LenH(Trim(sDayCnt(iCnt).sDayTitle)));
                Case "월"
                    Printer.Print Spc(1); sMonthCnt(iCnt).sMonthTitle; Spc(10 - LenH(Trim(sMonthCnt(iCnt).sMonthTitle)));
                Case "년"
                    Printer.Print Spc(1); sYearCnt(iCnt).sYearTitle; Spc(10 - LenH(Trim(sYearCnt(iCnt).sYearTitle)));
            End Select
    Next iCnt
    
    '라인을 꽉 채웠을 때
    If iMaxCnt Mod (iColMax) = 0 Then
        Select Case LeftH(sAddTitle, 2)
            Case "일"
                Printer.Print Spc(1); Spc(3); "소  계"
            Case "월"
                Printer.Print Spc(1); "소  계"
            Case "년"
                Printer.Print "소계"
        End Select
    Else
        Select Case LeftH(sAddTitle, 2)
            Case "일"
                Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(3); "소  계"
            Case "월"
                Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); "소  계"
            Case "년"
                Printer.Print Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); "소계"
        End Select
    End If
    
' LINE 출력
    Printer.FontSize = ciHeaderFntSize
    Printer.Print Spc(11); String(iLineLen, Chr(6))
    
    If Val(Printer.Page) = 1 Then
    ' 합계 출력
        Printer.FontSize = sFontSize
        Printer.Print Spc(11); Spc(3); Spc(1); "합    계"; Spc(6);
        
        For iCnt = 0 To iMaxCnt - 1
            If (iCnt Mod iColMax = 0) And (iCnt > 0) Then
                '이전 라인 내용 찍고
                Printer.Print
                
                '새 라인 내용 구성
                Printer.Print Spc(18);
            End If
            
            Select Case LeftH(sAddTitle, 2)
                Case "일"
                    sDayCnt(iCnt).sDaySum = Format$(Val(sDayCnt(iCnt).sDaySum), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sDayCnt(iCnt).sDaySum))); sDayCnt(iCnt).sDaySum;
                Case "월"
                    sMonthCnt(iCnt).sMonthSum = Format$(Val(sMonthCnt(iCnt).sMonthSum), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sMonthCnt(iCnt).sMonthSum))); sMonthCnt(iCnt).sMonthSum;
                Case "년"
                     sYearCnt(iCnt).sYearSum = Format$(Val(sYearCnt(iCnt).sYearSum), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sYearCnt(iCnt).sYearSum))); sYearCnt(iCnt).sYearSum;
            End Select
        Next
        
        '소계의 합계 출력
        If iMaxCnt Mod (iColMax) = 0 Then
            Select Case LeftH(sAddTitle, 2)
                Case "일"
                    sDayCnt(iCnt).sDaySum = Format$(Val(sDayCnt(iCnt).sDaySum), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sDayCnt(31).sDaySum))); sDayCnt(31).sDaySum
                Case "월"
                    sMonthCnt(iCnt).sMonthSum = Format$(Val(sMonthCnt(iCnt).sMonthSum), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sMonthCnt(12).sMonthSum))); sMonthCnt(12).sMonthSum
                Case "년"
                    sYearCnt(iCnt).sYearSum = Format$(Val(sYearCnt(iCnt).sYearSum), "##,###,##0")
                    Printer.Print Spc(1); Spc(10 - Len(Trim(sYearCnt(10).sYearSum))); sYearCnt(10).sYearSum
            End Select
        Else
            Select Case LeftH(sAddTitle, 2)
                Case "일"
                    sDayCnt(iCnt).sDaySum = Format$(Val(sDayCnt(iCnt).sDaySum), "##,###,##0")
                    Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sDayCnt(31).sDaySum))); sDayCnt(31).sDaySum
                Case "월"
                    sMonthCnt(iCnt).sMonthSum = Format$(Val(sMonthCnt(iCnt).sMonthSum), "##,###,##0")
                    Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sMonthCnt(12).sMonthSum))); sMonthCnt(12).sMonthSum
                Case "년"
                    sYearCnt(iCnt).sYearSum = Format$(Val(sYearCnt(iCnt).sYearSum), "##,###,##0")
                    Printer.Print Spc(1); Spc((10 + 1) * (iColMax - (iMaxCnt Mod iColMax))); Spc(10 - Len(Trim(sYearCnt(10).sYearSum))); sYearCnt(10).sYearSum
            End Select
        End If
        
    ' LINE 출력
        Printer.FontSize = ciHeaderFntSize
        Printer.Print Spc(11); String(iLineLen, Chr(6))
    End If
End Sub

' P3) 한 페이지 당 TAIL출력(공통)
Public Sub Print_Tail(ByVal iLineLen As Integer)

' LINE 출력
    Printer.Print
    Printer.FontSize = ciHeaderFntSize
    Printer.Print Spc(11); String(iLineLen, Chr(6))
' TAIL 내용 출력
    Printer.Print Spc(11); fCurAppTitle
    
End Sub

' P4) 년월일 검사건수 통계출력 PageSkip Check
Public Sub Print_CheckPage_검사건수통계(ByVal iPageLineMax As Integer, ByVal iLNowRow As Integer, ByVal iLineLen As Integer)
    
    If (iLNowRow <> 0) And (iLNowRow Mod (iPageLineMax / 2) = 0) Then
        Call Print_Tail(iLineLen)      ' TAIL인쇄후 HEAD인쇄
        Printer.NewPage
        Call Print_Head_검사건수통계(iLineLen)
    ElseIf iLNowRow = 0 Then
        Call Print_Head_검사건수통계(iLineLen)   ' HEAD만 인쇄
    End If
    iNowRow = iLNowRow + 2              ' 현재 Line번호 증가
    
End Sub

Private Sub Class_Initialize()

    Call InitializePart
    
End Sub
