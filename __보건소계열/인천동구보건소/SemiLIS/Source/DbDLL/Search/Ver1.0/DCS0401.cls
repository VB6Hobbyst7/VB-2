VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCS0401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type tPrt진료과정보
    sDeptNm     As String
End Type

Dim tPrtDept(99)            As tPrt진료과정보 '배열화 된 tPrt진료과정보

Dim iCntPrt진료과정보       As Integer

' S1) 슬립코드를 가지고 해당 슬립에 Delta 항목이 있으면 취득
Function Get_DeltaInfo(sSlipCd As String) As String

    Dim AdoRecord   As New ADODB.Recordset
    Dim iCnt        As Integer
    Dim RCnt        As Integer
    Dim sSql        As String
    Dim sDeltaCd    As String
    Dim sDeltaNm    As String
    Dim sDeltaVal   As String
    Dim sDeltaKind  As String
    Dim sDeltaInfo  As String
    Dim sTestNmCfg  As String
    
    sDeltaInfo = ""

    On Error GoTo ErrHandler
    
    sTestNmCfg = fGetCurTestItemNmCfg
    
    If sTestNmCfg = "T" Then
        sSql = " SELECT TESTITEMNM, DELTAVAL, DELTAGBN, SPECIMENCD + TESTITEMSEQ + SUBMCD AS ORDSEQ " & _
               "   FROM TESTITEM " & _
               "  WHERE PARTCD + PARTGBN = '" & sSlipCd & "'" & _
               "    AND DELTAGBN IN ('1','2') " & _
               "  ORDER BY PARTGBN, SPECIMENCD, TESTITEMSEQ "
    ElseIf sTestNmCfg = "P" Then
        sSql = " SELECT PRINTNM, DELTAVAL, DELTAGBN, SPECIMENCD + TESTITEMSEQ + SUBMCD AS ORDSEQ " & _
               "   FROM TESTITEM " & _
               "  WHERE PARTCD + PARTGBN = '" & sSlipCd & "'" & _
               "    AND DELTAGBN IN ('1','2') " & _
               "  ORDER BY PARTGBN, SPECIMENCD, TESTITEMSEQ "
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

    If AdoRecord.RecordCount <> 0 Then
' Packet {검사항목약명 | Delta값 | Delta 구분 |}n
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        For iCnt = 1 To RCnt
            sDeltaInfo = sDeltaInfo & AdoRecord(0) & "|"
            sDeltaInfo = sDeltaInfo & AdoRecord("DELTAVAL") & "|"
            sDeltaInfo = sDeltaInfo & AdoRecord("DELTAGBN") & "|"
            sDeltaInfo = sDeltaInfo & AdoRecord("ORDSEQ") & "|"
            AdoRecord.MoveNext
        Next iCnt
    Else
        sDeltaInfo = "0|"
    End If
    
    sDeltaInfo = Trim(Str(RCnt)) & "|" & sDeltaInfo
    
    AdoRecord.Close
    
    Get_DeltaInfo = sDeltaInfo

Exit Function

ErrHandler:
    MsgBox "GetItem(조회 DLL) - " & Err.Description

End Function

' S2) Delta의 값에 체크된 환자의 이전결과와 함께 결과 표시
Function Get_Delta(sDateGbn As String, sDateStr As String, eDateStr As String, sSlipCd As String, sDeltaLen As String, sOrdCd As String, orderBy As String) As String
    
    Dim AdoRecord       As New ADODB.Recordset
    Dim iCnt            As Integer
    Dim RCnt            As Integer
    Dim sSql            As String
    Dim sDelta          As String
    Dim sBefore         As String
    Dim sSex            As String
    Dim sDeptCd         As String
    Dim sJubSuGbn       As String
    Dim sEmerGbn        As String
    Dim sBeforeLabNo(999)  As String
    Dim sTestNmCfg      As String
    
' 먼저 이번에 해당 결과만 취득한다.

    sDelta = "": sBefore = ""
    
    On Error GoTo ErrHandler

    sTestNmCfg = fGetCurTestItemNmCfg

    If sTestNmCfg = "T" Then
        sSql = " SELECT JUBSU.REGNO, JUBSU.NAME, JUBSU.DEPTCD, JUBSU.JUBSUGBN, TESTITEM.TESTITEMNM, " & _
               "        JUBSU.LABDATE + '-' + JUBSU.LABSEQ AS LABNO, JUBSU.SEX + '/' + JUBSU.AGE AS SEXAGE, JUBSU.WARD, JUBSU.EMERGBN, SPECIMEN.SPECIMENBRIEFNM, RESULT.RESULT, RESULT.BEFORELABNO, RESULT.BEFORERESULT " & _
               "   FROM SPECIMEN, TESTITEM, H_JUBSU AS JUBSU, H_RESULT AS RESULT " & _
               "  WHERE RESULT." & sDateGbn & "    >= '" & sDateStr & "' " & _
               "    AND RESULT." & sDateGbn & "    <= '" & eDateStr & "' " & _
               "    AND RESULT.PARTGBN      = '" & Mid(sSlipCd, 2, 2) & "' " & sOrdCd & _
               "    AND RESULT.DELTAMARK    = 'D' " & _
               "    AND RESULT.INTERVAL     <= " & sDeltaLen & " " & _
               "    AND RESULT.BEFORELABNO  > '' " & _
               "    AND JUBSU.LABDATE       = RESULT.LABDATE " & _
               "    AND JUBSU.PARTGBN       = RESULT.PARTGBN " & _
               "    AND JUBSU.LABSEQ        = RESULT.LABSEQ " & _
               "    AND TESTITEM.PARTCD     = '" & Left(sSlipCd, 1) & "' " & _
               "    AND TESTITEM.PARTGBN    = RESULT.PARTGBN " & _
               "    AND TESTITEM.SPECIMENCD = RESULT.SPECIMENCD " & _
               "    AND TESTITEM.TESTITEMSEQ= RESULT.TESTITEMSEQ " & _
               "    AND TESTITEM.SUBMCD     = RESULT.SUBMCD " & _
               "    AND SPECIMEN.SPECIMENCD = RESULT.SPECIMENCD " & _
               orderBy
    ElseIf sTestNmCfg = "P" Then
        sSql = " SELECT JUBSU.REGNO, JUBSU.NAME, JUBSU.DEPTCD, JUBSU.JUBSUGBN, TESTITEM.PRINTNM, " & _
               "        JUBSU.LABDATE + '-' + JUBSU.LABSEQ AS LABNO, JUBSU.SEX + '/' + JUBSU.AGE AS SEXAGE, JUBSU.WARD, JUBSU.EMERGBN, SPECIMEN.SPECIMENBRIEFNM, RESULT.RESULT, RESULT.BEFORELABNO, RESULT.BEFORERESULT " & _
               "   FROM SPECIMEN, TESTITEM, H_JUBSU AS JUBSU, H_RESULT AS RESULT " & _
               "  WHERE RESULT." & sDateGbn & "    >= '" & sDateStr & "' " & _
               "    AND RESULT." & sDateGbn & "    <= '" & eDateStr & "' " & _
               "    AND RESULT.PARTGBN      = '" & Mid(sSlipCd, 2, 2) & "' " & sOrdCd & _
               "    AND RESULT.DELTAMARK    = 'D' " & _
               "    AND RESULT.INTERVAL     <= " & sDeltaLen & " " & _
               "    AND RESULT.BEFORELABNO  > '' " & _
               "    AND JUBSU.LABDATE       = RESULT.LABDATE " & _
               "    AND JUBSU.PARTGBN       = RESULT.PARTGBN " & _
               "    AND JUBSU.LABSEQ        = RESULT.LABSEQ " & _
               "    AND TESTITEM.PARTCD     = '" & Left(sSlipCd, 1) & "' " & _
               "    AND TESTITEM.PARTGBN    = RESULT.PARTGBN " & _
               "    AND TESTITEM.SPECIMENCD = RESULT.SPECIMENCD " & _
               "    AND TESTITEM.TESTITEMSEQ= RESULT.TESTITEMSEQ " & _
               "    AND TESTITEM.SUBMCD     = RESULT.SUBMCD " & _
               "    AND SPECIMEN.SPECIMENCD = RESULT.SPECIMENCD " & _
               orderBy
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet {등록번호 | 이름 | 진료과 | 접수구분 | 검사항목명 | 작업번호
'        | 성별/나이 | 병동 | 응급구분 | 검체명 | 결과 | 이전 LABNO | 이전 결과 |
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        For iCnt = 1 To RCnt
            sDelta = sDelta & AdoRecord("REGNO") & "|"
            sDelta = sDelta & AdoRecord("NAME") & "|"
            
            sDeptCd = Trim(AdoRecord("DEPTCD") & "")
            If sDeptCd <> "" Then
                sDelta = sDelta & Trim(tPrtDept(Val(sDeptCd)).sDeptNm) & "|"
            Else
                sDelta = sDelta & "|"
            End If
            
            sJubSuGbn = AdoRecord("JUBSUGBN") & ""
            Select Case sJubSuGbn
            Case "0"
                sJubSuGbn = "외래"
            Case "1"
                sJubSuGbn = "입원"
            Case "2"
                sJubSuGbn = "수탁"
            Case Else
                sJubSuGbn = ""
            End Select
            sDelta = sDelta & sJubSuGbn & "|"
            
            sDelta = sDelta & AdoRecord(4) & "|"
            sDelta = sDelta & AdoRecord("LABNO") & "|"
            sSex = AdoRecord("SEXAGE")
            If Left(Trim(sSex), 1) = "1" Or Left(Trim(sSex), 1) = "3" Then
                sSex = "남" & Mid(sSex, 2) & "|"
            ElseIf Left(Trim(sSex), 1) = "2" Or Left(Trim(sSex), 1) = "4" Then
                sSex = "여" & Mid(sSex, 2) & "|"
            Else
                sSex = "|"
            End If
            sDelta = sDelta & sSex
            sDelta = sDelta & AdoRecord("WARD") & "|"
            
            sEmerGbn = AdoRecord("EMERGBN") & ""
            If sEmerGbn = "0" Then
                sEmerGbn = "일반"
            ElseIf sEmerGbn = "1" Then
                sEmerGbn = "응급"
            End If
            sDelta = sDelta & sEmerGbn & "|"
            
            sDelta = sDelta & AdoRecord("SPECIMENBRIEFNM") & "|"
            sDelta = sDelta & AdoRecord("RESULT") & "|"
            sBeforeLabNo(iCnt) = Trim(AdoRecord("BEFORELABNO"))
            sDelta = sDelta & sBeforeLabNo(iCnt) & "|"
            sDelta = sDelta & AdoRecord("BEFORERESULT") & "|"
            AdoRecord.MoveNext
        Next iCnt
    Else
        sDelta = "0|"
    End If
    
    sDelta = Trim(Str(RCnt)) & "|" & sDelta
    
    AdoRecord.Close
    
' 이전 정보 취득
    For iCnt = 1 To RCnt
    
        sSql = " SELECT JUBSU.DEPTCD, JUBSU.JUBSUGBN, " & _
               "        JUBSU.SEX + '/' + JUBSU.AGE AS SEXAGE, JUBSU.WARD, JUBSU.EMERGBN, SPECIMEN.SPECIMENBRIEFNM " & _
               "   FROM SPECIMEN, H_JUBSU AS JUBSU " & _
               "  WHERE JUBSU." & sDateGbn & "    = '" & sDateStr & "' " & _
               "    AND JUBSU.LABDATE       = '" & Left(sBeforeLabNo(iCnt), 8) & "' " & _
               "    AND JUBSU.PARTGBN       = '" & Mid(sBeforeLabNo(iCnt), 10, 2) & "' " & _
               "    AND JUBSU.LABSEQ        = '" & Mid(sBeforeLabNo(iCnt), 12, 5) & "' " & _
               "    AND SPECIMEN.SPECIMENCD = JUBSU.SPECIMENCD "
           
        AdoRecord.CursorType = adOpenStatic
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
        If AdoRecord.RecordCount <> 0 Then
' Packet 진료과 | 접수구분 | 성별/나이 | 병동 | 응급구분 | 검체명 |
            RCnt = AdoRecord.RecordCount
            AdoRecord.MoveFirst
        
            sDeptCd = Trim(AdoRecord("DEPTCD") & "")
            If sDeptCd <> "" Then
                sBefore = sBefore & Trim(tPrtDept(Val(sDeptCd)).sDeptNm) & "|"
            Else
                sBefore = sBefore & "|"
            End If
            
            sJubSuGbn = AdoRecord("JUBSUGBN") & ""
            Select Case sJubSuGbn
            Case "0"
                sJubSuGbn = "외래"
            Case "1"
                sJubSuGbn = "입원"
            Case "2"
                sJubSuGbn = "수탁"
            Case Else
                sJubSuGbn = ""
            End Select
            sBefore = sBefore & sJubSuGbn & "|"
            
            sSex = AdoRecord("SEXAGE")
            If Left(Trim(sSex), 1) = "1" Or Left(Trim(sSex), 1) = "3" Then
                sSex = "남" & Mid(sSex, 2) & "|"
            ElseIf Left(Trim(sSex), 1) = "2" Or Left(Trim(sSex), 1) = "4" Then
                sSex = "여" & Mid(sSex, 2) & "|"
            Else
                sSex = "|"
            End If
            sBefore = sBefore & sSex
            sBefore = sBefore & AdoRecord("WARD") & "|"
        
            sEmerGbn = AdoRecord("EMERGBN") & ""
            If sEmerGbn = "0" Then
                sEmerGbn = "일반"
            ElseIf sEmerGbn = "1" Then
                sEmerGbn = "응급"
            End If
            sBefore = sBefore & sEmerGbn & "|"
        
            sBefore = sBefore & AdoRecord("SPECIMENBRIEFNM") & "|"
        Else
            sBefore = "||||||"
        End If
        sDelta = sDelta & sBefore
        
        AdoRecord.Close
    Next iCnt
    
    Get_Delta = sDelta

' 다음 결과의 내용을 취득한다.
    
Exit Function

ErrHandler:
    MsgBox "GetItem(조회 DLL) - " & Err.Description

End Function

' S3)  진료과 코드정보를 메모리에 적재
Sub Get_DeptInfo()

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim iCnt            As Integer
    Dim sSql            As String
         
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
            tPrtDept(Val(AdoRecord("DEPTCD") & "")).sDeptNm = AdoRecord("DEPTNM") & ""
            AdoRecord.MoveNext
        Next iCnt
    End If
    AdoRecord.Close
        
Exit Sub
ErrHandler:
    MsgBox "GetItem(조회 DLL) - " & Err.Description
    'AdoRecord.Close
    
End Sub
