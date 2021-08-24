VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISResultReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'+--------------------------------------------------------------------------------------+
'|  1. Class 명  : clsLISResultReview
'|  2. 기    능  : 결과내역을 조회한다.
'|  3. 작 성 자  : 김미경
'|
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+


'%  클래스 clsResultView의 Data Attributes

Public PTid As String
Public Sex As String
Public AgeDay As Long
Public ColId As String
Public RcvId As String
Public VfyId As String
Public ColDtTm As String
Public RcvDtTm As String
Public VfyDtTm As String
Public MultiFg As String
Public FootNoteFg As String
Public TestDiv As String
Public RemarkCd As String
Public StatFg As String

Public DeptCd As String
Public DeptNm As String
Public WardId As String
Public WardNm As String
Public RoomId As String
Public BedId As String
Public HosilId As String
Public BedinDt As String

Private Type ResultTable
   TestCd       As String   '/* 검사항목 코드   */
   RstCd        As String   '/* 결과코드(Alpha) */
   RstUnit      As String   '/* Unit            */
   RstDiv       As String
   HLDiv        As String   '/* High/Low(H:Hig,L:Low) */
   DPDiv        As String   '/* Delta/Panic(D:Delta,P:Panic) */
   SpcCd        As String   '/* 검체코드        */
   StatFg       As String   '/* 응급여부('0':부,'1':여) */
   LastRst      As String   '/* 최근결과        */
   LastVfyDtTm  As String   '/* 최근결과확인일  */
   'LastVfyTm As String     '/* 최근결과확인시간 */
   LastVfyId    As String   '/* 최근결과확인자  */
   VfyDt        As String   '/* 결과확인일자    */
   VfyTm        As String   '/* 결과확인시간    */
   VfyId        As String   '/* 결과확인자      */
   AttrCd       As String   '/* 속성코드        */
   MfyFg        As String   '/* 수정여부('0':부,'1':여) */
   GrpFg        As String   '/* 그래프 결과여부('0':부,'1':여) */
   TxtFg        As String   '/* TEXT 결과여부('0':부,'1':여) */
   ValFg        As String   '/* Valeu 결과여부('0':부,'1':여) - 기타검사에서만 사용 */
   RstType      As String   '/* 결과유형(N,Alpha,비율,Free) */
   DetailFg     As String   '/* 상세항목 SEQ  */
   EqpCd        As String   '/* 장비코드 */
   PTid         As String   '/* 환자 ID  */
   ORDDT        As String   '/* 처방일 */
   OrdNo        As Integer  '/* 처방번호 */
   OrdSeq       As String   '/* 처방Seq */
   RstText      As String   '/* Text 결과 */
   SuppText     As String   '/* Supplemental Report  */
   RefFromVal   As Double   '/* 기준치 From */
   RefToVal     As Double   '/* 기준치 To */
   RefCd        As String   '/* 기준치 Code */
   ClinicalNotice As String  '/* Text 결과 */
   MultiNotice  As String   '/*Multi 참고치 */
   TestShortNm  As String   '/* 검사명(Short)  */
   TestLongNm   As String   '/* 검사명(Long)  */
   RstCdNm      As String   '/* 결과코드 명 */
   SenFg        As String   '/* 감수성결과 여부 */
End Type

'연속검사 내역
Private Type MultiSpc
   WorkArea     As String
   AccDt        As String
   AccSeq       As String
   RemarkNm     As String  '/* 검체 Remark */
   FootNote     As String   '/* Foot Note */
End Type

'감수성결과 내역
Private Type SenResult
   RstCd        As String
   Row          As String
   ForeColor    As Long
End Type

Public Remark       As String  '/* 검체 Remark */
Public FootNote     As String   '/* Foot Note */
Public TextFg       As Boolean
Public SpecialFg    As Boolean
Public GeneralFg    As Boolean
Public CommentFg    As Boolean

'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As Object     'OraSession
'Public MyDb As Object     'OraDatabase

Public ResultCnt    As Integer
Public SortStartRow As Long
Public SortEndRow   As Long
Public SortFg       As Boolean
Public OffSet       As Integer
Public MicrobeCount As Long
Public SensiCount   As Long
Public TBFg         As Boolean    '항산성검사여부

Public RstRow           As Integer
Private MyResult()      As ResultTable
Private MySenResult()   As SenResult
Private RstForeColor()  As Long

Public ResultClipText   As String
Public SenClipText      As String
Public RstTextBuffer    As String
Public SamTextBuffer    As String
Public SpeTextBuffer    As String
Public SpeRstTitle      As String

Private blnMicFg(6)         As Boolean
Private lngSortStartRow(6)  As Long
Private lngSortEndRow(6)    As Long

Private objRstSQL   As New clsLISSqlReview
Private objCommSql  As New clsLISSqlStatement

Private Const MaxAntiCnt = 50
Private Const MaxColumns = 8

Private mIsMnm As Boolean   '해당접수번호의 항생제 존재여부

Public Sub SetBuildingDictionary(ByVal ObjDic As clsDictionary)
    Set objBuildings = ObjDic
End Sub

Public Sub SetDeptDictionary(ByVal ObjDic As clsDictionary)
    Set objDeptDic = ObjDic
End Sub

Public Property Get MicFg(ByVal lngIdx As Long) As Boolean
    MicFg = blnMicFg(lngIdx)
End Property

Public Property Get AntiSortStartRow(ByVal lngIdx As Long) As Long
    AntiSortStartRow = lngSortStartRow(lngIdx)
End Property

Public Property Get AntiSortEndRow(ByVal lngIdx As Long) As Long
    AntiSortEndRow = lngSortEndRow(lngIdx)
End Property

'% Method 1: Lab No를 기준으로 접수내역 및 결과내역 검색
Public Function POCResultQuery(ByVal pPtID As String, ByVal pVfyDt As String) As String

    Dim objRs As Recordset
    Dim strVfyTm As String
    
    RstRow = 0
    POCResultQuery = ""
    Set objRs = New Recordset
    objRs.Open objRstSQL.SqlGetPOCResult(pPtID, pVfyDt), DBConn
    
    ReDim Preserve RstForeColor(MaxColumns, objRs.RecordCount)
    While Not objRs.EOF
        RstRow = RstRow + 1
        strVfyTm = Mid(objRs.Fields("VfyTm").Value, 1, 2) & ":" & Mid(objRs.Fields("VfyTm").Value, 3, 2)
        POCResultQuery = POCResultQuery & _
                         strVfyTm & vbTab & objRs.Fields("RstCd").Value & vbTab & objRs.Fields("RstUnit").Value & vbTab & _
                         "" & vbTab & "by " & vbTab & objRs.Fields("VfyNm").Value & vbTab & _
                         objRs.Fields("TestNm").Value & vbTab & "" & vbTab & "" & vbTab & _
                         "" & vbTab & "" & vbTab & objRs.Fields("TestCd").Value & vbCrLf
        RstForeColor(2, RstRow) = &H747474      '진한 회색
        RstForeColor(3, RstRow) = &H404080      '갈색
        objRs.MoveNext
    Wend
    
    objRs.Close
    Set objRs = Nothing
    
End Function


'% Method 1: Lab No를 기준으로 접수내역 및 결과내역 검색
Public Function ResultQuery(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As Boolean

    Call InitRtn  '변수초기화
   
    ResultQuery = CollectQuery(pWorkArea, pAccDt, pAccSeq)
    If Not ResultQuery Then Exit Function
    
    '풋노트 & 검체리마크
    'If FootNoteFg > "0" Then
    Call ReadFootNote(pWorkArea, pAccDt, pAccSeq)
    If Trim(RemarkCd) <> "" Then Call ReadRemark(RemarkCd)
    SamTextBuffer = Remark & FootNote
    
    '결과내역 Query
    SensiCount = 0
    If Trim(MultiFg) <> "" Then
        Call MultiResult(pWorkArea, pAccDt, pAccSeq, Sex, AgeDay)
    Else
        Call GeneralResult(pWorkArea, pAccDt, pAccSeq, Sex, AgeDay)
        Call SpecialResult(pWorkArea, pAccDt, pAccSeq)
        Call MicrobeResult(pWorkArea, pAccDt, pAccSeq)
    End If
   
End Function


'% Method 1-1 : Lab No를 기준으로 감수성결과와 소견내역만 검색
Public Function ResultMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, _
                           ByVal pTestDiv As String, Optional ByVal pQuery As Boolean = True) As Boolean

    Call InitRtn  '변수초기화
   
      
        ResultMore = CollectQuery(pWorkArea, pAccDt, pAccSeq)
        If Not ResultMore Then Exit Function

        If Not pQuery Then Exit Function
      
        '풋노트 & 검체리마크
        If FootNoteFg > "0" Then Call ReadFootNote(pWorkArea, pAccDt, pAccSeq)
        If Trim(RemarkCd) <> "" Then Call ReadRemark(RemarkCd)
        SamTextBuffer = Remark & FootNote
        
        '결과내역 Query
        SensiCount = 0
        Select Case TestDiv
        Case "0", "1":  '일반검사결과,'기타검사결과
            Call GeneralMore(pWorkArea, pAccDt, pAccSeq, Sex, AgeDay)
            Call SpecialMore(pWorkArea, pAccDt, pAccSeq)
        Case "2":   '미생물검사결과
            Call MicrobeMore(pWorkArea, pAccDt, pAccSeq)
        End Select
   
End Function



'% 복수검체일 경우 : 처방번호/Seq를 기준으로 관련 검체들의 결과를 모두 가져온다.

Public Sub MultiResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, _
                       ByVal pSex As String, ByVal pAgeDay As Long)

    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim I As Integer
    Dim tmpWorkArea As String, tmpAccDt As String, tmpAccSeq As Long
   
    SqlStmt = objRstSQL.SqlMultiTest(pWorkArea, pAccDt, pAccSeq)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    With tmpRs
        If .EOF Then GoTo NoData
        While (Not .EOF)
            tmpWorkArea = "" & .Fields("WorkArea").Value
            tmpAccDt = "" & .Fields("AccDt").Value
            tmpAccSeq = Val("" & .Fields("AccSeq").Value)
            Call GeneralResult(tmpWorkArea, tmpAccDt, tmpAccSeq, pSex, pAgeDay)
            Call SpecialResult(tmpWorkArea, tmpAccDt, tmpAccSeq)
            Call MicrobeResult(tmpWorkArea, tmpAccDt, tmpAccSeq)
            .MoveNext
        Wend
    End With
         
NoData:
    tmpRs.Close
    Set tmpRs = Nothing
   
End Sub


'% 일반검사 ( LAB302 )  : Lab No를 기준으로 결과내역 검색
Public Sub GeneralResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, _
                         ByVal pSex As String, ByVal pAgeDay As Long)
    
    Dim tmpRs           As Recordset
    Dim tmpRs1          As Recordset
    Dim SqlStmt         As String
    Dim strApplyDate    As String
    Dim strLastRstVal   As String
    Dim strHLDiv        As String
    Dim strDPDiv        As String
    Dim ColCnt          As Integer
    Dim I               As Integer

   
   strApplyDate = Format(GetSystemDate, "YYYYMMDD")
   
    '결과내역 검색
    SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "0", pSex, pAgeDay)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    
    ' 그룹처방일시 배열 정리 안 함
    ' 2012-08-21 온승호
    If ResultCnt > 1 Then
        ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
    End If
    
    Dim J As Integer
    J = 1
    For I = 1 To ResultCnt
        With MyResult(J)
            '보고일자인경우 결과가 미입력일시 Skip
            If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then GoTo ResultSkip
            If Trim("" & tmpRs.Fields("RstDiv").Value) <> "*" Then
                If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                    If Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                        GoTo ResultSkip
                    End If
                End If
            End If
            
            Call AddRow(1)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            If .TestCd = P_AFBSENSCD Then TBFg = True
            .RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            .RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
            
            '-- By Mgchoi 2007/10/10
            ' - "N"을 Abnomal표시
            strHLDiv = Trim("" & tmpRs.Fields("HLDiv").Value)
            .HLDiv = IIf(strHLDiv = "N", "Abnormal", strHLDiv)
            
            '## 5.0.2: 이상대(2005-02-24)
            '   - "N"을 Abnomal표시
            strDPDiv = Trim$("" & tmpRs.Fields("DPDiv").Value)
            .DPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
                        
            '-- By 온승호 2012/07/10
            ' - "N"을 Abnomal표시
            ' - 둘다 Abnormal 이면 HLDiv를 "" 처리
            If .HLDiv = "Abnormal" And .DPDiv = "Abnormal" Then
                .HLDiv = ""
            End If
            
            .SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)
            .StatFg = Trim("" & tmpRs.Fields("StatFg").Value)
            .RstDiv = Trim("" & tmpRs.Fields("RstDiv").Value)
            
            '## 5.0.1: 이상대(2004-12-30)
            '   - 최근결과가 검사항목별 결과코드에 등록된 코드이면 결과명을 표시
            strLastRstVal = Trim$(tmpRs.Fields("lastrstval").Value & "")
            .LastRst = IIf(strLastRstVal = "", Trim("" & tmpRs.Fields("LastRst").Value), strLastRstVal)
            .LastVfyDtTm = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)
            .LastVfyId = Trim("" & tmpRs.Fields("LastVfyNm").Value)
            .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
            .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
            .AttrCd = Trim("" & tmpRs.Fields("AttrCd").Value)
            .MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)
            .GrpFg = Trim("" & tmpRs.Fields("GrpFg").Value)
            .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)
            .RstType = Trim("" & tmpRs.Fields("RstType").Value)
            .DetailFg = Trim("" & tmpRs.Fields("DetailFg").Value)
            .EqpCd = Trim("" & tmpRs.Fields("EqpCd").Value)
            .PTid = Trim("" & tmpRs.Fields("PtId").Value)
            .ORDDT = Trim("" & tmpRs.Fields("OrdDt").Value)
            .OrdNo = Val("" & tmpRs.Fields("OrdNo").Value)
            .OrdSeq = Val("" & tmpRs.Fields("OrdSeq").Value)
            .RstText = Trim("" & tmpRs.Fields("TextResult").Value)
            
            If (tmpRs.Fields("applydate").Value & "" = "" Or tmpRs.Fields("applydate").Value & "" >= strApplyDate) Then
                .ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
            End If
            
            'Supplemental Report 가 있는 경우...
            If .TxtFg = "2" Then .SuppText = GetSuppText(pWorkArea, pAccDt, pAccSeq, .TestCd)
         
            '검사명 (일반항목/상세항목)
            If .DetailFg = "" Or .RstDiv = "*" Then
                .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
            Else
                .TestShortNm = "   " & Trim("" & tmpRs.Fields("TestShortNm").Value)
            End If
            .TestLongNm = Trim("" & tmpRs.Fields("TestLongNm").Value)   '검사명 Full Name
            .RstCdNm = Trim("" & tmpRs.Fields("RstCdNm").Value)         '결과코드명

            '결과명(코드일 경우..)
            If .VfyDt = "" Then
               If .RstDiv <> "*" Then .RstCd = "미확": RstForeColor(3, RstRow) = &HC0C0FF
               .HLDiv = ""
            Else
               RstForeColor(3, RstRow) = &H404080   '갈색
               If .RstCdNm <> "" Then .RstCd = .RstCdNm
               'High / Low
               If .RstCd <> "" Then
'
'                  If .HLDiv = HLDIV_HIGH_CD Then .HLDiv = HLDIV_HIGH_FG: RstForeColor(5, RstRow) = &H7477EF '약간 붉은색
'                  If .HLDiv = HLDIV_LOW_CD Then .HLDiv = HLDIV_LOW_FG: RstForeColor(5, RstRow) = &HE48372 '약간 파란색
                  If strHLDiv = HLDIV_HIGH_CD Then .HLDiv = HLDIV_HIGH_FG: RstForeColor(5, RstRow) = &H7477EF '약간 붉은색
'                  Debug.Print RstForeColor(5, RstRow)
                  If strHLDiv = HLDIV_LOW_CD Then .HLDiv = HLDIV_LOW_FG: RstForeColor(5, RstRow) = &HE48372 '약간 파란색
               End If
            End If
         
         
            '기준치 검색
            SqlStmt = objRstSQL.SqlGetReference(.TestCd, .SpcCd, .VfyDt, "B", AgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
            If tmpRs1.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
                tmpRs1.Close
                SqlStmt = objRstSQL.SqlGetReference(.TestCd, .SpcCd, .VfyDt, Sex, AgeDay)
                Set tmpRs1 = Nothing
                Set tmpRs1 = New Recordset
                tmpRs1.Open SqlStmt, DBConn
            End If
            If tmpRs1.EOF Then
                .RefFromVal = 0: .RefToVal = 0: .RefCd = "": .MultiNotice = ""
            Else
                .RefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                .RefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                .RefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                .MultiNotice = Trim("" & tmpRs1.Fields("REFTEXT").Value & "")
                If .RefFromVal <> 0 Or .RefToVal <> 0 Then .RefCd = .RefFromVal & " - " & .RefToVal
            End If
            tmpRs1.Close
            Set tmpRs1 = Nothing
            
            '***************************************************************************
            '.RefFromVal = Val("" & tmpRs.Fields("RefValFrom").Value)
            '.RefToVal = Val("" & tmpRs.Fields("RefValTo").Value)
            '.RefCd = Trim("" & tmpRs.Fields("RefCd").Value)
            'If .RefFromVal <> 0 Or .RefToVal <> 0 Then .RefCd = .RefFromVal & " - " & .RefToVal
            '***************************************************************************
               
'            RstForeColor(2, RstRow) = &H747474 '&H747474 '진한 회색

            ' COVID 19
            If .TestCd = "COVGCRL1" Or .TestCd = "COVGCRL2" Then
                RstForeColor(2, RstRow) = &H7477EF
            Else
                RstForeColor(2, RstRow) = &H747474 '&H747474 '진한 회색
            End If
            
            ResultClipText = ResultClipText & OneRow(J)       '결과내역 Buffering
            RstTextBuffer = RstTextBuffer & OneTextResult(J)  '텍스트결과 Buffering
            
         J = J + 1
        End With
ResultSkip:
        tmpRs.MoveNext
    Next
    GeneralFg = True

NoData:
    tmpRs.Close
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Sub

'% 일반검사 ( LAB302 )  : Text 결과 및 Supplemental Report 만 검색...
Public Sub GeneralMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, _
                                     ByVal pSex As String, ByVal pAgeDay As Long)
 
   Dim I As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim tmpRs As New Recordset
   
   '결과내역 검색
   SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "0", pSex, pAgeDay)
   tmpRs.Open SqlStmt, DBConn
   If tmpRs.EOF Then GoTo NoData
   
   ResultCnt = tmpRs.RecordCount
   
   '결과내역 Array
   ReDim Preserve MyResult(ResultCnt)
   ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
   For I = 1 To ResultCnt
      
      Call AddRow(1)
      
      With MyResult(I)
         
         .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
         If .TestCd = P_AFBSENSCD Then TBFg = True
         .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
         .RstText = Trim("" & tmpRs.Fields("TextResult").Value)
         
         'Supplemental Report 가 있는 경우...
         If Trim("" & tmpRs.Fields("TxtFg").Value) = "2" Then .SuppText = GetSuppText(pWorkArea, pAccDt, pAccSeq, .TestCd)
         
         '텍스트결과 Buffering
         If .RstText <> "" Then
            ResultClipText = ResultClipText & "<< 검사 소견 >> - " & .TestShortNm & vbCrLf
            ResultClipText = ResultClipText & .RstText & vbCrLf & vbCrLf
            TextFg = True
         End If
         If .SuppText <> "" Then
            ResultClipText = ResultClipText & "<< Supplemental Report >> " & vbCrLf
            ResultClipText = ResultClipText & .SuppText & vbCrLf
            TextFg = True
         End If
         
      End With
         
      tmpRs.MoveNext
      
   'Wend
   Next
   
NoData:
   tmpRs.Close
   Set tmpRs = Nothing
End Sub

      
'% 기타검사 ( LAB351, LAB352, LAB353 ) : Lab No를 기준으로 결과내역 검색
Public Sub SpecialResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
 
    Dim I As Integer
    Dim J As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim SvTestCd As String
    Dim tmpRs As New Recordset
    Dim lngPreCnt As Long
    
    Dim objRichText As RichTextBox
    Dim objTempText As RichTextBox
    
    Dim blnFirst As Boolean
    
    '결과내역 검색
    SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "1")
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then GoTo NoData
    
    lngPreCnt = ResultCnt
    ResultCnt = ResultCnt + tmpRs.RecordCount
    
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, RstRow + ResultCnt)
    
    Set objRichText = frmControls.rtfTextBox
    Set objTempText = frmControls.rtfTempText
    
    objRichText.Text = ""
    objTempText.Text = ""
    blnFirst = True
    
    'For I = 1 To ResultCnt
    For I = lngPreCnt + 1 To ResultCnt
      
        Call AddRow(1)
        With MyResult(I)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .ValFg = Trim("" & tmpRs.Fields("ValFg").Value)
            .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)
            .MfyFg = Trim("" & tmpRs.Fields("MfySeq").Value)
            .RstType = Trim("" & tmpRs.Fields("RstType").Value)
            .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
            .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
            .PTid = Trim("" & tmpRs.Fields("PtId").Value)
            .ORDDT = Trim("" & tmpRs.Fields("OrdDt").Value)
            .OrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
            .OrdSeq = Trim("" & tmpRs.Fields("OrdSeq").Value)

            '텍스트 결과
            If .TxtFg = ERT_TxtRst Then
                
                If blnFirst Then
                    objRichText.TextRTF = "" & tmpRs.Fields("TextResult").Value
                    blnFirst = False
                Else
                    objTempText.TextRTF = "" & tmpRs.Fields("TextResult").Value
                    objRichText.Text = objRichText.Text & vbCrLf & objTempText.Text
                End If
                
                SpeRstTitle = "<< 특수검사 결과보고 >>  - " & "" & tmpRs.Fields("TestShortNm").Value
'                objRichText.Text = objRichText.Text & "<< 특수검사 결과보고 >>  - " & "" & tmpRs.Fields("TestShortNm").Value & vbCrLf & _
'                                   objRichText.Text & vbCrLf & vbCrLf
                

'                SpeTextBuffer = SpeTextBuffer & "<< 특수검사 결과보고 >>  - " & "" & tmpRs.Fields("TestShortNm").Value & vbCrLf
'                SpeTextBuffer = SpeTextBuffer & "" & tmpRs.Fields("TextResult").Value & vbCrLf & vbCrLf

'               ------------------------------------------------------------------
'               성모자애 병원에서는 Supplemental Report 사용안함 2001.7.2 김미경
'                'Supplemental Report
'                Dim SuppRs As Recordset
'                SqlStmt = objRstSql.SqlETextRst(pWorkArea, pAccDt, pAccSeq, .TestCd)
'                Set SuppRs = new recordset
'                If Not SuppRs.EOF Then
''                    SpeTextBuffer = SpeTextBuffer & "<< Supplemental Report >> " & vbCrLf
'                    objRichText.Text = objRichText.Text & "<< Supplemental Report >> " & vbCrLf
'                    While (Not SuppRs.EOF)
''                       SpeTextBuffer = SpeTextBuffer & "" & SuppRs.Fields("TextResult").Value & vbCrLf
'                        objTempText.TextRTF = "" & SuppRs.Fields("TextResult").Value
'                        objRichText.Text = objRichText.Text & objTempText.Text & vbCrLf
'                        SuppRs.MoveNext
'                    Wend
'                    SuppRs.Close
'                    Set SuppRs = Nothing
'                End If
'               ------------------------------------------------------------------
                SpecialFg = True
            End If
                  
            '검사명 (일반항목/상세항목)
            .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
            .TestLongNm = Trim("" & tmpRs.Fields("TestLongNm").Value)  '검사명 Full Name
            
            ResultClipText = ResultClipText & .TestLongNm & vbTab & vbTab
            If .VfyDt <> "" Then
                ResultClipText = ResultClipText & "☞ RESULT" & vbCrLf
                RstForeColor(4, RstRow) = DCM_LightBlue
            Else
                ResultClipText = ResultClipText & vbCrLf
            End If
            
            '상세 수치결과 검색...
            If .ValFg = ERT_ValRst Then Call GetRstValues(pWorkArea, pAccDt, pAccSeq, .TestCd, .MfyFg)
         
        End With
        tmpRs.MoveNext
      
    Next
    
    SpeTextBuffer = objRichText.TextRTF
    Unload frmControls

NoData:
   tmpRs.Close
   Set tmpRs = Nothing

End Sub
         
'% 기타검사 ( LAB351, LAB352, LAB353 ) : Text결과만...
Public Sub SpecialMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
 
    Dim I As Integer
    Dim J As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim SvTestCd As String
    Dim tmpRs As New Recordset
    Dim lngPreCnt As Long
   
    Dim objRichText As RichTextBox
    Dim objTempText As RichTextBox
    
    Dim blnFirst As Boolean
    
    '결과내역 검색
    SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "1")
    tmpRs.Open SqlStmt, DBConn
    If tmpRs.EOF Then GoTo NoData
    
    lngPreCnt = ResultCnt
    ResultCnt = ResultCnt + tmpRs.RecordCount
    
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
    
    Set objRichText = frmControls.rtfTextBox
    Set objTempText = frmControls.rtfTempText
    
    objRichText.Text = ""
    objTempText.Text = ""
    blnFirst = True
    
    For I = lngPreCnt + 1 To ResultCnt
       
        Call AddRow(1)
        With MyResult(I)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)
        
            '텍스트 결과
            If .TxtFg = ERT_TxtRst Then
                
                If blnFirst Then
                    objRichText.TextRTF = "" & tmpRs.Fields("TextResult").Value
                    blnFirst = False
                Else
                    objTempText.TextRTF = "" & tmpRs.Fields("TextResult").Value
                    objRichText.Text = objRichText.Text & vbCrLf & objTempText.Text
                End If
                
                SpeRstTitle = "<< 특수검사 결과보고 >>  - " & "" & tmpRs.Fields("TestShortNm").Value

                SpecialFg = True
            End If
                    
        End With
        tmpRs.MoveNext
       
    Next

    SpeTextBuffer = objRichText.TextRTF
    Unload frmControls

NoData:
   tmpRs.Close
   Set tmpRs = Nothing

End Sub
         
'% 미생물 검사 ( LAB404, LAB405 ) : Lab No를 기준으로 미생물 결과내역 검색
Public Sub MicrobeResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
 
    Dim I As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim lngPreCnt As Long
   
    '결과내역 검색
    SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then GoTo NoData
   
    lngPreCnt = ResultCnt
    ResultCnt = ResultCnt + tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, RstRow + ResultCnt)
   
        'For I = 1 To ResultCnt
        For I = lngPreCnt + 1 To ResultCnt
            
            With MyResult(I)
                .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
                .RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
                .RstDiv = Trim("" & tmpRs.Fields("RstDiv").Value)
                .LastRst = Trim("" & tmpRs.Fields("LastRst").Value)
                .LastVfyDtTm = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)
                .LastVfyId = Trim("" & tmpRs.Fields("LastVfyNm").Value)
                .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
                .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
                .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
                .MfyFg = Trim("" & tmpRs.Fields("MfySeq").Value)
                If .MfyFg = "0" Then .MfyFg = ""
                .RstType = Trim("" & tmpRs.Fields("RstType").Value)
                .DetailFg = Trim("" & tmpRs.Fields("DetailFg").Value)
                .PTid = Trim("" & tmpRs.Fields("PtId").Value)
                .ORDDT = Trim("" & tmpRs.Fields("OrdDt").Value)
                .OrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
                .OrdSeq = Trim("" & tmpRs.Fields("OrdSeq").Value)
                .SenFg = Trim("" & tmpRs.Fields("SenFg").Value)
               
                '감수성 결과 조회
                If .SenFg = "Y" Then
                   
                    '성모자애병원의 경우 미생물감수성 결과는 최종확인시에만 조회된다...
                    
                    If (Not P_ApplyMicMidVerify And _
                        Trim("" & tmpRs.Fields("StsCd").Value) >= enStsCd.StsCd_LIS_FinRst) Or _
                        P_ApplyMicMidVerify Then
         
                        Call AddRow(1)
                        SenClipText = ""
                        SenClipText = SenClipText & vbCrLf
            
                        Call AddRow(1)
                        RstForeColor(2, RstRow) = &H404040      '&H747474 '진한 회색
                        SenClipText = SenClipText & Trim("" & tmpRs.Fields("TestLongNm").Value) & vbCrLf
                        Call AddRow(1)
                        SenClipText = SenClipText & vbCrLf
                        
                        mIsMnm = False
                        Call SenResult(pWorkArea, pAccDt, pAccSeq, .TestCd)
                        
                        If mIsMnm Then
                            SensiCount = SensiCount + 1
                            lngSortStartRow(SensiCount) = SortStartRow
                            lngSortEndRow(SensiCount) = SortEndRow
                            
                            ResultClipText = ResultClipText & SenClipText
                            Call AddRow(1)
                            ResultClipText = ResultClipText & vbCrLf
                        End If
                    End If
                Else
                    '검사명 (일반항목/상세항목)
                    Call AddRow(1)
                    If .DetailFg = "" Or .RstDiv = "*" Then
                       .TestShortNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
                    Else
                       .TestShortNm = "   " & Trim("" & tmpRs.Fields("TestShortNm").Value)
                    End If
                    .TestLongNm = Trim("" & tmpRs.Fields("TestLongNm").Value)  '검사명 Full Name
                    .RstCdNm = Trim("" & tmpRs.Fields("RstCdNm1").Value)           '결과코드명
                    If .RstCdNm = "" Then .RstCdNm = Trim("" & tmpRs.Fields("RstCdNm2").Value)           '결과코드명
    
                   '결과명(코드일 경우..)
                    If .VfyDt = "" Then
                       If .RstDiv <> "*" Then .RstCd = "미확": RstForeColor(3, RstRow) = &HC0C0FF
                    Else
                       RstForeColor(3, RstRow) = &H404080   '갈색
                       If .RstCdNm <> "" Then .RstCd = .RstCdNm
                    End If
                    RstForeColor(2, RstRow) = &H747474  '&H404040      '&H747474 '진한 회색
                    'RstForeColor(3, RstRow) = &H404080 '갈색
                    ResultClipText = ResultClipText & OneRow(I)       '결과내역 Buffering
                    OffSet = OffSet + 1
                End If
   
            End With
            tmpRs.MoveNext
         
        'Wend
        Next

NoData:
    tmpRs.Close
    'tmpRs.CloseCursor
    Set tmpRs = Nothing
   
End Sub

'% 미생물 검사 ( LAB404, LAB405 ) : 감수성결과만...
Public Sub MicrobeMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
 
    Dim I As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
   
    '결과내역 검색
    SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
    For I = 1 To ResultCnt
            
        With MyResult(I)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .SenFg = Trim("" & tmpRs.Fields("SenFg").Value)
         
            '감수성 결과 조회
            If .SenFg = "Y" Then
                  
                '성모자애병원의 경우 미생물감수성 결과는 최종확인시에만 조회된다...
                
                If (Not P_ApplyMicMidVerify And _
                    Trim("" & tmpRs.Fields("StsCd").Value) >= enStsCd.StsCd_LIS_FinRst) Or _
                    P_ApplyMicMidVerify Then
             
                    Call AddRow(1)
                    SenClipText = ""
                    SenClipText = SenClipText & vbCrLf
                
                    Call AddRow(1)
                    RstForeColor(2, RstRow) = &H747474 '진한 회색
                    SenClipText = SenClipText & Trim("" & tmpRs.Fields("TestLongNm").Value) & vbCrLf
                    Call AddRow(1)
                    SenClipText = SenClipText & vbCrLf
                    Call SenResult(pWorkArea, pAccDt, pAccSeq, .TestCd)
                    
                    SensiCount = SensiCount + 1
                    lngSortStartRow(SensiCount) = SortStartRow
                    lngSortEndRow(SensiCount) = SortEndRow
                        
                    ResultClipText = ResultClipText & SenClipText
                    Call AddRow(1)
                    ResultClipText = ResultClipText & vbCrLf
                End If
            End If

        End With
        tmpRs.MoveNext
   
    Next

NoData:
    tmpRs.Close
    Set tmpRs = Nothing
   
End Sub

'% 미생물 검사 ( LAB404, LAB405 ) : 감수성결과만...
Public Sub MicrobeSensiRst(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)
 
    Dim I As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
   
    '결과내역 검색
    SqlStmt = objRstSQL.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
    For I = 1 To ResultCnt
            
        With MyResult(I)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .SenFg = Trim("" & tmpRs.Fields("SenFg").Value)
         
            '감수성 결과 조회
            If .SenFg = "Y" Then
                  
                '성모자애병원의 경우 미생물감수성 결과는 최종확인시에만 조회된다...
                
                If (Not P_ApplyMicMidVerify And _
                    Trim("" & tmpRs.Fields("StsCd").Value) >= enStsCd.StsCd_LIS_FinRst) Or _
                    P_ApplyMicMidVerify Then
             
                    Call AddRow(1)
                    SenClipText = ""
                    SenClipText = SenClipText & vbCrLf
                
                    Call AddRow(1)
                    RstForeColor(2, RstRow) = &H747474 '진한 회색
                    SenClipText = SenClipText & Trim("" & tmpRs.Fields("TestLongNm").Value) & vbCrLf
                    Call AddRow(1)
                    SenClipText = SenClipText & vbCrLf
                    Call GetOldSenResult(pWorkArea, pAccDt, pAccSeq, .TestCd)
                    ResultClipText = ResultClipText & SenClipText
                    Call AddRow(1)
                    ResultClipText = ResultClipText & vbCrLf
                End If
            End If

        End With
        tmpRs.MoveNext
   
    Next

NoData:
    tmpRs.Close
    Set tmpRs = Nothing
   
End Sub


'% 감수성결과 조회
Public Sub SenResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, ByVal pTestCd As String)

    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim I As Integer, J As Integer
    Dim AntiList As ListBox, AntiCnt As Integer
    Dim AntiRst As String, AntiCd As String, AntiNm As String
    Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
    Dim aryMicFg() As String
    Dim tmpTitle As String
    Dim tmpAntiCnt As Integer
   
    SqlStmt = objRstSQL.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo NoData
   
    MicroCnt = tmpRs.RecordCount  '균갯수
    MicrobeCount = MicroCnt
    ReDim MySenResult(MaxAntiCnt, MicroCnt)
    ReDim aryMicFg(MicroCnt)
   
    Set AntiList = frmControls.lstUnsortedList   '항생제 리스트(Unsorted)
    AntiList.Clear
   
    tmpAntiCnt = 0
    tmpTitle = "Antibiotics  " & Space(26 - Len("Antibiotics  ")) & vbTab
   
    '미생물 감수성 결과내역 Buffering (LAB405)
    For I = 1 To MicroCnt
       
        Call AddRow(1)
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        
        AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
        aryMicFg(I) = "" & tmpRs.Fields("MicFg").Value  'MIC여부
        
        If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
        tmpTitle = tmpTitle & " " & CStr(I) & " " & Space(15 - Len(CStr(I))) ' & vbTab

        '## 5.0.3: 이상대(2005-05-30)
        '   - 정도명이 한줄에 전부 표시되도록 수정
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
        
        Call AddRow(1)
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        
        GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
        SenClipText = SenClipText & "   " & GrowthQty & vbCrLf '배양정도
        
        For J = 1 To AntiCnt
            AntiRst = Trim("" & tmpRs.Fields("SRst" & CStr(J)).Value)
            AntiCd = medShift(AntiRst, ";")
            AntiSeq = medListFind(AntiList, AntiCd & ";")
            '새로운 항생제가 있을 경우만 리스트에 Add...
            If (AntiSeq < 0) Or (AntiCd & ";" <> AntiList.List(AntiSeq)) Then
                AntiList.AddItem AntiCd & ";"
                AntiSeq = AntiList.ListCount - 1
            End If
            
            MySenResult(AntiSeq, I).RstCd = AntiRst    '결과 Keeping
        Next
        tmpRs.MoveNext
    Next
   
    Call AddRow(1): SenClipText = SenClipText & vbCrLf
   
    If tmpAntiCnt > 0 Then  '감수성결과가 있으면...
        Call AddRow(1): SenClipText = SenClipText & "[ Susceptibility test ]" & vbCrLf
        RstForeColor(2, RstRow) = &HE48372  '약간 파란색
        Call AddRow(1): SenClipText = SenClipText & tmpTitle & vbCrLf
        RstForeColor(2, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(3, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(4, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(5, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(6, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(7, RstRow) = &H7477EF '약간 붉은색
    End If

    '감수성 결과 Buffering
    SortFg = False
   
    Dim strRst1 As String, strRst2 As String
   
    For I = 1 To AntiList.ListCount
        Call AddRow(1)
        
        If I = 1 Then SortStartRow = RstRow
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        AntiCd = AntiList.List(I - 1)
        AntiCd = Mid(AntiCd, 1, Len(AntiCd) - 1)
        
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open objCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
        If Not tmpRs1.EOF Then
            AntiNm = Trim("" & tmpRs1.Fields("text1").Value)   '항생제명
        End If
        
        If Trim(AntiNm) = "" Then AntiNm = AntiCd
        If Len(AntiNm) > 25 Then
            SenClipText = SenClipText & Mid(AntiNm, 1, 25) & Space(1) & vbTab  '항생제명
        Else
            SenClipText = SenClipText & Mid(AntiNm, 1, 25) & Space(26 - Len(AntiNm)) & vbTab  '항생제명
        End If
        tmpRs1.Close
        Set tmpRs1 = Nothing
        
        For J = 1 To MicroCnt
            strRst1 = medGetP(MySenResult(I - 1, J).RstCd, 1, ";")  'SRI결과
            strRst2 = medGetP(MySenResult(I - 1, J).RstCd, 2, ";")  'MIC결과
            If strRst1 = "" Then
                SenClipText = SenClipText & Space(17) '& vbTab
            Else
                If strRst1 = "N" Then
                    strRst1 = "NEGATIVE"
                ElseIf strRst1 = "P" Then
                    strRst1 = "POSITIVE"
                ElseIf strRst1 = "-" Then
                    If AntiNm <> "ESBL" Then
                        strRst1 = "SYN-R"
                    Else
                        strRst1 = strRst1
                    End If
                End If
                
                If Trim(strRst2) = "" Then
                    If Len(strRst1) > 15 Then
                        SenClipText = SenClipText & " " & strRst1 & Space(1) & vbTab
                    Else
                        SenClipText = SenClipText & " " & strRst1 & Space(15 - Len(strRst1)) & vbTab
                    End If
                    If Not blnMicFg(J) Then blnMicFg(J) = False
                Else
                    strRst1 = strRst1 & " (" & Trim(strRst2) & ")"
                    If Len(strRst1) > 15 Then
                        SenClipText = SenClipText & " " & strRst1 & Space(1) & vbTab
                    Else
                        SenClipText = SenClipText & " " & strRst1 & Space(15 - Len(strRst1)) & vbTab
                    End If
                    blnMicFg(J) = True
                End If
                
                If strRst1 Like "S*" Then
                    RstForeColor(J + 2, RstRow) = vbRed '&H7477EF
                    RstForeColor(2, RstRow) = &H404080  '&H7477EF
                Else
                    RstForeColor(J + 2, RstRow) = &H747474
                End If
                
            End If
        Next
        
        For J = MicroCnt + 1 To 5
            SenClipText = SenClipText & vbTab
        Next
        SenClipText = SenClipText & vbTab & AntiNm & vbCrLf
        SortFg = True
    Next
    SortEndRow = RstRow
    Unload frmControls
    mIsMnm = True
   
NoData:
   tmpRs.Close
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   Set AntiList = Nothing
End Sub

'% 감수성결과 조회
Public Sub GetOldSenResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, ByVal pTestCd As String)

    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim I As Integer, J As Integer
    Dim AntiList As ListBox, AntiCnt As Integer
    Dim AntiRst As String, AntiCd As String, AntiNm As String
    Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
    Dim aryMicFg() As String
    Dim tmpTitle As String
    Dim tmpAntiCnt As Integer
   
    SqlStmt = objRstSQL.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo NoData
   
    MicroCnt = tmpRs.RecordCount  '균갯수
    MicrobeCount = MicroCnt
    ReDim MySenResult(MaxAntiCnt, MicroCnt)
    ReDim aryMicFg(MicroCnt)
   
    Set AntiList = frmControls.lstUnsortedList   '항생제 리스트(Unsorted)
    AntiList.Clear
   
    tmpAntiCnt = 0
    tmpTitle = "Antibiotics  " & vbTab
   
    '미생물 감수성 결과내역 Buffering (LAB405)
    For I = 1 To MicroCnt
       
        Call AddRow(1)
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(4, RstRow) = &H404080 '갈색
        
        AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
        aryMicFg(I) = "" & tmpRs.Fields("MicFg").Value  'MIC여부
        
        If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
        If AntiCnt > 0 Then
            tmpTitle = tmpTitle & " " & CStr(I) & " " & vbTab
        End If
        GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
        SenClipText = SenClipText & GrowthQty & vbTab & vbTab '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
        
        For J = 1 To AntiCnt
            AntiRst = Trim("" & tmpRs.Fields("SRst" & CStr(J)).Value)
            AntiCd = medShift(AntiRst, ";")
            AntiSeq = medListFind(AntiList, AntiCd)
            '새로운 항생제가 있을 경우만 리스트에 Add...
            If (AntiSeq < 0) Or (AntiCd <> AntiList.List(AntiSeq)) Then
                AntiList.AddItem AntiCd
                AntiSeq = AntiList.ListCount - 1
            End If
            
            MySenResult(AntiSeq, I).RstCd = AntiRst    '결과 Keeping
        Next
        tmpRs.MoveNext
    Next
   
    Call AddRow(1): SenClipText = SenClipText & vbCrLf
   
    If tmpAntiCnt > 0 Then  '감수성결과가 있으면...
        Call AddRow(1): SenClipText = SenClipText & "[ Susceptibility test ]" & vbCrLf
        RstForeColor(2, RstRow) = &HE48372  '약간 파란색
        Call AddRow(1): SenClipText = SenClipText & tmpTitle & vbCrLf
        RstForeColor(2, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(3, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(4, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(5, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(6, RstRow) = &H7477EF '약간 붉은색
        RstForeColor(7, RstRow) = &H7477EF '약간 붉은색
    End If

    '감수성 결과 Buffering
    SortFg = False
   
    Dim strRst1 As String, strRst2 As String
   
    For I = 1 To AntiList.ListCount
        Call AddRow(1)
        
        If I = 1 Then SortStartRow = RstRow
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        AntiCd = AntiList.List(I - 1)
        
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open objCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
        
        If Not tmpRs1.EOF Then
            AntiNm = Trim("" & tmpRs1.Fields("Text1").Value)   '항생제명
        End If
        tmpRs1.Close
        
        SenClipText = SenClipText & AntiCd & vbTab  '항생제코드
        
        For J = 1 To MicroCnt
            strRst1 = medGetP(MySenResult(I - 1, J).RstCd, 1, ";")  'SRI결과
            strRst2 = medGetP(MySenResult(I - 1, J).RstCd, 2, ";")  'MIC결과
            If strRst1 = "" Then
                SenClipText = SenClipText & vbTab
            Else
                If strRst1 = "N" Then
                    strRst1 = "NEGATIVE"
                ElseIf strRst1 = "P" Then
                    strRst1 = "POSITIVE"
                ElseIf strRst1 = "-" Then
                    strRst1 = "SYN-R"
                End If
                
                If Trim(strRst2) = "" Then
                    SenClipText = SenClipText & " " & strRst1 & vbTab
                    If Not blnMicFg(J) Then blnMicFg(J) = False
                Else
                    SenClipText = SenClipText & " " & strRst1 & " (" & Trim(strRst2) & ")" & vbTab
                    blnMicFg(J) = True
                End If
                If strRst1 Like "S*" Then
                    RstForeColor(J + 2, RstRow) = vbRed '&H7477EF
                    RstForeColor(2, RstRow) = &H404080  '&H7477EF
                Else
                    RstForeColor(J + 2, RstRow) = &H747474
                End If
                
            End If
        Next
        For J = MicroCnt + 1 To 5
            SenClipText = SenClipText & vbTab
        Next
        SenClipText = SenClipText & vbTab & vbTab & AntiNm & vbCrLf
        SortFg = True
    Next
    SortEndRow = RstRow
   
NoData:
   tmpRs.Close
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   Set AntiList = Nothing
   
End Sub


'% 검체리마크 Text
'% -. LAB034 : CdIndex-'C404'의 Text1
Public Function ReadRemark(ByVal RmkCd As String) As String
   
   Dim tmpSQL As String
   Dim tmpRs As Recordset
   
   Remark = ""
   ReadRemark = ""
   tmpSQL = objRstSQL.SqlGetRemark(RmkCd)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSQL, DBConn
   
   If tmpRs.EOF Then GoTo NoData
   
   ReadRemark = Trim("" & tmpRs.Fields("Remark").Value)
   Remark = "<< Remark >>" & vbCrLf & Trim("" & tmpRs.Fields("Remark").Value) & vbCrLf & vbCrLf
   CommentFg = True

NoData:
   tmpRs.Close
   Set tmpRs = Nothing
   
End Function

'% FootNote Text
'% -. LAB304 : RstTxt
Public Function ReadFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String
   
    Dim I As Integer
    Dim tmpSQL As String
    Dim tmpRs As Recordset
   
    FootNote = ""
    ReadFootNote = ""
    tmpSQL = objRstSQL.SqlGetFootNote(pWorkArea, pAccDt, pAccSeq)
    Set tmpRs = New Recordset
    tmpRs.Open tmpSQL, DBConn
    
    If tmpRs.EOF Then GoTo NoData
   
    FootNote = "<< Foot Note >>" & vbCrLf
    ReadFootNote = Trim("" & tmpRs.Fields("FootNote").Value)
    While (Not tmpRs.EOF)
        If Trim("" & tmpRs.Fields("FootNote").Value) <> "" Then
            FootNote = FootNote & Trim("" & tmpRs.Fields("FootNote").Value) & vbCrLf
        Else
            FootNote = ""
        End If
        tmpRs.MoveNext
    Wend
    CommentFg = True
   
NoData:
    tmpRs.Close
    Set tmpRs = Nothing
   
End Function


'% 결과내역 Spread에 Display될 Clip Text를 작성한다.
'% -. 검사명(Short), 결과, 단위, High/Low, Delta/Panic, 기준치, 검사명(Long), 최근결과, 최근결과확인일, 속성코드
Public Function OneRow(ByVal Row As Integer)

    Dim strNotice   As String
    Dim strRow      As String
    
    OneRow = ""
    If Row > UBound(MyResult) Then Exit Function
    With MyResult(Row)
        OneRow = .TestShortNm & vbTab & .RstCd & vbTab & .RstUnit & vbTab & .HLDiv & vbTab & .DPDiv & vbTab & _
                 .RefCd & vbTab & .TestLongNm & vbTab & .LastRst & vbTab & .LastVfyDtTm & vbTab & _
                 .AttrCd & vbTab & .MfyFg & vbTab & .TestCd & vbCrLf
        
        If Trim(.MultiNotice) <> "" Then
            Call AddRow(1)
            RstForeColor(1, RstRow) = DCM_LightBlue
            OneRow = OneRow & "☞ Multi Reference " & Space(42) & vbCrLf
            strNotice = Replace(.MultiNotice, vbCr, "")
            strRow = medShift(strNotice, vbLf)
            While strRow <> ""
                Call AddRow(1)
                RstForeColor(1, RstRow) = &H747474
                OneRow = OneRow & strRow & Space(60 - Len(Mid(strRow, 1, 60))) & vbCrLf
                strRow = medShift(strNotice, vbLf)
            Wend
        End If
        strNotice = ""
        
        If Trim(.ClinicalNotice) <> "" Then
            Call AddRow(1)
            RstForeColor(1, RstRow) = DCM_LightRed
            'vbTab & vbTab &
            OneRow = OneRow & "☞ Clinical Notice " & Space(42) & vbCrLf
            strNotice = Replace(.ClinicalNotice, vbCr, "")
            strRow = medShift(strNotice, vbLf)
            While strRow <> ""
                Call AddRow(1)
                RstForeColor(1, RstRow) = &H747474
                OneRow = OneRow & strRow & Space(60 - Len(Mid(strRow, 1, 60))) & vbCrLf
                strRow = medShift(strNotice, vbLf)
            Wend
        End If
        
    End With
   
End Function

'% 일반결과 - 텍스트결과 내역 (수정내역포함)  Buffring
Public Function OneTextResult(ByVal Row As Integer)

   OneTextResult = ""
   If Row > UBound(MyResult) Then Exit Function
   If MyResult(Row).RstText <> "" Then
      OneTextResult = OneTextResult & "<< 검사 소견 >> - " & MyResult(Row).TestShortNm & vbCrLf
      OneTextResult = OneTextResult & MyResult(Row).RstText & vbCrLf & vbCrLf
      TextFg = True
   End If
   If MyResult(Row).SuppText <> "" Then
      OneTextResult = OneTextResult & "<< Supplemental Report >> " & vbCrLf
      OneTextResult = OneTextResult & MyResult(Row).SuppText & vbCrLf
      TextFg = True
   End If
      
End Function


'% Work Area 명 : LAB032 (C213) - Field1
Public Function GetWorkAreaNm(ByVal WorkAreaCd As String)

   Dim tmpRs As Recordset
   
   Set tmpRs = New Recordset
   tmpRs.Open objCommSql.SqlCommonCode(T_LAB032, LC3_WorkArea, WorkAreaCd), DBConn
   If tmpRs.EOF Then
      GetWorkAreaNm = ""
   Else
      GetWorkAreaNm = Trim("" & tmpRs.Fields("Field1").Value)
   End If
   tmpRs.Close
   Set tmpRs = Nothing
   
End Function

'% 텍스트결과 내역
'% -. LAB303
Public Function GetRstText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, ByVal pTestCd As String)

   Dim I As Integer
   Dim tmpRs As Recordset
   
   Set tmpRs = New Recordset
   tmpRs.Open objRstSQL.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
   If tmpRs.EOF Then
      GetRstText = ""
   Else
      GetRstText = Trim("" & tmpRs.Fields("RstTxt").Value)
   End If
   tmpRs.Close
   Set tmpRs = Nothing
   
End Function


'% 텍스트결과 수정내역
'% -. LAB305
Public Function GetSuppText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, ByVal pTestCd As String)

   Dim I As Integer
   Dim tmpRs As Recordset
   
   Set tmpRs = New Recordset
   tmpRs.Open objRstSQL.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
   
   If tmpRs.EOF Then
      GetSuppText = ""
   Else
      GetSuppText = ""
      While (Not tmpRs.EOF)
         GetSuppText = GetSuppText & Trim("" & tmpRs.Fields("RstTxt").Value) & vbCrLf & _
                              "♧ 수정자 : " & Trim("" & tmpRs.Fields("EmpNm").Value) & "   " & Trim("" & tmpRs.Fields("MfyDtTm").Value) & vbCrLf
         GetSuppText = GetSuppText & vbCrLf
         tmpRs.MoveNext
      Wend
      GetSuppText = Mid(GetSuppText, 1, Len(GetSuppText) - 2)
   End If
   tmpRs.Close
   Set tmpRs = Nothing
   
End Function

'% 기타검사 Value Result ( LAB352 )
Public Sub GetRstValues(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, _
                                    ByVal pTestCd As String, ByVal MfyFg As String)
   
   Dim tmpRs As Recordset
   Dim tmpSQL As String
   Dim I As Integer
   
   tmpSQL = objRstSQL.SqlEValueRst(pWorkArea, pAccDt, pAccSeq, pTestCd, MfyFg)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSQL, DBConn
   
   If tmpRs.EOF Then GoTo NoData
   
   While (Not tmpRs.EOF)
      Call AddRow(1)
      ResultClipText = ResultClipText & "   " & Trim("" & tmpRs.Fields("Head").Value) & vbTab
      ResultClipText = ResultClipText & Trim("" & tmpRs.Fields("ValRst").Value) & vbTab
      ResultClipText = ResultClipText & Trim("" & tmpRs.Fields("Unit").Value) & vbCrLf
      RstForeColor(2, RstRow) = &H747474 '진한 회색
      RstForeColor(3, RstRow) = &H404080 '갈색
      tmpRs.MoveNext
   Wend
   
NoData:
   tmpRs.Close
   Set tmpRs = Nothing

End Sub

'% 직원 성명
'Public Function GetEmpNm(ByVal EmpId As String)
'
'   Dim tmpRs As Recordset
'   Dim SqlStmt As String
'
'   GetEmpNm = ""
'   If EmpId = "" Then Exit Function
'
'   SqlStmt = objCommSql.SqlLAB015Read(EmpId, 0)
'   Set tmpRs = New Recordset
'   tmpRs.Open SqlStmt, dbconn
'
'   If tmpRs.EOF Then
'      GetEmpNm = ""
'   Else
'      GetEmpNm = Trim("" & tmpRs.Fields("EmpNm").Value)
'   End If
'   tmpRs.Close
'   Set tmpRs = Nothing
'
'End Function

'% 의사이름
'Public Function GetDoctNm(ByVal DoctId As String)
'
'   Dim tmpRs As Recordset
'   Dim SqlStmt As String
'
'   GetDoctNm = ""
'   If DoctId = "" Then Exit Function
'
'   SqlStmt = objCommSql.SqlHIS007CodeList(DoctId)
'   Set tmpRs = New Recordset
'   tmpRs.Open SqlStmt, dbconn
'
'   If tmpRs.EOF Then
'      GetDoctNm = ""
'   Else
'      GetDoctNm = Trim("" & tmpRs.Fields("EmpNm").Value)
'   End If
'   tmpRs.Close
'   Set tmpRs = Nothing
'
'End Function

'% 부서이름
'Public Function GetDeptNm(ByVal DeptCd As String)
''    Dim objDept As clsBasisData
'
'    GetDeptNm = ""
'    If DeptCd = "" Then Exit Function
'
''    Set objDept = New clsBasisData
'
'    GetDeptNm = GetDeptNm(DeptCd)
'
''    Set objDept = Nothing
'
''    'SqlStmt = objCommSql.SqlHIS003CodeList(DeptCd)
''    'Set tmpRs = new recordset
''
''    If Not ObjLISComCode.DeptCd.Exists(DeptCd) Then
''        GetDeptNm = ""
''    Else
''        ObjLISComCode.DeptCd.KeyChange (DeptCd)
''        GetDeptNm = ObjLISComCode.DeptCd.Fields("DeptNm")
''    End If
''    'tmpRs.Close
''    'Set tmpRs = Nothing
'
'End Function

Private Sub AddRow(ByVal Rows As Integer)
    RstRow = RstRow + Rows
    ReDim Preserve RstForeColor(MaxColumns, RstRow)
   
End Sub


'% High/Low 컬럼 ForeColor
Public Function Get_ForeColor(ByVal Col As Integer, ByVal Row As Integer) As Long

    Get_ForeColor = RstForeColor(Col, Row)
    'Get_HL_ForeColor = MyResult(Row).HLForeColor
   
End Function

'% Initialize Routine
Private Sub InitRtn()
   
    '변수 초기화
    RstRow = 0
    ResultCnt = 0
    Remark = ""
    FootNote = ""
    ResultClipText = ""
    SenClipText = ""
    SamTextBuffer = ""
    RstTextBuffer = ""
    SortStartRow = 0
    SortEndRow = 0
    OffSet = 0
    
    CommentFg = False
    TextFg = False
    SortFg = False
    
    GeneralFg = False
    SpecialFg = False
    
    Erase MyResult
    Erase MySenResult
    ReDim MyResult(1)
    ReDim RstForeColor(MaxColumns, 1)

End Sub

Public Sub GetRelTest(ByRef objCombo As Object, ByVal pLabNo As String)
    Dim tmpRs       As New Recordset
    Dim SqlStmt     As String
    Dim pWorkArea   As String
    Dim pAccDt      As String
    Dim pAccSeq     As String
    Dim strRstVal   As String
    Dim I           As Long
   
    pWorkArea = medGetP(pLabNo, 1, "-")
    pAccDt = medGetP(pLabNo, 2, "-")
    pAccSeq = medGetP(pLabNo, 3, "-")
   
    If Mid(pAccDt, 1, 1) = "9" Then
        pAccDt = "19" & pAccDt
    Else
        pAccDt = "20" & pAccDt
    End If
' 2009.07.22 양성현 검체코드 추가
    SqlStmt = objRstSQL.SqlGetRelTest(pWorkArea, pAccDt, pAccSeq)
    tmpRs.Open SqlStmt, DBConn
   
    objCombo.Clear
    If tmpRs.EOF Then
        objCombo.AddItem "< 없음 >"
        GoTo NoData
    End If
       
    '## 5.0.1: 이상대(2004-12-30)
    '   - 검사항목별 결과코드에 등록된 결과코드이면 결과값표시
    With tmpRs
        For I = 1 To .RecordCount
            strRstVal = .Fields("rstval").Value & ""
            strRstVal = IIf(strRstVal = "", .Fields("rstcd").Value & "", strRstVal)
            objCombo.AddItem "" & .Fields("AbbrNm5").Value & "   " & strRstVal & "   " & _
                             "" & .Fields("RstUnit").Value & "   " & .Fields("spcnm").Value & "   " & .Fields("VfyDtTm").Value & "   " & _
                             "by " & "" & .Fields("VfyNm").Value & Space(50) & vbTab & _
                             .Fields("testcd").Value & "" & vbTab & _
                             .Fields("testdiv").Value & "" & vbTab & _
                             .Fields("labno").Value & "" & vbTab & _
                             .Fields("vfydttm").Value & ""
            .MoveNext
        Next
    End With
   
NoData:
    tmpRs.Close
    Set tmpRs = Nothing
End Sub

Public Sub GetMicRelTest(ByRef objCombo As Object, ByVal pLabNo As String)
    Dim tmpRs       As New Recordset
    Dim SqlStmt     As String
    Dim pWorkArea   As String
    Dim pAccDt      As String
    Dim pAccSeq     As String
    Dim strRstVal   As String
    Dim I, J        As Long
   
    pWorkArea = medGetP(pLabNo, 1, "-")
    pAccDt = medGetP(pLabNo, 2, "-")
    pAccSeq = medGetP(pLabNo, 3, "-")
   
'    If Mid(pAccDt, 1, 1) = "9" Then
'        pAccDt = "19" & pAccDt
'    Else
'        pAccDt = "20" & pAccDt
'    End If
    
    SqlStmt = objRstSQL.SqlGetMicRelTest(pWorkArea, pAccDt, pAccSeq)
    tmpRs.Open SqlStmt, DBConn
   
    objCombo.Clear
    If tmpRs.EOF Then
        objCombo.AddItem "< 없음 >"
        GoTo NoData
    End If
       
    '## 5.0.1: 이상대(2004-12-30)
    '   - 검사항목별 결과코드에 등록된 결과코드이면 결과값표시
    With tmpRs
        For I = 1 To .RecordCount
    ' 090520 양성현 검체명 추가
            strRstVal = .Fields("rstval").Value & ""
            strRstVal = IIf(strRstVal = "", .Fields("rstcd").Value & "", strRstVal)
            SqlStmt = "" & .Fields("spcnm").Value & " | " & strRstVal
            If Len(SqlStmt) < 35 Then
                For J = 0 To (35 - Len(SqlStmt))
                    SqlStmt = SqlStmt & " "
                Next J
            End If
            objCombo.AddItem "" & .Fields("AbbrNm5").Value & " | " & _
                             SqlStmt & _
                             "" & .Fields("RstUnit").Value & " | " & _
                             "" & .Fields("VfyDtTm").Value & " | " & _
                             "by " & "" & .Fields("VfyNm").Value & Space(50) & vbTab & _
                             .Fields("testcd").Value & "" & vbTab & _
                             .Fields("testdiv").Value & "" & vbTab & _
                             .Fields("labno").Value & "" & vbTab & _
                             .Fields("vfydttm").Value & ""
            .MoveNext
        Next
    End With
   
NoData:
    tmpRs.Close
    Set tmpRs = Nothing
End Sub

Public Function GetTAT(ByVal pTestCd As String, ByVal pSpcCd As String, ByVal pErFg As String) As String

    Dim tmpRs As New Recordset
    
    tmpRs.Open objRstSQL.SqlGetTAT(pTestCd, pSpcCd), DBConn
    
    GetTAT = ""
    If Not tmpRs.EOF Then
        If Trim(pErFg) = "" Then '일반
            GetTAT = medGetP("" & tmpRs.Fields("TATS").Value, 1, ";")
        Else
            GetTAT = medGetP("" & tmpRs.Fields("TATS").Value, 2, ";")
        End If
    End If
    tmpRs.Close
    Set tmpRs = Nothing
    
End Function


Public Function GetBuildNoForTAT(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Long

    Dim tmpRs As New Recordset
    
    tmpRs.Open objRstSQL.SqlGetBuildNoForTAT(pWorkArea, pAccDt, pAccSeq), DBConn
    If Not tmpRs.EOF Then
        GetBuildNoForTAT = Val("" & tmpRs.Fields("BldNo").Value)
        tmpRs.Close
    Else
        GetBuildNoForTAT = 1
    End If
    Set tmpRs = Nothing
    
End Function


Public Function CollectQuery(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean

    Dim tmpRs As New Recordset
    
    tmpRs.Open objRstSQL.SqlGetCollectData(pWorkArea, pAccDt, pAccSeq), DBConn
    
    If Not tmpRs.EOF Then
        With tmpRs
            PTid = "" & .Fields("PtId").Value
            Sex = "" & .Fields("Sex").Value
            AgeDay = "" & .Fields("AgeDay").Value
            ' 2018.01.29  채혈/접수 시 채혈자ID 누락으로 접수자로 대체 함
            If "" & .Fields("ColId").Value = "" Then
                ColId = "" & .Fields("RcvId").Value
            Else
                ColId = "" & .Fields("ColId").Value
            End If
            RcvId = "" & .Fields("RcvId").Value
            VfyId = "" & .Fields("VfyId").Value
            ColDtTm = Format("" & .Fields("ColDt").Value, CS_DateMask) & " " & Format(Mid("" & .Fields("ColTm").Value, 1, 4), CS_TimeShortMask)
            RcvDtTm = Format("" & .Fields("RcvDt").Value, CS_DateMask) & " " & Format(Mid("" & .Fields("RcvTm").Value, 1, 4), CS_TimeShortMask)
            VfyDtTm = Format("" & .Fields("VfyDt").Value, CS_DateMask) & " " & Format(Mid("" & .Fields("VfyTm").Value, 1, 4), CS_TimeShortMask)
            MultiFg = "" & .Fields("MultiFg").Value
            FootNoteFg = "" & .Fields("FootNoteFg").Value
            TestDiv = "" & .Fields("TestDiv").Value
            RemarkCd = "" & .Fields("RmkCd").Value
            StatFg = "" & .Fields("StatFg").Value
            
            DeptCd = "" & .Fields("DeptCd").Value
            DeptNm = GetDeptNm(DeptCd)
            WardId = "" & .Fields("WardId").Value
            RoomId = "" & .Fields("RoomId").Value
            HosilId = "" & .Fields("HosilId").Value
            BedId = "" & .Fields("BedId").Value
            BedinDt = Format("" & .Fields("BedInDt").Value, CS_DateMask)
        End With
        CollectQuery = True
    Else
        CollectQuery = False
    End If
    Set tmpRs = Nothing
    
End Function


Public Function Archive_Data(ByVal pDeptCd As String) As Boolean

    Dim strYesterday As String
    Dim SqlStmt As String
    Dim strTmp As String
    
    On Error GoTo Err_Trap
    
    strYesterday = Format(DateAdd("d", -30, GetSystemDate), CS_DateDbFormat)
    SqlStmt = objRstSQL.SqlArchiveVerifiedList(pDeptCd, strYesterday)
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    Archive_Data = True
    
    Exit Function
    
Err_Trap:
    MsgBox Err.Description, vbExclamation
    Archive_Data = False

End Function


Public Function Get_VerifiedList(ByVal pDeptCd As String, ByVal pDoneFg As String, _
                                  ByRef tblPtList As Object)

    Dim I As Integer
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim DivRow As Long
    
    If pDeptCd = "" Then pDeptCd = "ER"
    
    DivRow = 0
'    QueryFg = True
    
    Set tmpRs = New Recordset
    tmpRs.Open objRstSQL.SqlGetVerifiedList(pDeptCd, pDoneFg), DBConn
    
    With tblPtList
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = tmpRs.RecordCount
        For I = 1 To tmpRs.RecordCount
            .Row = I
            .Col = 1: .CellType = CellTypeCheckBox
                      .Value = Val("" & tmpRs.Fields("DoneFg").Value)
                      If .Value = 1 Then .Col = -1: .ForeColor = &H808080
            .Col = 2: .Value = "" & tmpRs.Fields("PtId").Value
            .Col = 3: .Value = "" & tmpRs.Fields("PtNm").Value
            .Col = 4: .Value = "" & tmpRs.Fields("VfyTm").Value
            .Col = 5: .Value = Choose(Val("" & tmpRs.Fields("MfyFg").Value) + 1, "", "*")
            .Col = 6: .Value = "" & tmpRs.Fields("VfyDt").Value
            .Col = 7: .Value = "" & tmpRs.Fields("DeptCd").Value
            .Col = 8: .Value = "" & tmpRs.Fields("OrdDiv").Value
            
            If DivRow = 0 And "" & tmpRs.Fields("VfyDt").Value = Format(Now, CS_DateDbFormat) Then DivRow = I
            tmpRs.MoveNext
        Next
        
        If DivRow > 0 Then
            .Row = 1: .Row2 = DivRow - 1
            .Col = -1
            .BlockMode = True
            .BackColor = RGB(252, 252, 252)
            .BlockMode = False
        End If
        .ReDraw = True
    End With
          
    tmpRs.Close
    Set tmpRs = Nothing
'    QueryFg = False
    
End Function

Public Function Get_PtListByDoctor(ByVal pDoctId As String, ByRef tblPtList As Object)

    Dim I As Integer
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim DivRow As Long
    
    DivRow = 0
'    QueryFg = True
    
    Set tmpRs = New Recordset
    tmpRs.Open objRstSQL.SqlPtListByDoctor(pDoctId), DBConn
    
    With tblPtList
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = tmpRs.RecordCount
        For I = 1 To tmpRs.RecordCount
            .Row = I
            '-- 변경 Colum Type CheckBox 유지 By Choi.M.G
            '- 원본 =====================================
            '.Col = 1: .CellType = CellTypeStaticText
            '          .Value = ""
            '============================================
            .Col = 1: .Value = ""
            .Col = 2: .Value = "" & tmpRs.Fields("PtId").Value
            .Col = 3: .Value = "" & tmpRs.Fields("PtNm").Value
            .Col = 4: .Value = "" & tmpRs.Fields("WardId").Value & "-" & tmpRs.Fields("RoomId").Value
            .Col = 6: .Value = Format(Now, CS_DateDbFormat)
            .Col = 8: .Value = "" & tmpRs.Fields("OrdDiv").Value
            
            tmpRs.MoveNext
        Next
        
    End With
          
    tmpRs.Close
    Set tmpRs = Nothing
'    QueryFg = False
    
End Function


Public Function UpdateVerifiedList(ByVal pDeptCd As String, ByVal pVerifyDt As String, _
                                   ByVal pPtID As String, ByVal pMfyFg As String, _
                                   ByVal pEmpId As String) As Boolean

    Dim SqlStmt As String
            
    On Error GoTo Err_Trap
    SqlStmt = objRstSQL.SqlUpdateVerifiedList(pDeptCd, pVerifyDt, pPtID, pMfyFg, pEmpId)
    Call DBConn.BeginTrans
    Call DBConn.Execute(SqlStmt)
    Call DBConn.CommitTrans
    UpdateVerifiedList = True
    Exit Function
    
Err_Trap:
    MsgBox Err.Description, vbExclamation
    UpdateVerifiedList = False

End Function

Public Function SetStartDt(ByVal pDeptCd As String) As Date
    
    Dim RS As Recordset
    Dim sDate As Date
    
    Set RS = New Recordset
    RS.Open objRstSQL.SqlGetStartDate(pDeptCd), DBConn
    If Not RS.EOF Then
        SetStartDt = DateAdd("m", Val("" & RS.Fields("Field1").Value) * (-1), Now)
        RS.Close
    Else
        SetStartDt = DateAdd("m", -1, Now)
    End If
    Set RS = Nothing
    
End Function


Public Sub LoadWorkArea(ByRef cboWorkArea As Object)

    Dim I%
    Dim SqlStmt As String
    Dim tmpRs As Recordset

    Set tmpRs = New Recordset
    tmpRs.Open objRstSQL.SqlGetWorkAreaList, DBConn

    If tmpRs.EOF = True Then ' record가 존재하지 않을 경우
        tmpRs.Close
        Set tmpRs = Nothing
        Exit Sub
    End If

    cboWorkArea.Clear
    With tmpRs
        .MoveFirst
        For I = 1 To .RecordCount
            cboWorkArea.AddItem "" & .Fields("cdval1").Value & "  " & "" & .Fields("field1").Value
            .MoveNext
        Next I
    End With
    cboWorkArea.ListIndex = 0
    
    tmpRs.Close
    Set tmpRs = Nothing

End Sub

Public Sub LoadItem(ByRef lstList As Object, ByVal pWorkArea As String, ByVal pSpcCd As String)

    Dim SqlStmt As String
    Dim RS As Recordset
    Dim tmpStr As String
    Dim I%
    
    '상세항목 제외...
    Set RS = New Recordset
    RS.Open objRstSQL.SqlGetItemList(pWorkArea, pSpcCd), DBConn
    
    If RS.EOF Then GoTo NoData
    
    lstList.Clear
    For I = 1 To RS.RecordCount
        tmpStr = "" & RS.Fields("TestCd").Value & Space(9)
        lstList.AddItem Mid(tmpStr, 1, 10) & _
                        "" & RS.Fields("TestNm").Value & Space(30) & vbTab & _
                        "" & RS.Fields("TestDiv").Value & vbTab & _
                        "" & RS.Fields("WorkArea").Value & vbTab & _
                        "" & RS.Fields("PanelFg").Value & vbTab & _
                        pSpcCd & vbTab & _
                        "" & RS.Fields("SpcNm").Value
        RS.MoveNext
    Next I
    
NoData:
    RS.Close
    Set RS = Nothing
    
End Sub
Public Function GetOrderDept(ByVal PTid As String, ByVal ORDDT As String, ByVal OrdNo As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " select deptcd from " & T_LAB101 & " where " & _
                    DBW("ptid=", PTid) & " and " & DBW("orddt=", ORDDT) & " and " & DBW("ordno=", OrdNo)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        GetOrderDept = GetDeptNm(RS.Fields("deptcd").Value & "")
    End If
    Set RS = Nothing
End Function
Public Function GetOrderColid(ByVal PTid As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " select colid ,coldt,coltm,rcvid from " & T_BBS201 & " where " & _
                    DBW("ptid=", PTid) & " and " & _
                    " coldt=(select max(coldt) from " & T_BBS201 & " where " & DBW("ptid=", PTid) & ")"
                 
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    '******************************************
    ' 수정자 : 온승호
    ' 2010.12.07
    ' 혈액은행 결과조회에서 채혈자/접수자 누락
    ' GestDoctNm => GetEmpNm 으로 수정
    '******************************************
    
    If Not RS.EOF Then
        GetOrderColid = GetEmpNm(Trim(RS.Fields("colid").Value & ""))
        GetOrderColid = GetOrderColid & COL_DIV & Format(RS.Fields("coldt").Value & "", "####-##-##") & " " & Format(Mid(RS.Fields("coltm").Value & "", 1, 4), "##:##")
        GetOrderColid = GetOrderColid & COL_DIV & GetEmpNm(Trim(RS.Fields("rcvid").Value & ""))
    End If
    Set RS = Nothing
End Function


Public Function GetOrderRemark(ByVal sWorkarea As String, ByVal sAccDt As String, ByVal sAccSeq As String) As String
    Dim SSQL As String
    
    SSQL = " select a.orddt,a.ordno,a.ordseq,a.ordcd,b.abbrnm10 ,a.mesg " & _
           " from " & T_LAB001 & " b," & T_LAB102 & " a " & _
           " where " & _
                     DBW("a.workarea=", sWorkarea) & _
           " and " & DBW("a.accdt=", sAccDt) & _
           " and " & DBW("a.accseq=", sAccSeq) & _
           " and  (a.mesg <>'' or a.mesg is not null)" & _
           " and a.ordcd=b.testcd" & _
           " order by a.ordno,a.ordseq"
    GetOrderRemark = SSQL
    
    Debug.Print GetOrderRemark
End Function
