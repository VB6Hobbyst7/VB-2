VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCmtResultReview"
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

Public ptid As String
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
Public BedID As String
Public HosilId As String
Public BedinDt As String

Private Type ResultTable
   TestCd As String  '/* 검사항목 코드   */
   RstCd As String  '/* 결과코드(Alpha) */
   RstUnit As String  '/* Unit            */
   RstDiv As String
   HLDiv As String  '/* High/Low(H:Hig,L:Low) */
   DPDiv As String  '/* Delta/Panic(D:Delta,P:Panic) */
   SpcCd As String  '/* 검체코드        */
   StatFg As String  '/* 응급여부('0':부,'1':여) */
   LastRst As String '/* 최근결과        */
   LastVfyDtTm As String '/* 최근결과확인일  */
   'LastVfyTm As String '/* 최근결과확인시간 */
   LastVfyId As String  '/* 최근결과확인자  */
   VfyDt As String  '/* 결과확인일자    */
   VfyTm As String  '/* 결과확인시간    */
   VfyId As String  '/* 결과확인자      */
   AttrCd As String  '/* 속성코드        */
   MfyFg As String  '/* 수정여부('0':부,'1':여) */
   GrpFg As String  '/* 그래프 결과여부('0':부,'1':여) */
   TxtFg As String  '/* TEXT 결과여부('0':부,'1':여) */
   ValFg As String  '/* Valeu 결과여부('0':부,'1':여) - 기타검사에서만 사용 */
   RstType As String  '/* 결과유형(N,Alpha,비율,Free) */
   DetailFg As String  '/* 상세항목 SEQ  */
   EqpCd As String  '/* 장비코드 */
   ptid As String  '/* 환자 ID  */
   OrdDt As String  '/* 처방일 */
   OrdNo As Integer  '/* 처방번호 */
   OrdSeq As String  '/* 처방Seq */
   
   RstText As String  '/* Text 결과 */
   SuppText As String  '/* Supplemental Report  */
   
   RefFromVal As Double  '/* 기준치 From */
   RefToVal As Double  '/* 기준치 To */
   RefCd As String  '/* 기준치 Code */
   ClinicalNotice As String  '/* Text 결과 */
   
   TestShortNm As String  '/* 검사명(Short)  */
   TestLongNm As String  '/* 검사명(Long)  */
   RstCdNm As String '/* 결과코드 명 */
   
   SenFg As String  '/* 감수성결과 여부 */
End Type

'연속검사 내역
Private Type MultiSpc
   WorkArea As String
   AccDt As String
   AccSeq As String
   RemarkNm As String  '/* 검체 Remark */
   FootNote As String   '/* Foot Note */
End Type

'감수성결과 내역
Private Type SenResult
   RstCd As String
   Row As String
   ForeColor As Long
End Type

Public Remark As String  '/* 검체 Remark */
Public FootNote As String   '/* Foot Note */
Public TextFg As Boolean
Public SpecialFg As Boolean
Public GeneralFg As Boolean
Public CommentFg As Boolean

'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As Object     'OraSession
'Public MyDb As Object     'OraDatabase

Public ResultCnt As Integer
Public SortStartRow As Long
Public SortEndRow As Long
Public SortFg As Boolean
Public OffSet As Integer
Public MicrobeCount As Long
Public SensiCount As Long

Public RstRow As Integer
Private MyResult() As ResultTable
Private MySenResult() As SenResult
Private RstForeColor() As Long

Public ResultClipText As String
Public SenClipText As String
Public RstTextBuffer As String
Public SamTextBuffer As String
Public SpeTextBuffer As String
Public SpeRstTitle As String

Private blnMicFg(6) As Boolean
Private lngSortStartRow(6) As Long
Private lngSortEndRow(6) As Long

Private objRstSql As New clsLISSqlReview
Private objCommSql As New clsLISSqlStatement
'Private objCollect As New clsLISCollectioin   '접수내역

Private Const MaxAntiCnt = 50
Private Const MaxColumns = 8

'Public Sub SetDatabase(ByVal objDB As DrDatabase)
'    Set DBConn = objDB
'End Sub
'
'Public Sub SetSession(ByVal objSE As Object)
'    Set dbSess = objSE
'End Sub

'Public Sub SetBuildingDictionary(ByVal objDic As clsDictionary)
'    Set objBuildings = objDic
'End Sub

'Public Sub SetDeptDictionary(ByVal objDic As clsDictionary)
'    Set objDeptDic = objDic
'End Sub

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
Public Function POCResultQuery(ByVal pPtId As String, ByVal pVfyDt As String) As String

    Dim objRs As Recordset
    Dim strVfyTm As String
    
    RstRow = 0
    POCResultQuery = ""
    Set objRs = New Recordset 'OpenRecordSet(objRstSql.SqlGetPOCResult(pPtId, pVfyDt))
    objRs.Open objRstSql.SqlGetPOCResult(pPtId, pVfyDt), DBConn
    
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
    
'    objRs.RsClose
    Set objRs = Nothing
End Function

'% Method 1: Lab No를 기준으로 접수내역 및 결과내역 검색
Public Function ResultQuery(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

    Call InitRtn  '변수초기화
   
    ResultQuery = CollectQuery(pWorkArea, pAccDt, pAccSeq)
    If Not ResultQuery Then Exit Function
    
    '풋노트 & 검체리마크
    If FootNoteFg > "0" Then Call ReadFootNote(pWorkArea, pAccDt, pAccSeq)
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
Public Function ResultMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
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

Public Sub MultiResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                       ByVal pSex As String, ByVal pAgeDay As Long)

    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim i As Integer
    Dim tmpWorkArea As String, tmpAccDt As String, tmpAccSeq As Integer
   
    SqlStmt = objRstSql.SqlMultiTest(pWorkArea, pAccDt, pAccSeq)
    Set tmpRs = New Recordset ' OpenRecordSet(SqlStmt)
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
'    tmpRs.RsClose
    Set tmpRs = Nothing
   
End Sub


'% 일반검사 ( LAB302 )  : Lab No를 기준으로 결과내역 검색
Public Sub GeneralResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                         ByVal pSex As String, ByVal pAgeDay As Long)
 
    Dim i As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As Recordset
   
    '결과내역 검색
    SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "0", pSex, pAgeDay)
'    Set tmpRs = OpenRecordSet(SqlStmt)
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
    For i = 1 To ResultCnt
      
        Call AddRow(1)
        With MyResult(i)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            .RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
            .HLDiv = Trim("" & tmpRs.Fields("HLDiv").Value)
            .DPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
            .SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)
            .StatFg = Trim("" & tmpRs.Fields("StatFg").Value)
            .RstDiv = Trim("" & tmpRs.Fields("RstDiv").Value)
            .LastRst = Trim("" & tmpRs.Fields("LastRst").Value)
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
            .ptid = Trim("" & tmpRs.Fields("PtId").Value)
            .OrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
            .OrdNo = Val("" & tmpRs.Fields("OrdNo").Value)
            .OrdSeq = Val("" & tmpRs.Fields("OrdSeq").Value)

            .RstText = Trim("" & tmpRs.Fields("TextResult").Value)
            .ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
         
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
                  If .HLDiv = HLDIV_HIGH_CD Then .HLDiv = HLDIV_HIGH_FG: RstForeColor(5, RstRow) = &H7477EF '약간 붉은색
                  If .HLDiv = HLDIV_LOW_CD Then .HLDiv = HLDIV_LOW_FG: RstForeColor(5, RstRow) = &HE48372 '약간 파란색
               End If
            End If
         
            '기준치 검색
            SqlStmt = objRstSql.SqlGetReference(.TestCd, .SpcCd, .VfyDt, "B", AgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
'            Set tmpRs1 = OpenRecordSet(SqlStmt)
            If tmpRs1.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
'               tmpRs1.RsClose
               SqlStmt = objRstSql.SqlGetReference(.TestCd, .SpcCd, .VfyDt, Sex, AgeDay)
               Set tmpRs1 = Nothing
               Set tmpRs1 = New Recordset
               tmpRs1.Open SqlStmt, DBConn
'               Set tmpRs1 = OpenRecordSet(SqlStmt)
            End If
            If tmpRs1.EOF Then
               .RefFromVal = 0: .RefToVal = 0: .RefCd = ""
            Else
               .RefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
               .RefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
               .RefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
               If .RefFromVal <> 0 Or .RefToVal <> 0 Then .RefCd = .RefFromVal & " - " & .RefToVal
            End If
'            tmpRs1.RsClose
            Set tmpRs1 = Nothing
            '***************************************************************************
            '.RefFromVal = Val("" & tmpRs.Fields("RefValFrom").Value)
            '.RefToVal = Val("" & tmpRs.Fields("RefValTo").Value)
            '.RefCd = Trim("" & tmpRs.Fields("RefCd").Value)
            'If .RefFromVal <> 0 Or .RefToVal <> 0 Then .RefCd = .RefFromVal & " - " & .RefToVal
            '***************************************************************************
               
            RstForeColor(2, RstRow) = &H747474 '진한 회색
            'RstForeColor(3, RstRow) = &H404080 '갈색
            ResultClipText = ResultClipText & OneRow(i)       '결과내역 Buffering
            RstTextBuffer = RstTextBuffer & OneTextResult(i)  '텍스트결과 Buffering
         
        End With
         
        tmpRs.MoveNext
      
   'Wend
    Next
    GeneralFg = True
   
NoData:
'    tmpRs.RsClose
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Sub

'% 일반검사 ( LAB302 )  : Text 결과 및 Supplemental Report 만 검색...
Public Sub GeneralMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                                     ByVal pSex As String, ByVal pAgeDay As Long)
 
   Dim i As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim tmpRs As New Recordset
   Dim tmpRs1 As Recordset
   
   '결과내역 검색
   SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "0", pSex, pAgeDay)
   tmpRs.Open SqlStmt, DBConn
'   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then GoTo NoData
   
   ResultCnt = tmpRs.RecordCount
   
   '결과내역 Array
   ReDim Preserve MyResult(ResultCnt)
   ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
   For i = 1 To ResultCnt
      
      Call AddRow(1)
      
      With MyResult(i)
         
         .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
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
'   tmpRs.RsClose
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   
End Sub

      
'% 기타검사 ( LAB351, LAB352, LAB353 ) : Lab No를 기준으로 결과내역 검색
Public Sub SpecialResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim i As Integer
    Dim j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim SvTestCd As String
    Dim tmpRs As New Recordset
    Dim lngPreCnt As Long
    
    Dim objRichText As RichTextBox
    Dim objTempText As RichTextBox
    
    Dim blnFirst As Boolean
    
    '결과내역 검색
    SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "1")
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
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
    For i = lngPreCnt + 1 To ResultCnt
      
        Call AddRow(1)
        With MyResult(i)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .ValFg = Trim("" & tmpRs.Fields("ValFg").Value)
            .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)
            .MfyFg = Trim("" & tmpRs.Fields("MfySeq").Value)
            .RstType = Trim("" & tmpRs.Fields("RstType").Value)
            .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
            .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
            .ptid = Trim("" & tmpRs.Fields("PtId").Value)
            .OrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
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
'                Dim SuppRs As recordset
'                SqlStmt = objRstSql.SqlETextRst(pWorkArea, pAccDt, pAccSeq, .TestCd)
'                Set SuppRs = OpenRecordSet(SqlStmt)
'                If Not SuppRs.EOF Then
''                    SpeTextBuffer = SpeTextBuffer & "<< Supplemental Report >> " & vbCrLf
'                    objRichText.Text = objRichText.Text & "<< Supplemental Report >> " & vbCrLf
'                    While (Not SuppRs.EOF)
''                       SpeTextBuffer = SpeTextBuffer & "" & SuppRs.Fields("TextResult").Value & vbCrLf
'                        objTempText.TextRTF = "" & SuppRs.Fields("TextResult").Value
'                        objRichText.Text = objRichText.Text & objTempText.Text & vbCrLf
'                        SuppRs.MoveNext
'                    Wend
'                    SuppRs.RsClose
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
'   tmpRs.RsClose
   Set tmpRs = Nothing

End Sub
         
'% 기타검사 ( LAB351, LAB352, LAB353 ) : Text결과만...
Public Sub SpecialMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim i As Integer
    Dim j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim SvTestCd As String
    Dim tmpRs As New Recordset
    Dim lngPreCnt As Long
   
    Dim objRichText As RichTextBox
    Dim objTempText As RichTextBox
    
    Dim blnFirst As Boolean
    
    '결과내역 검색
    SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "1")
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
    
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
    
    For i = lngPreCnt + 1 To ResultCnt
       
        Call AddRow(1)
        With MyResult(i)
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

'                SpeTextBuffer = SpeTextBuffer & "<< 특수검사 결과보고 >>  - " & "" & tmpRs.Fields("TestShortNm").Value & vbCrLf
'                SpeTextBuffer = SpeTextBuffer & "" & tmpRs.Fields("TextResult").Value & vbCrLf & vbCrLf
'                'Supplemental Report
'                Dim SuppRs As recordset
'                SqlStmt = objRstSql.SqlETextRst(pWorkArea, pAccDt, pAccSeq, .TestCd)
'                Set SuppRs = OpenRecordSet(SqlStmt)
'                If Not SuppRs.EOF Then
'                    SpeTextBuffer = SpeTextBuffer & "<< Supplemental Report >> " & vbCrLf
'                    While (Not SuppRs.EOF)
'                        SpeTextBuffer = SpeTextBuffer & "" & SuppRs.Fields("TextResult").Value & vbCrLf
'                        SuppRs.MoveNext
'                    Wend
'                    SuppRs.RsClose
'                    Set SuppRs = Nothing
'                End If
                SpecialFg = True
            End If
                    
        End With
        tmpRs.MoveNext
       
    Next

    SpeTextBuffer = objRichText.TextRTF
    Unload frmControls

NoData:
'   tmpRs.RsClose
   Set tmpRs = Nothing

End Sub
         
'% 미생물 검사 ( LAB404, LAB405 ) : Lab No를 기준으로 미생물 결과내역 검색
Public Sub MicrobeResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim i As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim lngPreCnt As Long
   
    '결과내역 검색
    SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
    If tmpRs.EOF Then GoTo NoData
   
    lngPreCnt = ResultCnt
    ResultCnt = ResultCnt + tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, RstRow + ResultCnt)
   
        'For I = 1 To ResultCnt
        For i = lngPreCnt + 1 To ResultCnt
            
            With MyResult(i)
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
                .ptid = Trim("" & tmpRs.Fields("PtId").Value)
                .OrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
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
                        Call SenResult(pWorkArea, pAccDt, pAccSeq, .TestCd)
                        
                        SensiCount = SensiCount + 1
                        lngSortStartRow(SensiCount) = SortStartRow
                        lngSortEndRow(SensiCount) = SortEndRow
                        
                        ResultClipText = ResultClipText & SenClipText
                        Call AddRow(1)
                        ResultClipText = ResultClipText & vbCrLf
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
                    ResultClipText = ResultClipText & OneRow(i)       '결과내역 Buffering
                    OffSet = OffSet + 1
                End If
   
            End With
            tmpRs.MoveNext
         
        'Wend
        Next

NoData:
'    tmpRs.RsClose
    'tmpRs.CloseCursor
    Set tmpRs = Nothing
   
End Sub

'% 미생물 검사 ( LAB404, LAB405 ) : 감수성결과만...
Public Sub MicrobeMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim i As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
   
    '결과내역 검색
    SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
    For i = 1 To ResultCnt
            
        With MyResult(i)
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
'    tmpRs.RsClose
    Set tmpRs = Nothing
   
End Sub

'% 미생물 검사 ( LAB404, LAB405 ) : 감수성결과만...
Public Sub MicrobeSensiRst(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim i As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
   
    '결과내역 검색
    SqlStmt = objRstSql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '결과내역 Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
    For i = 1 To ResultCnt
            
        With MyResult(i)
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
'    tmpRs.RsClose
    Set tmpRs = Nothing
   
End Sub


'% 감수성결과 조회
Public Sub SenResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim i As Integer, j As Integer
    Dim AntiList As ListBox, AntiCnt As Integer
    Dim AntiRst As String, AntiCd As String, AntiNm As String
    Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
    Dim aryMicFg() As String
    Dim tmpTitle As String
    Dim tmpAntiCnt As Integer
   
    SqlStmt = objRstSql.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
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
    For i = 1 To MicroCnt
       
        Call AddRow(1)
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        
        AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
        aryMicFg(i) = "" & tmpRs.Fields("MicFg").Value  'MIC여부
        
        If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
        If AntiCnt > 0 Then
            tmpTitle = tmpTitle & " " & CStr(i) & " " & Space(11 - Len(CStr(i))) & vbTab
'            If aryMicFg(I) = MRT_GenSen Then
'                tmpTitle = tmpTitle & " " & CStr(I) & " "
'            ElseIf aryMicFg(I) = MRT_MicSen Then
'                tmpTitle = tmpTitle & " " & CStr(I) & Space(7) & " "
'            End If
        End If
        GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
        'SenClipText = SenClipText & GrowthQty & Space(6 - Len(GrowthQty)) & vbTab '배양정도
        SenClipText = SenClipText & GrowthQty & vbTab '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
        
        For j = 1 To AntiCnt
            AntiRst = Trim("" & tmpRs.Fields("SRst" & CStr(j)).Value)
            AntiCd = medShift(AntiRst, ";")
            AntiSeq = medListFind(AntiList, AntiCd & ";")
            '새로운 항생제가 있을 경우만 리스트에 Add...
            If (AntiSeq < 0) Or (AntiCd & ";" <> AntiList.List(AntiSeq)) Then
                AntiList.AddItem AntiCd & ";"
                AntiSeq = AntiList.ListCount - 1
            End If
            
            MySenResult(AntiSeq, i).RstCd = AntiRst    '결과 Keeping
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
   
    For i = 1 To AntiList.ListCount
        Call AddRow(1)
        
        If i = 1 Then SortStartRow = RstRow
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        AntiCd = AntiList.List(i - 1)
        AntiCd = Mid(AntiCd, 1, Len(AntiCd) - 1)
        
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open objCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
        
'        Set tmpRs1 = OpenRecordSet(objCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd))
        'AntiNm = Mid(Trim(tmpRs1.Fields("Text1").Value), 1, 15)   '항생제명
        'SenClipText = SenClipText & AntiNm & Space(15 - Len(AntiNm)) & vbTab  '항생제명
        If tmpRs1.EOF = False Then
            AntiNm = Trim("" & tmpRs1.Fields("Text1").Value)   '항생제명
        End If
        
        If Trim(AntiNm) = "" Then AntiNm = AntiCd
        
        If Len(AntiNm) > 26 Then
            SenClipText = SenClipText & Mid(AntiNm, 1, 25) & vbTab  '항생제명
        Else
            SenClipText = SenClipText & Mid(AntiNm, 1, 25) & Space(26 - Len(AntiNm)) & vbTab  '항생제명
        End If
        
'        tmpRs1.RsClose
        
        For j = 1 To MicroCnt
            
            strRst1 = medGetP(MySenResult(i - 1, j).RstCd, 1, ";")  'SRI결과
            strRst2 = medGetP(MySenResult(i - 1, j).RstCd, 2, ";")  'MIC결과
            If strRst1 = "" Then
                SenClipText = SenClipText & Space(13) & vbTab
'                If aryMicFg(j) = MRT_MicSen Then
'                    SenClipText = SenClipText & "   " & Space(7)
'                Else
'                    SenClipText = SenClipText & "   "
'                End If
            Else
'                SenClipText = SenClipText & " " & strRst1  '균별 감수성결과
                If strRst1 = "N" Then
                    strRst1 = "NEGATIVE"
                ElseIf strRst1 = "P" Then
                    strRst1 = "POSITIVE"
                ElseIf strRst1 = "-" Then
                    strRst1 = "SYN-R"
                End If
                
                If Trim(strRst2) = "" Then
                    SenClipText = SenClipText & " " & strRst1 & Space(11 - Len(strRst1)) & vbTab
                    If Not blnMicFg(j) Then blnMicFg(j) = False
                Else
                    strRst1 = strRst1 & " (" & Trim(strRst2) & ")"
                    SenClipText = SenClipText & " " & strRst1 & Space(11 - Len(strRst1)) & vbTab
                    blnMicFg(j) = True
                End If
                If strRst1 Like "S*" Then
                    RstForeColor(j + 2, RstRow) = vbRed '&H7477EF
                    RstForeColor(2, RstRow) = &H404080  '&H7477EF
                Else
                    RstForeColor(j + 2, RstRow) = &H747474
                End If
                
'                If strRst2 <> "" Then
'                    If Mid(strRst2, 1, 1) = "≤" Or Mid(strRst2, 1, 1) = "≥" Then
'                        SenClipText = SenClipText & "(" & Trim(strRst2) & ")" & Space(6 - Len(strRst2) - 1)
'                    Else
'                        SenClipText = SenClipText & "(" & Trim(strRst2) & ")" & Space(6 - Len(strRst2))
'                    End If
'                Else
'                    SenClipText = SenClipText & " "
'                End If
            End If
        Next
        For j = MicroCnt + 1 To 5
            SenClipText = SenClipText & vbTab
        Next
        SenClipText = SenClipText & AntiNm & vbCrLf
'        SenClipText = SenClipText & "                        " & vbTab
'        SenClipText = SenClipText & vbTab & vbTab & vbTab & vbTab & AntiNm & vbCrLf
        'Call AddRow(1)
        SortFg = True
    Next
    SortEndRow = RstRow
    
    Unload frmControls
   'SortFg = True
   
NoData:
'   tmpRs.RsClose
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   Set AntiList = Nothing
   
End Sub

'% 감수성결과 조회
Public Sub GetOldSenResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim i As Integer, j As Integer
    Dim AntiList As ListBox, AntiCnt As Integer
    Dim AntiRst As String, AntiCd As String, AntiNm As String
    Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
    Dim aryMicFg() As String
    Dim tmpTitle As String
    Dim tmpAntiCnt As Integer
   
    SqlStmt = objRstSql.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
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
    For i = 1 To MicroCnt
       
        Call AddRow(1)
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(4, RstRow) = &H404080 '갈색
        
        AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
        aryMicFg(i) = "" & tmpRs.Fields("MicFg").Value  'MIC여부
        
        If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
        If AntiCnt > 0 Then
            tmpTitle = tmpTitle & " " & CStr(i) & " " & vbTab
        End If
        GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
        SenClipText = SenClipText & GrowthQty & vbTab & vbTab '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
        
        For j = 1 To AntiCnt
            AntiRst = Trim("" & tmpRs.Fields("SRst" & CStr(j)).Value)
            AntiCd = medShift(AntiRst, ";")
            AntiSeq = medListFind(AntiList, AntiCd)
            '새로운 항생제가 있을 경우만 리스트에 Add...
            If (AntiSeq < 0) Or (AntiCd <> AntiList.List(AntiSeq)) Then
                AntiList.AddItem AntiCd
                AntiSeq = AntiList.ListCount - 1
            End If
            
            MySenResult(AntiSeq, i).RstCd = AntiRst    '결과 Keeping
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
   
    For i = 1 To AntiList.ListCount
        Call AddRow(1)
        
        If i = 1 Then SortStartRow = RstRow
        RstForeColor(2, RstRow) = &H747474 '진한 회색
        RstForeColor(3, RstRow) = &H404080 '갈색
        AntiCd = AntiList.List(i - 1)
        
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open objCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
        
'        Set tmpRs1 = OpenRecordSet(objCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd))
        If Not tmpRs1.EOF Then
            AntiNm = Trim("" & tmpRs1.Fields("Text1").Value)   '항생제명
        End If
'        tmpRs1.RsClose
        
        SenClipText = SenClipText & AntiCd & vbTab  '항생제코드
        
        For j = 1 To MicroCnt
            strRst1 = medGetP(MySenResult(i - 1, j).RstCd, 1, ";")  'SRI결과
            strRst2 = medGetP(MySenResult(i - 1, j).RstCd, 2, ";")  'MIC결과
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
                    If Not blnMicFg(j) Then blnMicFg(j) = False
                Else
                    SenClipText = SenClipText & " " & strRst1 & " (" & Trim(strRst2) & ")" & vbTab
                    blnMicFg(j) = True
                End If
                If strRst1 Like "S*" Then
                    RstForeColor(j + 2, RstRow) = vbRed '&H7477EF
                    RstForeColor(2, RstRow) = &H404080  '&H7477EF
                Else
                    RstForeColor(j + 2, RstRow) = &H747474
                End If
                
            End If
        Next
        For j = MicroCnt + 1 To 5
            SenClipText = SenClipText & vbTab
        Next
        SenClipText = SenClipText & vbTab & vbTab & AntiNm & vbCrLf
        SortFg = True
    Next
    SortEndRow = RstRow
   
NoData:
'   tmpRs.RsClose
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   Set AntiList = Nothing
   
End Sub


'% 검체리마크 Text
'% -. LAB034 : CdIndex-'C404'의 Text1
Public Function ReadRemark(ByVal RmkCd As String) As String
   
   Dim tmpSql As String
   Dim tmpRs As Recordset
   
   Remark = ""
   ReadRemark = ""
   tmpSql = objRstSql.SqlGetRemark(RmkCd)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
'   Set tmpRs = OpenRecordSet(tmpSql)
   
   If tmpRs.EOF Then GoTo NoData
   
   ReadRemark = Trim("" & tmpRs.Fields("Remark").Value)
   Remark = "<< Remark >>" & vbCrLf & Trim("" & tmpRs.Fields("Remark").Value) & vbCrLf & vbCrLf
   CommentFg = True

NoData:
'   tmpRs.RsClose
   Set tmpRs = Nothing
End Function

'% FootNote Text
'% -. LAB304 : RstTxt
Public Function ReadFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer) As String
   
    Dim i As Integer
    Dim tmpSql As String
    Dim tmpRs As Recordset
   
    FootNote = ""
    ReadFootNote = ""
    tmpSql = objRstSql.SqlGetFootNote(pWorkArea, pAccDt, pAccSeq)
    Set tmpRs = New Recordset
    tmpRs.Open tmpSql, DBConn
'    Set tmpRs = OpenRecordSet(tmpSql)
    If tmpRs.EOF Then GoTo NoData
   
    FootNote = "<< Foot Note >>" & vbCrLf
    ReadFootNote = Trim("" & tmpRs.Fields("FootNote").Value)
    While (Not tmpRs.EOF)
        FootNote = FootNote & Trim("" & tmpRs.Fields("FootNote").Value) & vbCrLf
        tmpRs.MoveNext
    Wend
    CommentFg = True
   
NoData:
'    tmpRs.RsClose
    Set tmpRs = Nothing
   
End Function


'% 결과내역 Spread에 Display될 Clip Text를 작성한다.
'% -. 검사명(Short), 결과, 단위, High/Low, Delta/Panic, 기준치, 검사명(Long), 최근결과, 최근결과확인일, 속성코드
Public Function OneRow(ByVal Row As Integer)

    Dim strNotice As String
    Dim strRow As String
    
    OneRow = ""
    If Row > UBound(MyResult) Then Exit Function
    With MyResult(Row)
        OneRow = .TestShortNm & vbTab & .RstCd & vbTab & .RstUnit & vbTab & .HLDiv & vbTab & .DPDiv & vbTab & _
                 .RefCd & vbTab & .TestLongNm & vbTab & .LastRst & vbTab & .LastVfyDtTm & vbTab & _
                 .AttrCd & vbTab & .MfyFg & vbTab & .TestCd & vbCrLf
        If Trim(.ClinicalNotice) <> "" Then
            Call AddRow(1)
            RstForeColor(4, RstRow) = DCM_LightRed
            OneRow = OneRow & vbTab & vbTab & "☞ Clinical Notice " & Space(42) & vbCrLf
            strNotice = Replace(.ClinicalNotice, vbCr, "")
            strRow = medShift(strNotice, vbLf)
            While strRow <> ""
                Call AddRow(1)
                RstForeColor(4, RstRow) = &H747474
                OneRow = OneRow & vbTab & vbTab & strRow & Space(60 - Len(Mid(strRow, 1, 60))) & vbCrLf
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
   
'   Set tmpRs = OpenRecordSet(objCommSql.SqlCommonCode(T_LAB032, LC3_WorkArea, WorkAreaCd))
   If tmpRs.EOF Then
      GetWorkAreaNm = ""
   Else
      GetWorkAreaNm = Trim("" & tmpRs.Fields("Field1").Value)
   End If
'   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function

'% 텍스트결과 내역
'% -. LAB303
Public Function GetRstText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

   Dim i As Integer
   Dim tmpRs As Recordset
   
   Set tmpRs = New Recordset
   tmpRs.Open objRstSql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
'   Set tmpRs = OpenRecordSet(objRstSql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd))
   If tmpRs.EOF Then
      GetRstText = ""
   Else
      GetRstText = Trim("" & tmpRs.Fields("RstTxt").Value)
   End If
'   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function


'% 텍스트결과 수정내역
'% -. LAB305
Public Function GetSuppText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

   Dim i As Integer
   Dim tmpRs As Recordset
   
   Set tmpRs = New Recordset
   tmpRs.Open objRstSql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
'   Set tmpRs = OpenRecordSet(objRstSql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd))
   If tmpRs.EOF Then
      GetSuppText = ""
   Else
      GetSuppText = ""
      While (Not tmpRs.EOF)
         GetSuppText = GetSuppText & Trim("" & tmpRs.Fields("RstTxt").Value) & vbCrLf & _
                              "♧ 수정자 : " & Trim("" & tmpRs.Fields("EmpNm").Value) & "   " & Trim("" & tmpRs.Fields("MfyDtTm").Value) & vbCrLf
         tmpRs.MoveNext
      Wend
   End If
'   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function

'% 기타검사 Value Result ( LAB352 )
Public Sub GetRstValues(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                                    ByVal pTestCd As String, ByVal MfyFg As String)
   
   Dim tmpRs As Recordset
   Dim tmpSql As String
   Dim i As Integer
   
   tmpSql = objRstSql.SqlEValueRst(pWorkArea, pAccDt, pAccSeq, pTestCd, MfyFg)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
'   Set tmpRs = OpenRecordSet(tmpSql)
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
'   tmpRs.RsClose
   Set tmpRs = Nothing

End Sub

'% 직원 성명
Public Function GetEmpNm(ByVal EmpId As String)

   Dim tmpRs As Recordset
   Dim SqlStmt As String
   
   GetEmpNm = ""
   If EmpId = "" Then Exit Function
   
   SqlStmt = objCommSql.SqlLAB015Read(EmpId, 0)
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
'   Set tmpRs = OpenRecordSet(SqlStmt)
   If tmpRs.EOF Then
      GetEmpNm = ""
   Else
      GetEmpNm = Trim("" & tmpRs.Fields("EmpNm").Value)
   End If
'   tmpRs.RsClose
   Set tmpRs = Nothing

End Function

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
'   tmpRs.Open SqlStmt, DBConn
'
''   Set tmpRs = OpenRecordSet(SqlStmt)
'   If tmpRs.EOF Then
'      GetDoctNm = ""
'   Else
'      GetDoctNm = Trim("" & tmpRs.Fields("EmpNm").Value)
'   End If
''   tmpRs.RsClose
'   Set tmpRs = Nothing
'
'End Function

'% 부서이름
'Public Function GetDeptNm(ByVal DeptCd As String)
'
'    Dim tmpRs As Recordset
'    Dim SqlStmt As String
'
'    GetDeptNm = ""
'    If DeptCd = "" Then Exit Function
'
'    SqlStmt = objCommSql.SqlHIS003CodeList(DeptCd)
'    Set tmpRs = New Recordset
'    tmpRs.Open SqlStmt, DBConn
''    Set tmpRs = OpenRecordSet(SqlStmt)
'
''    If Not ObjLISComCode.DeptCd.Exists(DeptCd) Then
''        GetDeptNm = ""
''    Else
''        ObjLISComCode.DeptCd.KeyChange (DeptCd)
''        GetDeptNm = ObjLISComCode.DeptCd.Fields("DeptNm")
''    End If
'    If tmpRs.EOF Then
'        GetDeptNm = ""
'    Else
'        GetDeptNm = "" & tmpRs.Fields("DeptNm").Value
'    End If
'
''    tmpRs.RsClose
'    Set tmpRs = Nothing
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
'    SuppText = ""
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

    Dim SqlStmt As String
    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    Dim tmpRs As New Recordset
    Dim i As Long
   
    pWorkArea = medGetP(pLabNo, 1, "-")
    pAccDt = medGetP(pLabNo, 2, "-")
    pAccSeq = medGetP(pLabNo, 3, "-")
   
    If Mid(pAccDt, 1, 1) = "9" Then
        pAccDt = "19" & pAccDt
    Else
        pAccDt = "20" & pAccDt
    End If
    SqlStmt = objRstSql.SqlGetRelTest(pWorkArea, pAccDt, pAccSeq)
    tmpRs.Open SqlStmt, DBConn
'    Set tmpRs = OpenRecordSet(SqlStmt)
   
    objCombo.Clear
    If tmpRs.EOF Then
        objCombo.AddItem "< 없음 >"
        GoTo NoData
    End If
   
    With tmpRs
        For i = 1 To .RecordCount
            objCombo.AddItem "" & .Fields("AbbrNm5").Value & "   " & "" & .Fields("RstCd").Value & "   " & _
                             "" & .Fields("RstUnit").Value & "   " & "" & .Fields("VfyDtTm").Value & "   " & _
                             "by " & "" & .Fields("VfyNm").Value
            .MoveNext
        Next
    End With
   
NoData:
'    tmpRs.RsClose
    Set tmpRs = Nothing
   
End Sub

Public Function GetTAT(ByVal pTestCd As String, ByVal pSpcCd As String, ByVal pErFg As String) As String

    Dim tmpRs As New Recordset
    
'    Set tmpRs = OpenRecordSet(objRstSql.SqlGetTAT(pTestCd, pSpcCd))
    tmpRs.Open objRstSql.SqlGetTAT(pTestCd, pSpcCd), DBConn
    
    GetTAT = ""
    If Not tmpRs.EOF Then
        If Trim(pErFg) = "" Then '일반
            GetTAT = medGetP("" & tmpRs.Fields("TATS").Value, 1, ";")
        Else
            GetTAT = medGetP("" & tmpRs.Fields("TATS").Value, 2, ";")
        End If
    End If
'    tmpRs.RsClose
    Set tmpRs = Nothing
    
End Function


Public Function GetBuildNoForTAT(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Long

    Dim tmpRs As New Recordset
    
'    Set tmpRs = OpenRecordSet(objRstSql.SqlGetBuildNoForTAT(pWorkArea, pAccDt, pAccSeq))
    tmpRs.Open objRstSql.SqlGetBuildNoForTAT(pWorkArea, pAccDt, pAccSeq), DBConn
    If Not tmpRs.EOF Then
        GetBuildNoForTAT = Val("" & tmpRs.Fields("BldNo").Value)
'        tmpRs.RsClose
    Else
        GetBuildNoForTAT = 1
    End If
    Set tmpRs = Nothing
    
End Function


Public Function CollectQuery(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean

    Dim tmpRs As New Recordset
    
'    Set tmpRs = OpenRecordSet(objRstSql.SqlGetCollectData(pWorkArea, pAccDt, pAccSeq))
    tmpRs.Open objRstSql.SqlGetCollectData(pWorkArea, pAccDt, pAccSeq), DBConn
    
    If Not tmpRs.EOF Then
        With tmpRs
            ptid = "" & .Fields("PtId").Value
            Sex = "" & .Fields("Sex").Value
            AgeDay = "" & .Fields("AgeDay").Value
            ColId = "" & .Fields("ColId").Value
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
            BedID = "" & .Fields("BedId").Value
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
    
    strYesterday = Format(DateAdd("d", -1, GetSystemDate), CS_DateDbFormat)
    SqlStmt = objRstSql.SqlArchiveVerifiedList(pDeptCd, strYesterday)
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    Archive_Data = True
    
    Exit Function
    
Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    Archive_Data = False
    MsgBox Err.Description, vbExclamation
End Function


Public Function Get_VerifiedList(ByVal pDeptCd As String, ByVal pDoneFg As String, _
                                  ByRef tblPtList As Object)

    Dim i As Integer
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim DivRow As Long
    
    If pDeptCd = "" Then pDeptCd = "ER"
    
    DivRow = 0
'    QueryFg = True
    
'    Set tmpRs = OpenRecordSet(objRstSql.SqlGetVerifiedList(pDeptCd, pDoneFg))

    Set tmpRs = New Recordset
    tmpRs.Open objRstSql.SqlGetVerifiedList(pDeptCd, pDoneFg), DBConn
    
    With tblPtList
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = tmpRs.RecordCount
        For i = 1 To tmpRs.RecordCount
            .Row = i
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
            
            If DivRow = 0 And "" & tmpRs.Fields("VfyDt").Value = Format(Now, CS_DateDbFormat) Then DivRow = i
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
          
'    tmpRs.RsClose
    Set tmpRs = Nothing
'    QueryFg = False
    
End Function

Public Function Get_PtListByDoctor(ByVal pDoctID As String, ByRef tblPtList As Object)

    Dim i As Integer
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim DivRow As Long
    
    DivRow = 0
'    QueryFg = True
    
'    Set tmpRs = OpenRecordSet(objRstSql.SqlPtListByDoctor(pDoctId))
    tmpRs.Open objRstSql.SqlPtListByDoctor(pDoctID), DBConn
    
    With tblPtList
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = tmpRs.RecordCount
        For i = 1 To tmpRs.RecordCount
            .Row = i
            .Col = 1: .CellType = CellTypeStaticText
                      .Value = ""
            .Col = 2: .Value = "" & tmpRs.Fields("PtId").Value
            .Col = 3: .Value = "" & tmpRs.Fields("PtNm").Value
            .Col = 4: .Value = "" & tmpRs.Fields("WardId").Value & "-" & tmpRs.Fields("RoomId").Value
            .Col = 6: .Value = Format(Now, CS_DateDbFormat)
            .Col = 8: .Value = "" & tmpRs.Fields("OrdDiv").Value
            
            tmpRs.MoveNext
        Next
        
    End With
          
'    tmpRs.RsClose
    Set tmpRs = Nothing
'    QueryFg = False
    
End Function


Public Function UpdateVerifiedList(ByVal pDeptCd As String, ByVal pVerifyDt As String, _
                                   ByVal pPtId As String, ByVal pMfyFg As String, _
                                   ByVal pEmpId As String) As Boolean

    Dim SqlStmt As String
            
    On Error GoTo Err_Trap
    SqlStmt = objRstSql.SqlUpdateVerifiedList(pDeptCd, pVerifyDt, pPtId, pMfyFg, pEmpId)
    Call DBConn.BeginTrans
    Call DBConn.Execute(SqlStmt)
    Call DBConn.CommitTrans
    UpdateVerifiedList = True
    Exit Function
    
Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    UpdateVerifiedList = False
    MsgBox Err.Description, vbExclamation

End Function

'Public Function UpdateVerifiedAPS(ByVal pDeptCd As String, ByVal pVerifyDt As String, _
'                                  ByVal pPtId As String, ByVal pMfyFg As String, _
'                                  ByVal pEmpId As String) As Boolean
'
''    Dim SqlStmt As String
''
''    On Error GoTo Err_Trap
''    SqlStmt = objRstSql.SqlUpdateVerifiedAPS(pDeptCd, pVerifyDt, pPtId, pMfyFg, pEmpId)
''    Call dbconn.BeginTrans
''    Call dbconn.Execute(SqlStmt)
''    Call dbconn.CommitTrans
'    UpdateVerifiedAPS = True
'    Exit Function
'
'Err_Trap:
''    Call Error_Routine
'    DBConn.RollbackTrans
'    UpdateVerifiedAPS = False
'    MsgBox Err.Description, vbExclamation
'
'End Function

Public Function SetStartDt(ByVal pDeptCd As String) As Date
    
    Dim Rs As Recordset
    Dim sDate As Date
    
'    Set rs = OpenRecordSet(objRstSql.SqlGetStartDate(pDeptCd))
    Set Rs = New Recordset
    Rs.Open objRstSql.SqlGetStartDate(pDeptCd), DBConn
    
    If Not Rs.EOF Then
        SetStartDt = DateAdd("m", Val("" & Rs.Fields("Field1").Value) * (-1), Now)
'        Rs.RsClose
    Else
        SetStartDt = DateAdd("m", -1, Now)
    End If
    Set Rs = Nothing
    
End Function


Public Sub LoadWorkArea(ByRef cboWorkArea As Object)

    Dim i%
    Dim SqlStmt As String
    Dim tmpRs As Recordset

'    Set tmpRs = OpenRecordSet(objRstSql.SqlGetWorkAreaList)
    Set tmpRs = New Recordset
    tmpRs.Open objRstSql.SqlGetWorkAreaList, DBConn

    If tmpRs.EOF = True Then ' record가 존재하지 않을 경우
'        tmpRs.RsClose
        Set tmpRs = Nothing
        Exit Sub
    End If

    cboWorkArea.Clear
    With tmpRs
        .MoveFirst
        For i = 1 To .RecordCount
            cboWorkArea.AddItem "" & .Fields("cdval1").Value & "  " & "" & .Fields("field1").Value
            .MoveNext
        Next i
    End With
    cboWorkArea.ListIndex = 0
    
'    tmpRs.RsClose
    Set tmpRs = Nothing

End Sub

Public Sub LoadItem(ByRef lstList As Object, ByVal pWorkArea As String, ByVal pSpcCd As String)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim tmpStr As String
    Dim i%
    
    '상세항목 제외...
'    Set Rs = OpenRecordSet(objRstSql.SqlGetItemList(pWorkArea, pSpcCd))
    Set Rs = New Recordset
    Rs.Open objRstSql.SqlGetItemList(pWorkArea, pSpcCd), DBConn
    
    If Rs.EOF Then GoTo NoData
    
    lstList.Clear
    For i = 1 To Rs.RecordCount
        tmpStr = "" & Rs.Fields("TestCd").Value & Space(9)
        lstList.AddItem Mid(tmpStr, 1, 10) & _
                        "" & Rs.Fields("TestNm").Value & Space(30) & vbTab & _
                        "" & Rs.Fields("TestDiv").Value & vbTab & _
                        "" & Rs.Fields("WorkArea").Value & vbTab & _
                        "" & Rs.Fields("PanelFg").Value & vbTab & _
                        pSpcCd & vbTab & _
                        "" & Rs.Fields("SpcNm").Value
        Rs.MoveNext
    Next i
    
NoData:
'    Rs.RsClose
    Set Rs = Nothing
    
End Sub



