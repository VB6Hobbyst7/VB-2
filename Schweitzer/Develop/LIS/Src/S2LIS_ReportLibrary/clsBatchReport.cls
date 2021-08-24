VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatchReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPtId   As String        '환자ID
Private mvarPtNm   As String        '환자명
Private mvarPtSex  As String        '나이
Private mvarPtAge  As String        '성별
Private mvarFromDt As String        '기간(시작)
Private mvarToDt   As String        '기간(끝)

Private mvarDept   As String        '진료과
Private mvarDeptNm As String        '진료과
Private mvarWard   As String        '병동
Private mvarWardId As String
Private mvarDoct   As String        '의뢰의사
Private mvarVfyNm  As String        '보고자
Private mvarVfyDt  As String        '보고일
Private mvarMdfDt  As String        '수정일
Private mvarICD    As String        '임상진단

Private mvarRouding As Boolean      '회진여부
Private mvarReprint As Boolean      '재발행여부
Private mvarSpecial As Boolean      '기타검사여부
Private mvarBatchReprint As Boolean '임상병리 일괄재발행여부
Private mvarRealPrinter As Boolean  '즉시출력여부

Private mvarErrorFg As Boolean

Private iPWACoHeRow As Integer
Private fWhich As Object
Private iPageWidth As Integer
Private iPageHeight As Integer
Private iCurY As Integer
Private DataExist As Boolean
Private sLastDt As String
Private sLastTm As String
Private iRecordCount As Integer

Private HeadString1 As String
Private HeadString2 As String
Private HeadString3 As String

Private PageNumber  As Integer

Private strImgPath As String
Private strSlidePath As String
Private picESign As Object


'Const iCm = 567
Const iCm = 10
Const iLineHeight = 10

Private iposOrdDt%, iposSpcNm%, iposTestNm%, iposRstCd%, iposLastRst%, _
        iposUnit%, iposHL, iposDP%, iposRefRng%, iposText%
Private CurWorkArea As String, CurAccDt As String, CurAccSeq As String
Private CurColDtTm As String
Private CurRcvDtTm As String
Private CurVfyDtTm As String

Private mResults As Collection
Private mvarCount As Integer '지역 복사
Private FirstFg As Boolean

Private Type SenResult
   RstCd As String
   Row As String
   ForeColor As Long
End Type

Private lngErrFile As Long
'핵의학 처리를 위해서 추가
Private mvarRiPrint As String

'-- 성바오로병원 결과지 출력
Private mvarRptDt As String
Private mvarRptTm As String
Private mvarRptID As String

Private mvarTestNm  As String
Private mvarTestCd As String

Private mvarImage As Boolean


Public Property Let RiPrint(ByVal vNewValue As String)
    mvarRiPrint = vNewValue
End Property
Public Property Get ptid() As String
    ptid = mvarPtId
End Property

Public Property Let ptid(ByVal vNewValue As String)
    mvarPtId = vNewValue
End Property

Public Property Get PtNm() As String
    PtNm = mvarPtNm
End Property

Public Property Let PtNm(ByVal vNewValue As String)
    mvarPtNm = vNewValue
End Property

Public Property Get PtSex() As String
    PtSex = mvarPtSex
End Property

Public Property Let PtSex(ByVal vNewValue As String)
    mvarPtSex = vNewValue
End Property

Public Property Get PtAge() As String
    PtAge = mvarPtAge
End Property

Public Property Let PtAge(ByVal vNewValue As String)
    mvarPtAge = vNewValue
End Property

Public Property Let TestCd(ByVal vNewValue As String)
    mvarTestCd = vNewValue
End Property

Public Property Get FromDt() As String
    FromDt = mvarFromDt
End Property

Public Property Let FromDt(ByVal vNewValue As String)
    mvarFromDt = vNewValue
End Property

Public Property Get TODT() As String
    TODT = mvarToDt
End Property

Public Property Let TODT(ByVal vNewValue As String)
    mvarToDt = vNewValue
End Property
Public Property Let Dept(ByVal vData As String)
    mvarDept = vData
End Property
Public Property Get Dept() As String
    Dept = mvarDept
End Property
Public Property Let DeptNm(ByVal vData As String)
    mvarDeptNm = vData
End Property
Public Property Get DeptNm() As String
    DeptNm = mvarDeptNm
End Property

Public Property Let Ward(ByVal vData As String)
    mvarWard = vData
End Property

Public Property Get Ward() As String
    Ward = mvarWard
End Property

Public Property Let WardID(ByVal vData As String)
    mvarWardId = vData
End Property

Public Property Get WardID() As String
    WardID = mvarWardId
End Property

Public Property Let Doct(ByVal vData As String)
    mvarDoct = vData
End Property

Public Property Get Doct() As String
    Doct = mvarDoct
End Property
Public Property Let VfyNM(ByVal vData As String)
    mvarVfyNm = vData
End Property
Public Property Get VfyNM() As String
    VfyNM = mvarVfyNm
End Property
Public Property Let VfyDt(ByVal vData As String)
    mvarVfyDt = vData
End Property
Public Property Get VfyDt() As String
    VfyDt = mvarVfyDt
End Property
Public Property Let MdfDt(ByVal vData As String)
    mvarMdfDt = vData
End Property
Public Property Get MdfDt() As String
    MdfDt = mvarMdfDt
End Property
Public Property Let ICD(ByVal vData As String)
    mvarICD = vData
End Property
Public Property Get ICD() As String
    ICD = mvarICD
End Property

Public Property Let Rouding(ByVal vData As Boolean)
    mvarRouding = vData
End Property
Public Property Get Rouding() As Boolean
    Rouding = mvarRouding
End Property
Public Property Let Reprint(ByVal vData As Boolean)
    mvarReprint = vData
End Property
Public Property Get Reprint() As Boolean
    Reprint = mvarReprint
End Property
Public Property Let BatchReprint(ByVal vData As Boolean)
    mvarBatchReprint = vData
End Property
Public Property Get BatchReprint() As Boolean
    BatchReprint = mvarBatchReprint
End Property
Public Property Let Special(ByVal vData As Boolean)
    mvarSpecial = vData
End Property
Public Property Get Special() As Boolean
    Special = mvarSpecial
End Property

Public Property Get ErrorFg() As Boolean
    ErrorFg = mvarErrorFg
End Property


Public Property Get Count() As Integer
    Count = mResults.Count
End Property

Public Property Get Item(ByVal KeyIndex As Variant) As clsResults
    Set Item = mResults(KeyIndex)
End Property

Public Sub ReportForImageOnePatient(ByVal pPtId As String, ByVal pVfyDt As String, _
                               ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByVal pTestDiv As String, _
                               ByVal pSlidePath As String, ByVal pImage As Object, _
                               ByVal pImgPath As String, ByVal pESign As Object, Optional ByVal pCnt As Long = 1)
    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String, SvWorkArea As String
    Dim objRptSql As New clsLISSqlReview
    Dim MyData As New clsResults
    Dim blnFirst As Boolean
    Dim strDPDiv As String
    
    Dim objDic As New clsDictionary
    
    mvarRealPrinter = False
    
    objDic.Clear
    objDic.FieldInialize "seq", "workarea,accdt,accseq,testcd,testdiv"
    objDic.Sort = False
    
    'Dim me As New clsResultReport
   
    mvarImage = True
   
    MouseRunning
'    pProgBar.Value = 20
    
    mvarErrorFg = False
    
    strSlidePath = pSlidePath
    strImgPath = pImgPath
    Set picESign = pESign
    
    '처방일/접수일 기준
    objRptSql.RiPrint = mvarRiPrint
    objRptSql.TestCd = mvarTestCd
    SqlStmt = objRptSql.SqlQueryImageForReport(pPtId, pVfyDt, pWorkArea, pAccDt, pAccSeq, pTestDiv)
    
    'Query
    blnFirst = False
    tmpRs.Open SqlStmt, DBConn
    
    SvWorkArea = "": SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    '====================================================
    'sybase 구동시 tmpRs.EOF가 데이타가 있음에도 불구하고
    '항상 TRUE로 됨
    '====================================================
'    If dbconn.Whatsthis = dbconn.ThisIsOracle Then
'        If tmpRs.EOF Then      '성모자애
'            MouseDefault
'            Call PrtImageBody
'            Set tmpRs = Nothing
'            Exit Sub
'        End If
'    Else
        If tmpRs.EOF Then  'Or tmpRs.EOF Then '길병원,울산동강병운
            MouseDefault
            Call PrtImageBody
            Set tmpRs = Nothing
            Exit Sub
        End If
'    End If
    
    Dim ii As Integer
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    Do Until tmpRs.EOF
        
        ii = ii + 1
'        If pProgBar.Value >= pProgBar.Max Then pProgBar.Max = pProgBar.Max + 50
'        pProgBar.Value = pProgBar.Value + 1
    
        DoEvents
        
        If Not mvarReprint Then
            If mvarWard <> "" Then  '병동
'                If Trim("" & tmpRs.fields("WardId").value) <> medGetP(mvarWard, 1, "-") Then GoTo Skip
            Else
               ' If Trim("" & tmpRs.fields("DeptCd").value) <> mvarDept Then GoTo Skip
            End If
        End If
        
        If Not blnFirst Then
'            Set objDept = Nothing
'            Set objDept = New clsBasisData
            
            strDept = GetDeptNm(tmpRs.Fields("deptcd").Value & "")
'            Set objDept = Nothing
            
            If mvarWard <> "" And strDept <> "" Then
                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
                mvarDeptNm = strDept
                blnFirst = True
            End If

'            If mvarWard <> "" And ObjLISComCode.DeptCd.Exists(Trim("" & tmpRs.Fields("DeptCd").Value)) Then
'                Call ObjLISComCode.DeptCd.KeyChange(Trim("" & tmpRs.Fields("DeptCd").Value))
'                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
'                mvarDeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                blnFirst = True
'            End If
        End If
        
        mvarWardId = Trim("" & tmpRs.Fields("wardid").Value)
        
        '**------------------------------------------------------------------------------
        '** 결과구분이 Alternative이면서 결과가 없을경우는 출력에서 제외한다. 2001.6.20
        '**------------------------------------------------------------------------------
        
        If Trim("" & tmpRs.Fields("RstDiv").Value) = "A" And Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
            
            GoTo Skip
        End If
        
        If P_MicSensiReport = False Then    'MIC Sensi인경우 출력에서 제외
            If Trim("" & tmpRs.Fields("TestCd").Value) Like P_MicSensiTestCd & "*" Then
                
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
                
                
                GoTo Skip
            End If
        End If
        '**------------------------------------------------------------------------------
        
    
'        If pLastDt < Trim("" & tmpRs.fields("VfyDt").value) Then
'            pLastDt = Trim("" & tmpRs.fields("VfyDt").value)
'            pLastTm = Trim("" & tmpRs.fields("VfyTm").value)
'        ElseIf pLastTm < Trim("" & tmpRs.fields("Vfytm").value) Then
'            pLastTm = Trim("" & tmpRs.fields("VfyTm").value)
'        End If
            
        
        If SvWorkArea <> Trim("" & tmpRs.Fields("WorkArea").Value) Then
            SvKeyDt = ""
            SvSpcNm = ""
            SvWorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        End If
        
        'If SvKeyDt <> Trim("" & tmpRs.fields("KeyDate").value) Then
        If SvKeyDt <> Trim("" & tmpRs.Fields("ColDtTm").Value) Then
            MyData.OrdDt = Trim("" & tmpRs.Fields("ColDtTm").Value)     '처방일
            MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            SvKeyDt = Trim("" & tmpRs.Fields("ColDtTm").Value)
            SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
        Else
            MyData.OrdDt = "":
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                MyData.SpcNm = ""
            End If
        End If
        
        'MyData.OrdDate = Trim("" & tmpRs.fields("KeyDate").value)       '처방일
        MyData.OrdDate = Trim("" & tmpRs.Fields("ColDtTm").Value)       '처방일
        MyData.SpcName = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
        
        '-- 검사명
        mvarTestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        If Trim("" & tmpRs.Fields("TestDiv").Value) = enTestDiv.TST_SpeTest Then
            MyData.TestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        Else
            tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestShortNm").Value), 1, 24)
            If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                
                MyData.TestNm = tmpTestNm & " " & String(24 - Len(tmpTestNm), ".")
            Else
                MyData.TestNm = Space(4) & tmpTestNm & " " & String(24 - Len("  " & tmpTestNm), ".")
            End If
        End If
                    
        '-- 결과명(코드일 경우..)
        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
            MyData.RstCd = "미확"
        Else
            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                MyData.RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            Else
                MyData.RstCd = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
            End If
            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                MyData.RstCd = "Growth"
            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                MyData.RstCd = Space(3)
            End If
        End If
                    
        '-- 결과단위
        MyData.RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
        
        '-- High / Low
        MyData.HLDiv = ""
        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then MyData.HLDiv = HLDIV_HIGH_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then MyData.HLDiv = HLDIV_LOW_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then MyData.HLDiv = "*" 'Abnormal_Flag
        End If
        
        '## 5.0.8: 이상대(2005-05-30)
        '   - Alpha결과 참고치의 판정을 "Abnormal"로 표시하도록 수정
        '-- Delta/Panic
        strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
        strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
        MyData.DPDiv = strDPDiv
        
        MyData.WorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        MyData.AccDt = Trim("" & tmpRs.Fields("AccDt").Value):
        MyData.AccSeq = Trim("" & tmpRs.Fields("AccSeq").Value)
        MyData.LastRst = Trim("" & tmpRs.Fields("LastRst").Value)           '-- 최근결과
        MyData.TestCd = Trim("" & tmpRs.Fields("TestCd").Value):            '-- 검사코드
        MyData.SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)               '-- 검체코드
        MyData.VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)               '-- 보고일
        MyData.VfyId = Trim("" & tmpRs.Fields("VfyId").Value)               '-- 보고자
        MyData.TestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)           '-- 검사구분
        MyData.TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)               '-- 소견결과여부
        MyData.FootNoteFg = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
        MyData.RmkCd = Trim("" & tmpRs.Fields("RmkCd").Value)               '-- Remark 코드
        MyData.SenFg = Trim("" & tmpRs.Fields("SenFg").Value)               '-- 감수성 여부
        MyData.PtSex = Trim("" & tmpRs.Fields("Sex").Value)                 '-- Sex
        MyData.AgeDay = Trim("" & tmpRs.Fields("AgeDay").Value)             '-- AgeDay
        MyData.RefRng = CS_QuestionMark
        MyData.MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)               '-- 수정SEQ
        MyData.StsCd = Trim("" & tmpRs.Fields("StsCd").Value)               '-- 상태
        MyData.VfyDtTm = Trim("" & tmpRs.Fields("VfyDtTm").Value)
        MyData.RcvDtTm = Trim("" & tmpRs.Fields("RcvDtTm").Value)
        MyData.ColDtTm = Trim("" & tmpRs.Fields("ColDtTm").Value)
        MyData.ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
        'Debug.Print tmpRs.fields("testcd")
        
        '-- 성바오로병원 결과지 출력(By M.G.Choi 2002.09.24)...FeedBack
'        mvarRptDt = Trim("" & tmpRs.fields("rptdt").value)
'        mvarRptTm = Trim("" & tmpRs.fields("rpttm").value)
'        mvarRptID = Trim("" & tmpRs.fields("rptid").value)
        
        Call Me.Add(MyData)
Skip:
        tmpRs.MoveNext
    Loop
    
    objDic.Sort = True
      
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    
'    pProgBar.Value = pProgBar.Max - 10
'    pProgBar.Msg = "임상 참고치를 검색하고 있습니다.."
    DoEvents
    
    '결과지출력을 않해도 되는항목들.....UPDATE
    If objDic.RecordCount > 0 Then
        objDic.MoveFirst
        Do Until objDic.EOF
            Call PrintUpdate(objDic.Fields("Workarea"), objDic.Fields("accdt"), objDic.Fields("accseq"), objDic.Fields("testcd"), objDic.Fields("testdiv"))
            objDic.MoveNext
        Loop
    End If

    If Me.Count = 0 Then
        
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    For i = 1 To Me.Count
        
 '       If pProgBar.Value < pProgBar.Max Then pProgBar.Value = pProgBar.Value + 1
        '참고치 검색
        If Me.Item(i).RefRng <> CS_QuestionMark Then GoTo RefSkip
        
        tmpSex = Me.Item(i).PtSex
        tmpAgeDay = Me.Item(i).AgeDay
        tmpTestCd = Me.Item(i).TestCd
        tmpSpcCd = Me.Item(i).SpcCd
        tmpVfyDt = Me.Item(i).VfyDt
        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
     
        SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open SqlStmt, DBConn
        
        If tmpRs1.EOF Then
            '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
            '--> 거의 Both로 등록됨.
            SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
        End If
        If tmpRs1.EOF Then
            tmpRefCd = Space(5)
        Else
            tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
            tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
            tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
            If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
               tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        Set tmpRs1 = Nothing
        For j = i To Me.Count
            If Me.Item(j).TestCd = tmpTestCd Then Me.Item(j).RefRng = tmpRefCd
        Next
     
        DoEvents
        
RefSkip:
    Next
      
ExitPos:
'    pProgBar.Value = pProgBar.Max
    DoEvents
      
Nodata:
    
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
    mvarRealPrinter = False
    
    lngErrFile = FreeFile
    Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_ERROR.log" For Append As lngErrFile
    
    For ii = 1 To pCnt
        PageNumber = 0
        Call Print_Report
    Next ii
    
    Close #lngErrFile
                               
End Sub


Public Sub ReportForOnePatient(ByVal pPtId As String, ByVal pFromDate As String, _
                               ByVal pToDate As String, ByVal pTestDiv As String, _
                               ByVal pImgPath As String, ByVal pESign As Object, _
                               ByRef pProgBar As Object, ByRef pLastDt As String, _
                               ByRef pLastTm As String)

    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String, SvWorkArea As String
    Dim objRptSql As New clsLISSqlReview
    Dim MyData As New clsResults
    Dim blnFirst As Boolean
    Dim strDPDiv As String
    
    
    Dim objDic As New clsDictionary
    
    mvarRealPrinter = False
    
    objDic.Clear
    objDic.FieldInialize "seq", "workarea,accdt,accseq,testcd,testdiv"
    objDic.Sort = False
    
    MouseRunning
    
    mvarErrorFg = False
    
    strImgPath = pImgPath
    Set picESign = pESign
    
    '처방일/접수일 기준
    objRptSql.RiPrint = mvarRiPrint
    SqlStmt = objRptSql.SqlQueryForReport(pPtId, pFromDate, pToDate, pTestDiv, mvarReprint Or mvarRouding, _
                                          mvarBatchReprint, gUsingInWardMenu)
    
    'Query
    blnFirst = False
    tmpRs.Open SqlStmt, DBConn
    
    SvWorkArea = "": SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    '====================================================
    'sybase 구동시 tmpRs.EOF가 데이타가 있음에도 불구하고
    '항상 TRUE로 됨
    '====================================================
'    If dbconn.Whatsthis = dbconn.ThisIsOracle Then
        If tmpRs.EOF Then      '성모자애
            MouseDefault
            Set tmpRs = Nothing
            Exit Sub
        End If
'    Else
'        If ColCnt = 0 Then 'Or tmpRs.EOF Then '길병원,울산동강병운
'            MouseDefault
'            Set tmpRs = Nothing
'            Exit Sub
'        End If
'    End If
    
    Dim ii As Integer
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    Do Until tmpRs.EOF
        
        ii = ii + 1
'        If pProgBar.Value >= pProgBar.Max Then pProgBar.Max = pProgBar.Max + 50
'        pProgBar.Value = pProgBar.Value + 1
    
        DoEvents
        
        If Not mvarReprint Then
            If mvarWard <> "" Then  '병동
'                If Trim("" & tmpRs.fields("WardId").value) <> medGetP(mvarWard, 1, "-") Then GoTo Skip
            Else
               ' If Trim("" & tmpRs.fields("DeptCd").value) <> mvarDept Then GoTo Skip
            End If
        End If
        
        If Not blnFirst Then
'            Set objDept = Nothing
'            Set objDept = New clsBasisData
            strDept = GetDeptNm(tmpRs.Fields("deptcd").Value & "")
'            Set objDept = Nothing
            
            If mvarWard <> "" And strDept <> "" Then
                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
                mvarDeptNm = strDept
                blnFirst = True
            End If
            
'            If mvarWard <> "" And ObjLISComCode.DeptCd.Exists(Trim("" & tmpRs.Fields("DeptCd").Value)) Then
'                Call ObjLISComCode.DeptCd.KeyChange(Trim("" & tmpRs.Fields("DeptCd").Value))
'                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
'                mvarDeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                blnFirst = True
'            End If
        End If
        
        mvarWardId = Trim("" & tmpRs.Fields("wardid").Value)
        
        '**------------------------------------------------------------------------------
        '** 결과구분이 Alternative이면서 결과가 없을경우는 출력에서 제외한다. 2001.6.20
        '**------------------------------------------------------------------------------
        
        If Trim("" & tmpRs.Fields("RstDiv").Value) = "A" And Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
            
            GoTo Skip
        End If
        
        If P_MicSensiReport = False Then    'MIC Sensi인경우 출력에서 제외
            If Trim("" & tmpRs.Fields("TestCd").Value) Like P_MicSensiTestCd & "*" Then
                
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
                
                
                GoTo Skip
            End If
        End If
        '**------------------------------------------------------------------------------
        
    
        If pLastDt < Trim("" & tmpRs.Fields("VfyDt").Value) Then
            pLastDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        ElseIf pLastTm < Trim("" & tmpRs.Fields("Vfytm").Value) Then
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        End If
        
'2003/10/08 수정전 원본소스
'        If SvWorkArea <> Trim("" & tmpRs.fields("WorkArea").value) Then
'            SvKeyDt = ""
'            SvSpcNm = ""
'            SvWorkArea = Trim("" & tmpRs.fields("WorkArea").value)
'        End If

'Modify By Legends 2003/10/08
'같은 workarea여도 상태가 다른 경우 채혈일시를 담아주기위해서 수정했음
        Dim svStsCd As String

        If SvWorkArea <> Trim("" & tmpRs.Fields("WorkArea").Value) Then
            SvKeyDt = ""
            SvSpcNm = ""
            SvWorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
            svStsCd = Trim("" & tmpRs.Fields("stscd").Value)
        ElseIf SvWorkArea = Trim("" & tmpRs.Fields("workarea").Value) Then
            If svStsCd <> Trim("" & tmpRs.Fields("stscd").Value) Then
                SvKeyDt = ""
                SvSpcNm = ""
                svStsCd = Trim("" & tmpRs.Fields("stscd").Value)
            End If
        End If
        
        'If SvKeyDt <> Trim("" & tmpRs.fields("KeyDate").value) Then
        If SvKeyDt <> Trim("" & tmpRs.Fields("ColDtTm").Value) Then
            MyData.OrdDt = Trim("" & tmpRs.Fields("ColDtTm").Value)     '채혈일
            MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            SvKeyDt = Trim("" & tmpRs.Fields("ColDtTm").Value)
            SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
        Else
            MyData.OrdDt = "":
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                MyData.SpcNm = ""
            End If
        End If
        
        'MyData.OrdDate = Trim("" & tmpRs.fields("KeyDate").value)       '처방일
        MyData.OrdDate = Trim("" & tmpRs.Fields("ColDtTm").Value)       '처방일
        MyData.SpcName = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
        
        '-- 검사명
        If Trim("" & tmpRs.Fields("TestDiv").Value) = enTestDiv.TST_SpeTest Then
            MyData.TestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        Else
            tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestShortNm").Value), 1, 24)
            If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                
                MyData.TestNm = tmpTestNm & " " & String(24 - Len(tmpTestNm), ".")
            Else
                MyData.TestNm = Space(4) & tmpTestNm & " " & String(24 - Len("  " & tmpTestNm), ".")
            End If
        End If
                    
        '-- 결과명(코드일 경우..)
        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
            MyData.RstCd = "미확"
        Else
            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                MyData.RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            Else
                MyData.RstCd = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
            End If
            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                MyData.RstCd = "Growth"
            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                MyData.RstCd = Space(3)
            End If
        End If
                    
        '-- 결과단위
        MyData.RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
        
        '-- High / Low
        MyData.HLDiv = ""
        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then MyData.HLDiv = HLDIV_HIGH_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then MyData.HLDiv = HLDIV_LOW_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then MyData.HLDiv = "*" 'Abnormal_Flag
        End If
        
        '## 5.0.8: 이상대(2005-05-30)
        '   - Alpha결과 참고치의 판정을 "Abnormal"로 표시하도록 수정
        '-- Delta/Panic
        strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
        strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
        MyData.DPDiv = strDPDiv
        
        MyData.WorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        MyData.AccDt = Trim("" & tmpRs.Fields("AccDt").Value):
        MyData.AccSeq = Trim("" & tmpRs.Fields("AccSeq").Value)
        MyData.LastRst = Trim("" & tmpRs.Fields("LastRst").Value)           '-- 최근결과
        MyData.TestCd = Trim("" & tmpRs.Fields("TestCd").Value):            '-- 검사코드
        MyData.SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)               '-- 검체코드
        MyData.VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)               '-- 보고일
        MyData.VfyId = Trim("" & tmpRs.Fields("VfyId").Value)               '-- 보고자
        MyData.TestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)           '-- 검사구분
        MyData.TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)               '-- 소견결과여부
        MyData.FootNoteFg = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
        MyData.RmkCd = Trim("" & tmpRs.Fields("RmkCd").Value)               '-- Remark 코드
        MyData.SenFg = Trim("" & tmpRs.Fields("SenFg").Value)               '-- 감수성 여부
        MyData.PtSex = Trim("" & tmpRs.Fields("Sex").Value)                 '-- Sex
        MyData.AgeDay = Trim("" & tmpRs.Fields("AgeDay").Value)             '-- AgeDay
        MyData.RefRng = CS_QuestionMark
        MyData.MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)               '-- 수정SEQ
        MyData.StsCd = Trim("" & tmpRs.Fields("StsCd").Value)               '-- 상태
        MyData.VfyDtTm = Trim("" & tmpRs.Fields("VfyDtTm").Value)
        MyData.RcvDtTm = Trim("" & tmpRs.Fields("RcvDtTm").Value)
        MyData.ColDtTm = Trim("" & tmpRs.Fields("ColDtTm").Value)
        MyData.ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
        MyData.DetailFg = Trim("" & tmpRs.Fields("detailfg").Value)
        'Debug.Print tmpRs.fields("testcd")
        
        '-- 성바오로병원 결과지 출력(By M.G.Choi 2002.09.24)...FeedBack
'        mvarRptDt = Trim("" & tmpRs.fields("rptdt").value)
'        mvarRptTm = Trim("" & tmpRs.fields("rpttm").value)
'        mvarRptID = Trim("" & tmpRs.fields("rptid").value)
        
        Call Me.Add(MyData)
Skip:
        tmpRs.MoveNext
    Loop
    objDic.Sort = True
      
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    
'    pProgBar.Value = pProgBar.Max - 10
'    pProgBar.Msg = "임상 참고치를 검색하고 있습니다.."
    DoEvents
    
    '결과지출력을 않해도 되는항목들.....UPDATE
    If objDic.RecordCount > 0 Then
        objDic.MoveFirst
        Do Until objDic.EOF
            Call PrintUpdate(objDic.Fields("Workarea"), objDic.Fields("accdt"), objDic.Fields("accseq"), objDic.Fields("testcd"), objDic.Fields("testdiv"))
            objDic.MoveNext
        Loop
    End If

    If Me.Count = 0 Then
        
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    For i = 1 To Me.Count
        
 '       If pProgBar.Value < pProgBar.Max Then pProgBar.Value = pProgBar.Value + 1
        '참고치 검색
        If Me.Item(i).RefRng <> CS_QuestionMark Then GoTo RefSkip
        
        tmpSex = Me.Item(i).PtSex
        tmpAgeDay = Me.Item(i).AgeDay
        tmpTestCd = Me.Item(i).TestCd
        tmpSpcCd = Me.Item(i).SpcCd
        tmpVfyDt = Me.Item(i).VfyDt
        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
     
        SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open SqlStmt, DBConn
        
        If tmpRs1.EOF Then
            '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
            '--> 거의 Both로 등록됨.
            SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
        End If
        If tmpRs1.EOF Then
            tmpRefCd = Space(5)
        Else
            tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
            tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
            tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
            If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
               tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        Set tmpRs1 = Nothing
        For j = i To Me.Count
            If Me.Item(j).TestCd = tmpTestCd Then Me.Item(j).RefRng = tmpRefCd
        Next
     
        DoEvents
        
RefSkip:
    Next
      
ExitPos:
'    pProgBar.Value = pProgBar.Max
    DoEvents
      
Nodata:
    
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
    mvarRealPrinter = False
    
    lngErrFile = FreeFile
    Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_ERROR.log" For Append As lngErrFile
    
    Call Print_Report
    
    Close #lngErrFile

End Sub

'== 예수병원 추가루틴 =======================================================================
Public Sub ReportForOnePatient_Special(ByVal pPtId As String, ByVal pFromDate As String, _
                               ByVal pToDate As String, ByVal pTestDiv As String, _
                               ByVal pImgPath As String, ByVal pESign As Object, _
                               ByRef pProgBar As Object, ByRef pLastDt As String, _
                               ByRef pLastTm As String)

    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String, SvWorkArea As String
    Dim objRptSql As New clsLISSqlReview
    Dim MyData As New clsResults
    Dim blnFirst As Boolean
    Dim strDPDiv As String
    
    
    Dim objDic As New clsDictionary
    
    mvarRealPrinter = False
    
    objDic.Clear
    objDic.FieldInialize "seq", "workarea,accdt,accseq,testcd,testdiv"
    objDic.Sort = False
    
    MouseRunning
    
    mvarErrorFg = False
    
    strImgPath = pImgPath
    Set picESign = pESign
    
    '처방일/접수일 기준
    objRptSql.RiPrint = mvarRiPrint
    SqlStmt = objRptSql.SqlQueryForReport_Special(pPtId, pFromDate, pToDate, pTestDiv, mvarReprint Or mvarRouding, _
                                          mvarBatchReprint, gUsingInWardMenu)
    
    'Query
    blnFirst = False
    tmpRs.Open SqlStmt, DBConn
    
    SvWorkArea = "": SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    If tmpRs.EOF Then
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    Dim ii As Integer
    Dim strDept As String
    
    Do Until tmpRs.EOF
        
        ii = ii + 1
    
        DoEvents
                
        If Not blnFirst Then
            strDept = GetDeptNm(tmpRs.Fields("deptcd").Value & "")
            
            If mvarWard <> "" And strDept <> "" Then
                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
                mvarDeptNm = strDept
                blnFirst = True
            End If
            
        End If
        
        mvarWardId = Trim("" & tmpRs.Fields("wardid").Value)
        
        '**------------------------------------------------------------------------------
        '** 결과구분이 Alternative이면서 결과가 없을경우는 출력에서 제외한다. 2001.6.20
        '**------------------------------------------------------------------------------
        
        If Trim("" & tmpRs.Fields("RstDiv").Value) = "A" And Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
            
            GoTo Skip
        End If
        
        If P_MicSensiReport = False Then    'MIC Sensi인경우 출력에서 제외
            If Trim("" & tmpRs.Fields("TestCd").Value) Like P_MicSensiTestCd & "*" Then
                
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
                
                
                GoTo Skip
            End If
        End If
        '**------------------------------------------------------------------------------
        
    
        If pLastDt < Trim("" & tmpRs.Fields("VfyDt").Value) Then
            pLastDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        ElseIf pLastTm < Trim("" & tmpRs.Fields("Vfytm").Value) Then
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        End If
        
        Dim svStsCd As String

        If SvWorkArea <> Trim("" & tmpRs.Fields("WorkArea").Value) Then
            SvKeyDt = ""
            SvSpcNm = ""
            SvWorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
            svStsCd = Trim("" & tmpRs.Fields("stscd").Value)
        ElseIf SvWorkArea = Trim("" & tmpRs.Fields("workarea").Value) Then
            If svStsCd <> Trim("" & tmpRs.Fields("stscd").Value) Then
                SvKeyDt = ""
                SvSpcNm = ""
                svStsCd = Trim("" & tmpRs.Fields("stscd").Value)
            End If
        End If
        
        'If SvKeyDt <> Trim("" & tmpRs.fields("KeyDate").value) Then
        If SvKeyDt <> Trim("" & tmpRs.Fields("ColDtTm").Value) Then
            MyData.OrdDt = Trim("" & tmpRs.Fields("ColDtTm").Value)     '채혈일
            MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            SvKeyDt = Trim("" & tmpRs.Fields("ColDtTm").Value)
            SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
        Else
            MyData.OrdDt = "":
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                MyData.SpcNm = ""
            End If
        End If
        
        'MyData.OrdDate = Trim("" & tmpRs.fields("KeyDate").value)       '처방일
        MyData.OrdDate = Trim("" & tmpRs.Fields("ColDtTm").Value)       '처방일
        MyData.SpcName = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
        
        '-- 검사명
        If Trim("" & tmpRs.Fields("TestDiv").Value) = enTestDiv.TST_SpeTest Then
            MyData.TestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        Else
            tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestShortNm").Value), 1, 24)
            If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                
                MyData.TestNm = tmpTestNm & " " & String(24 - Len(tmpTestNm), ".")
            Else
                MyData.TestNm = Space(4) & tmpTestNm & " " & String(24 - Len("  " & tmpTestNm), ".")
            End If
        End If
                    
        '-- 결과명(코드일 경우..)
        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
            MyData.RstCd = "미확"
        Else
            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                MyData.RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            Else
                MyData.RstCd = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
            End If
            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                MyData.RstCd = "Growth"
            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                MyData.RstCd = Space(3)
            End If
        End If
                    
        '-- 결과단위
        MyData.RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
        
        '-- High / Low
        MyData.HLDiv = ""
        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then MyData.HLDiv = HLDIV_HIGH_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then MyData.HLDiv = HLDIV_LOW_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then MyData.HLDiv = "*" 'Abnormal_Flag
        End If
        
        '## 5.0.8: 이상대(2005-05-30)
        '   - Alpha결과 참고치의 판정을 "Abnormal"로 표시하도록 수정
        '-- Delta/Panic
        strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
        strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
        MyData.DPDiv = strDPDiv
        
        MyData.WorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        MyData.AccDt = Trim("" & tmpRs.Fields("AccDt").Value):
        MyData.AccSeq = Trim("" & tmpRs.Fields("AccSeq").Value)
        MyData.LastRst = Trim("" & tmpRs.Fields("LastRst").Value)           '-- 최근결과
        MyData.TestCd = Trim("" & tmpRs.Fields("TestCd").Value):            '-- 검사코드
        MyData.SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)               '-- 검체코드
        MyData.VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)               '-- 보고일
        MyData.VfyId = Trim("" & tmpRs.Fields("VfyId").Value)               '-- 보고자
        MyData.TestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)           '-- 검사구분
        MyData.TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)               '-- 소견결과여부
        MyData.FootNoteFg = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
        MyData.RmkCd = Trim("" & tmpRs.Fields("RmkCd").Value)               '-- Remark 코드
        MyData.SenFg = Trim("" & tmpRs.Fields("SenFg").Value)               '-- 감수성 여부
        MyData.PtSex = Trim("" & tmpRs.Fields("Sex").Value)                 '-- Sex
        MyData.AgeDay = Trim("" & tmpRs.Fields("AgeDay").Value)             '-- AgeDay
        MyData.RefRng = CS_QuestionMark
        MyData.MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)               '-- 수정SEQ
        MyData.StsCd = Trim("" & tmpRs.Fields("StsCd").Value)               '-- 상태
        MyData.VfyDtTm = Trim("" & tmpRs.Fields("VfyDtTm").Value)
        MyData.RcvDtTm = Trim("" & tmpRs.Fields("RcvDtTm").Value)
        MyData.ColDtTm = Trim("" & tmpRs.Fields("ColDtTm").Value)
        MyData.ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
        MyData.DetailFg = Trim("" & tmpRs.Fields("detailfg").Value)
        
        Call Me.Add(MyData)
Skip:
        tmpRs.MoveNext
    Loop
    objDic.Sort = True
      
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    
    DoEvents
    
    '결과지출력을 않해도 되는항목들.....UPDATE
    If objDic.RecordCount > 0 Then
        objDic.MoveFirst
        Do Until objDic.EOF
            Call PrintUpdate(objDic.Fields("Workarea"), objDic.Fields("accdt"), objDic.Fields("accseq"), objDic.Fields("testcd"), objDic.Fields("testdiv"))
            objDic.MoveNext
        Loop
    End If

    If Me.Count = 0 Then
        
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    For i = 1 To Me.Count
        
        '참고치 검색
        If Me.Item(i).RefRng <> CS_QuestionMark Then GoTo RefSkip
        
        tmpSex = Me.Item(i).PtSex
        tmpAgeDay = Me.Item(i).AgeDay
        tmpTestCd = Me.Item(i).TestCd
        tmpSpcCd = Me.Item(i).SpcCd
        tmpVfyDt = Me.Item(i).VfyDt
        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
     
        SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open SqlStmt, DBConn
        
        If tmpRs1.EOF Then
            '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
            '--> 거의 Both로 등록됨.
            SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
        End If
        If tmpRs1.EOF Then
            tmpRefCd = Space(5)
        Else
            tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
            tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
            tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
            If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
               tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        Set tmpRs1 = Nothing
        For j = i To Me.Count
            If Me.Item(j).TestCd = tmpTestCd Then Me.Item(j).RefRng = tmpRefCd
        Next
     
        DoEvents
        
RefSkip:
    Next
      
ExitPos:
'    pProgBar.Value = pProgBar.Max
    DoEvents
      
Nodata:
    
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
    mvarRealPrinter = False
    
    lngErrFile = FreeFile
    Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_ERROR.log" For Append As lngErrFile
    
    Call Print_Report
    
    Close #lngErrFile

End Sub
'============================================================================================

Public Sub ReportForOneERPatient(ByVal pPtId As String, ByVal pVfyDt As String, _
                                 ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, _
                                 ByVal pESign As Object, ByRef pLastDt As String, _
                                 ByRef pLastTm As String)

    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String, SvWorkArea As String
    Dim objRptSql As New clsLISSqlReview
    Dim MyData As New clsResults
    Dim blnFirst As Boolean
    Dim strDPDiv As String
    'Dim me As New clsResultReport
   
    
   
    MouseRunning
'    pProgBar.Value = 20
    
    mvarErrorFg = False
    
    strImgPath = ""
    Set picESign = pESign
    
    '처방일/접수일 기준
    objRptSql.RiPrint = mvarRiPrint
    SqlStmt = objRptSql.SqlQueryERForReport(pPtId, pVfyDt, pWorkArea, pAccDt, pAccSeq)
    
    'Query
    blnFirst = False
    tmpRs.Open SqlStmt, DBConn
    
    SvWorkArea = "": SvKeyDt = "": SvSpcNm = ""
    
'    DoEvents
   
    '====================================================
    'sybase 구동시 tmpRs.EOF가 데이타가 있음에도 불구하고
    '항상 TRUE로 됨
    '====================================================
'    If dbconn.Whatsthis = dbconn.ThisIsOracle Then
        If tmpRs.EOF Then      '성모자애
            MouseDefault
            Set tmpRs = Nothing
            Exit Sub
        End If
'    Else
'        If ColCnt = 0 Then 'Or tmpRs.EOF Then '길병원,울산동강병운
'            MouseDefault
'            Set tmpRs = Nothing
'            Exit Sub
'        End If
'    End If
    
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    Do Until tmpRs.EOF
     
     
'        If pProgBar.Value >= pProgBar.Max Then pProgBar.Max = pProgBar.Max + 50
'        pProgBar.Value = pProgBar.Value + 1
    
'        DoEvents
        
        If Not mvarReprint Then
            If mvarWard <> "" Then  '병동
'                If Trim("" & tmpRs.fields("WardId").value) <> medGetP(mvarWard, 1, "-") Then GoTo Skip
            Else
               ' If Trim("" & tmpRs.fields("DeptCd").value) <> mvarDept Then GoTo Skip
            End If
        End If
        
        If Not blnFirst Then
'            Set objDept = Nothing
'            Set objDept = New clsBasisData
            strDept = GetDeptNm(tmpRs.Fields("deptcd").Value & "")
'            Set objDept = Nothing
            
            If mvarWard <> "" And strDept <> "" Then
                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
                mvarDeptNm = strDept
                blnFirst = True
            End If
            
'            If mvarWard <> "" And ObjLISComCode.DeptCd.Exists(Trim("" & tmpRs.Fields("DeptCd").Value)) Then
'                Call ObjLISComCode.DeptCd.KeyChange(Trim("" & tmpRs.Fields("DeptCd").Value))
'                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
'                mvarDeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                blnFirst = True
'            End If
        End If
        '**------------------------------------------------------------------------------
        '** 결과구분이 Alternative이면서 결과가 없을경우는 출력에서 제외한다. 2001.6.20
        '**------------------------------------------------------------------------------
        If Trim("" & tmpRs.Fields("RstDiv").Value) = "A" And _
           Trim("" & tmpRs.Fields("RstCd").Value) = "" Then GoTo Skip
        If P_MicSensiReport = False Then    'MIC Sensi인경우 출력에서 제외
            If Trim("" & tmpRs.Fields("TestCd").Value) Like P_MicSensiTestCd & "*" Then GoTo Skip
        End If
        '**------------------------------------------------------------------------------
        
    
        If pLastDt < Trim("" & tmpRs.Fields("VfyDt").Value) Then
            pLastDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        ElseIf pLastTm < Trim("" & tmpRs.Fields("Vfytm").Value) Then
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        End If
            
        
        If SvWorkArea <> Trim("" & tmpRs.Fields("WorkArea").Value) Then
            SvKeyDt = ""
            SvSpcNm = ""
            SvWorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        End If
        
        'If SvKeyDt <> Trim("" & tmpRs.fields("KeyDate").value) Then
        If SvKeyDt <> Trim("" & tmpRs.Fields("ColDtTm").Value) Then
            MyData.OrdDt = Trim("" & tmpRs.Fields("ColDtTm").Value)     '처방일
            MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            SvKeyDt = Trim("" & tmpRs.Fields("ColDtTm").Value)
            SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
        Else
            MyData.OrdDt = "":
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                MyData.SpcNm = ""
            End If
        End If
        
        'MyData.OrdDate = Trim("" & tmpRs.fields("KeyDate").value)       '처방일
        MyData.OrdDate = Trim("" & tmpRs.Fields("ColDtTm").Value)       '처방일
        MyData.SpcName = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
        
        '-- 검사명
        If Trim("" & tmpRs.Fields("TestDiv").Value) = enTestDiv.TST_SpeTest Then
            MyData.TestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        Else
            tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestShortNm").Value), 1, 24)
            If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                
                MyData.TestNm = tmpTestNm & " " & String(24 - Len(tmpTestNm), ".")
            Else
                MyData.TestNm = Space(4) & tmpTestNm & " " & String(24 - Len("  " & tmpTestNm), ".")
            End If
        End If
                    
        '-- 결과명(코드일 경우..)
        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
            MyData.RstCd = "미확"
        Else
            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                MyData.RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            Else
                MyData.RstCd = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
            End If
            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                MyData.RstCd = "Growth"
            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                MyData.RstCd = Space(3)
            End If
        End If
                    
        '-- 결과단위
        MyData.RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
        
        '-- High / Low
        MyData.HLDiv = ""
        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then MyData.HLDiv = HLDIV_HIGH_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then MyData.HLDiv = HLDIV_LOW_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then MyData.HLDiv = "*" 'Abnormal_Flag
        End If
        
        '## 5.0.8: 이상대(2005-05-30)
        '   - Alpha결과 참고치의 판정을 "Abnormal"로 표시하도록 수정
        '-- Delta/Panic
        strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
        strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
        MyData.DPDiv = strDPDiv
        
        MyData.WorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        MyData.AccDt = Trim("" & tmpRs.Fields("AccDt").Value)
        MyData.AccSeq = Trim("" & tmpRs.Fields("AccSeq").Value)
        MyData.LastRst = Trim("" & tmpRs.Fields("LastRst").Value)           '-- 최근결과
        MyData.TestCd = Trim("" & tmpRs.Fields("TestCd").Value)             '-- 검사코드
        MyData.SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)               '-- 검체코드
        MyData.VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)               '-- 보고일
        MyData.VfyId = Trim("" & tmpRs.Fields("VfyId").Value)               '-- 보고자
        MyData.TestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)           '-- 검사구분
        MyData.TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)               '-- 소견결과여부
        MyData.FootNoteFg = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
        MyData.RmkCd = Trim("" & tmpRs.Fields("RmkCd").Value)               '-- Remark 코드
        MyData.SenFg = Trim("" & tmpRs.Fields("SenFg").Value)               '-- 감수성 여부
        MyData.PtSex = Trim("" & tmpRs.Fields("Sex").Value)                 '-- Sex
        MyData.AgeDay = Trim("" & tmpRs.Fields("AgeDay").Value)             '-- AgeDay
        MyData.RefRng = CS_QuestionMark
        MyData.MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)               '-- 수정SEQ
        MyData.StsCd = Trim("" & tmpRs.Fields("StsCd").Value)               '-- 상태
        MyData.VfyDtTm = Trim("" & tmpRs.Fields("VfyDtTm").Value)
        MyData.RcvDtTm = Trim("" & tmpRs.Fields("RcvDtTm").Value)
        MyData.ColDtTm = Trim("" & tmpRs.Fields("ColDtTm").Value)
        MyData.DetailFg = Trim("" & tmpRs.Fields("detailfg").Value)
'        MyData.ClinicalNotice = Trim("" & tmpRs.fields("Notice").value)
        'Debug.Print tmpRs.fields("testcd")
        
        Call Me.Add(MyData)
Skip:
        tmpRs.MoveNext
    Loop
      
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    
'    pProgBar.Value = pProgBar.Max - 10
'    pProgBar.Msg = "임상 참고치를 검색하고 있습니다.."
'    DoEvents
    
    If Me.Count = 0 Then
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    For i = 1 To Me.Count
        
 '       If pProgBar.Value < pProgBar.Max Then pProgBar.Value = pProgBar.Value + 1
        '참고치 검색
        If Me.Item(i).RefRng <> CS_QuestionMark Then GoTo RefSkip
        
        tmpSex = Me.Item(i).PtSex
        tmpAgeDay = Me.Item(i).AgeDay
        tmpTestCd = Me.Item(i).TestCd
        tmpSpcCd = Me.Item(i).SpcCd
        tmpVfyDt = Me.Item(i).VfyDt
        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
     
        SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open SqlStmt, DBConn
        
        If tmpRs1.EOF Then
            '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
            '--> 거의 Both로 등록됨.
            SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
        End If
        If tmpRs1.EOF Then
            tmpRefCd = Space(5)
        Else
            tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
            tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
            tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
            If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
               tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        Set tmpRs1 = Nothing
        For j = i To Me.Count
            If Me.Item(j).TestCd = tmpTestCd Then Me.Item(j).RefRng = tmpRefCd
        Next
     
'        DoEvents
    
RefSkip:
    Next
      
ExitPos:
'    pProgBar.Value = pProgBar.Max
'    DoEvents
      
Nodata:
    MouseDefault
'    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
    Dim strMsg      As VbMsgBoxResult
    
    If mvarDept = "ER" Then Call Select_Print("응급실")
    If mvarDept = "24" Then Call Select_Print("수술실")
    
    
    If strMsg = vbNo Then Exit Sub
    
    mvarRealPrinter = True
    
    lngErrFile = FreeFile
    Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_ERROR.log" For Append As lngErrFile
    
    Call Print_Report
    
    Close #lngErrFile
    
End Sub

Public Sub MicSensiReport(ByVal pWorkArea As String, ByVal pAccDt As String, _
                          ByVal pAccSeq As String, ByVal pESign As Object)

    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String, SvWorkArea As String
    Dim objRptSql As New clsLISSqlReview
    Dim MyData As New clsResults
    Dim blnFirst As Boolean
    Dim strDPDiv As String
    'Dim me As New clsResultReport
   
    
   
    MouseRunning
'    pProgBar.Value = 20
    
    mvarErrorFg = False
    
    strImgPath = ""
    Set picESign = pESign
    
    '처방일/접수일 기준
    SqlStmt = objRptSql.SqlQueryForMicReport(pWorkArea, pAccDt, pAccSeq)
    
    'Query
    blnFirst = False
    tmpRs.Open SqlStmt, DBConn
    
    'ColCnt = tmpRs.OpenCursor(, SqlStmt)
    
    SvWorkArea = "": SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
'    If ColCnt = 0 Or tmpRs.EOF Then
'        MouseDefault
'        Set tmpRs = Nothing
'        Exit Sub
'    End If
    
    'While (tmpRs.FetchCursor(ColCnt))
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    While (Not tmpRs.EOF)
     
     
'        If pProgBar.Value >= pProgBar.Max Then pProgBar.Max = pProgBar.Max + 50
'        pProgBar.Value = pProgBar.Value + 1
    
        If Not blnFirst Then
'            Set objDept = Nothing
'            Set objDept = New clsBasisData
            strDept = GetDeptNm(tmpRs.Fields("deptcd").Value & "")
'            Set objDept = Nothing
            
            If mvarWard <> "" And strDept <> "" Then
                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
                mvarDeptNm = strDept
                blnFirst = True
            End If
            
'            If mvarWard <> "" And ObjLISComCode.DeptCd.Exists(Trim("" & tmpRs.Fields("DeptCd").Value)) Then
'                Call ObjLISComCode.DeptCd.KeyChange(Trim("" & tmpRs.Fields("DeptCd").Value))
'                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
'                mvarDeptNm = ObjLISComCode.DeptCd.Fields("deptnm").Value
'                blnFirst = True
'            End If
        End If
        '**------------------------------------------------------------------------------
        '** 결과구분이 Alternative이면서 결과가 없을경우는 출력에서 제외한다. 2001.6.20
        '**------------------------------------------------------------------------------
        If Trim("" & tmpRs.Fields("RstDiv").Value) = "A" And _
           Trim("" & tmpRs.Fields("RstCd").Value) = "" Then GoTo Skip
        'MIC Sensi인경우 출력에서 제외
        
        If Trim("" & tmpRs.Fields("TestCd").Value) = "L4062" Then GoTo Skip
        '**------------------------------------------------------------------------------
        
    
        If SvWorkArea <> Trim("" & tmpRs.Fields("WorkArea").Value) Then
            SvKeyDt = ""
            SvSpcNm = ""
            SvWorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        End If
        
        'If SvKeyDt <> Trim("" & tmpRs.fields("KeyDate").value) Then
        If SvKeyDt <> Trim("" & tmpRs.Fields("ColDtTm").Value) Then
            MyData.OrdDt = Trim("" & tmpRs.Fields("ColDtTm").Value)     '처방일
            MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            SvKeyDt = Trim("" & tmpRs.Fields("ColDtTm").Value)
            SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
        Else
            MyData.OrdDt = "":
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                MyData.SpcNm = ""
            End If
        End If
        
        'MyData.OrdDate = Trim("" & tmpRs.fields("KeyDate").value)       '처방일
        MyData.OrdDate = Trim("" & tmpRs.Fields("ColDtTm").Value)       '처방일
        MyData.SpcName = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
        
        '-- 검사명
        If Trim("" & tmpRs.Fields("TestDiv").Value) = TST_SpeTest Then
            MyData.TestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        Else
            tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestShortNm").Value), 1, 24)
            If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                
                MyData.TestNm = tmpTestNm & " " & String(24 - Len(tmpTestNm), ".")
            Else
                MyData.TestNm = Space(4) & tmpTestNm & " " & String(24 - Len("  " & tmpTestNm), ".")
            End If
        End If
                    
        '-- 결과명(코드일 경우..)
        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
            MyData.RstCd = "미확"
        Else
            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                MyData.RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            Else
                MyData.RstCd = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
            End If
            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                MyData.RstCd = "Growth"
            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                MyData.RstCd = Space(3)
            End If
        End If
                    
        '-- 결과단위
        MyData.RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
        
        '-- High / Low
        MyData.HLDiv = ""
        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then MyData.HLDiv = HLDIV_HIGH_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then MyData.HLDiv = HLDIV_LOW_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then MyData.HLDiv = "*" 'Abnormal_Flag
        End If
        
        '## 5.0.8: 이상대(2005-05-30)
        '   - Alpha결과 참고치의 판정을 "Abnormal"로 표시하도록 수정
        '-- Delta/Panic
        strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
        strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
        MyData.DPDiv = strDPDiv
        
        MyData.WorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        MyData.AccDt = Trim("" & tmpRs.Fields("AccDt").Value)
        MyData.AccSeq = Trim("" & tmpRs.Fields("AccSeq").Value)
        MyData.LastRst = Trim("" & tmpRs.Fields("LastRst").Value)           '-- 최근결과
        MyData.TestCd = Trim("" & tmpRs.Fields("TestCd").Value)             '-- 검사코드
        MyData.SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)               '-- 검체코드
        MyData.VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)               '-- 보고일
        MyData.VfyId = Trim("" & tmpRs.Fields("VfyId").Value)               '-- 보고자
        MyData.TestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)           '-- 검사구분
        MyData.TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)               '-- 소견결과여부
        MyData.FootNoteFg = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
        MyData.RmkCd = Trim("" & tmpRs.Fields("RmkCd").Value)               '-- Remark 코드
        MyData.SenFg = Trim("" & tmpRs.Fields("SenFg").Value)               '-- 감수성 여부
        MyData.PtSex = Trim("" & tmpRs.Fields("Sex").Value)                 '-- Sex
        MyData.AgeDay = Trim("" & tmpRs.Fields("AgeDay").Value)             '-- AgeDay
        MyData.RefRng = ""
        MyData.MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)               '-- 수정SEQ
        MyData.StsCd = Trim("" & tmpRs.Fields("StsCd").Value)               '-- 상태
        MyData.VfyDtTm = Trim("" & tmpRs.Fields("VfyDtTm").Value)
        MyData.RcvDtTm = Trim("" & tmpRs.Fields("RcvDtTm").Value)
        MyData.ColDtTm = Trim("" & tmpRs.Fields("ColDtTm").Value)
        MyData.ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
        
        Call Me.Add(MyData)
Skip:
        tmpRs.MoveNext
        
    Wend
      
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    
'    pProgBar.Value = pProgBar.Max - 10
'    pProgBar.Msg = "임상 참고치를 검색하고 있습니다.."
    DoEvents
    
    If Me.Count = 0 Then
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
          
ExitPos:
'    pProgBar.Value = pProgBar.Max
    DoEvents
      
Nodata:
    
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    
    lngErrFile = FreeFile
    Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_ERROR.log" For Append As lngErrFile
    
    Call Print_Report
    
    Close #lngErrFile

End Sub

Private Sub ErrorWrite(ByVal strErrMsg As String)
    mvarErrorFg = True
    Print #lngErrFile, strErrMsg
End Sub

Private Sub Class_Initialize()
    
    Set mResults = New Collection
    
    Printer.Font = "굴림체"
    Printer.FontSize = 10
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    Twidth = Printer.ScaleWidth
    
    Select Case Printer.PaperSize
        Case 9
            PrtLeft = 10
            'LineSpace = 6
            LineSpace = 5
        Case 7
            PrtLeft = 0
            LineSpace = 4
        Case Else
            PrtLeft = 0
            LineSpace = 4
    End Select
    LastLineYpos = Printer.ScaleHeight - iCm             '마지막라인Y위치
    
    mvarCount = 0
    'iposOrdDt = iCm * 1
    'iposOrdDt = iCm * 0
    iposOrdDt = PrtLeft * 0.1 * iCm
    iposSpcNm = iposOrdDt + iCm * 2
    iposTestNm = iposSpcNm + 2.5 * iCm
    iposRstCd = iposTestNm + iCm * 4
    iposLastRst = iposRstCd + iCm * 2
    iposUnit = iposLastRst + iCm * 1.8
    iposHL = iposUnit + iCm * 1.4
    iposDP = iposHL + iCm * 1.5 + 2
    iposRefRng = iposDP + iCm * 1.2
    'iposText = iposRefRng + 7.8 * iCm
    
End Sub


Public Function Add(ByVal clsData As clsResults, Optional ByVal clsKey As Variant)
   
   Dim newItem As New clsResults
   
   With newItem
        .OrdDt = clsData.OrdDt
        .SpcNm = clsData.SpcNm
        .TestNm = clsData.TestNm
        .VfyDt = clsData.VfyDt
        .VfyId = clsData.VfyId
        .RstCd = clsData.RstCd
        .RstUnit = clsData.RstUnit
        .HLDiv = clsData.HLDiv
        .DPDiv = clsData.DPDiv
        .RefRng = clsData.RefRng
        .TxtFg = clsData.TxtFg
        .SenFg = clsData.SenFg
        .FootNoteFg = clsData.FootNoteFg
        .RmkCd = clsData.RmkCd
        .DcFg = clsData.DcFg
        .WorkArea = clsData.WorkArea
        .AccDt = clsData.AccDt
        .AccSeq = clsData.AccSeq
        .LastRst = clsData.LastRst
        .TestCd = clsData.TestCd
        .SpcCd = clsData.SpcCd
        .TestDiv = clsData.TestDiv
        .OrdDate = clsData.OrdDate
        .SpcName = clsData.SpcName
        .MfyFg = clsData.MfyFg
        .PtSex = clsData.PtSex
        .AgeDay = clsData.AgeDay
        .StsCd = clsData.StsCd
        .ColDtTm = clsData.ColDtTm
        .RcvDtTm = clsData.RcvDtTm
        .VfyDtTm = clsData.VfyDtTm
        .ClinicalNotice = clsData.ClinicalNotice
        .DetailFg = clsData.DetailFg
   End With
   
   mvarCount = mvarCount + 1
   If IsMissing(clsKey) Then clsKey = "Key" & CStr(mvarCount)
   mResults.Add newItem, clsKey
   Set Add = newItem
   
End Function


Public Sub Clear()
   Set mResults = Nothing
   Set mResults = New Collection
   mvarCount = 0
End Sub

Public Sub Print_Report()
    
    'Call InitReport
    Call PrtBody
    
End Sub

Private Sub PrtHeader()
    Dim Header1 As Integer
    Dim Header2 As Integer
    Dim strTmp  As String
    Dim sICSString As String
    
    Header1 = Twidth * (1 / 3) + PrtLeft + iCm / 2
    Header2 = Twidth * (2 / 3) + PrtLeft
    
    lngCurYPos = 10
    Printer.FontSize = 18: Printer.FontBold = True
    
    If gUsingInWardMenu Then
        If mvarRouding Then
            Call Print_Setting("(회진용) 결과보고서", 0, 12, Twidth, "C", "C")
        ElseIf mvarReprint Then
            Call Print_Setting("(재발행) 결과보고서", 0, 12, Twidth, "C", "C")
        Else
            Call Print_Setting("진단검사의학과 검사보고서", 0, 12, Twidth, "C", "C")
        End If
        Printer.FontSize = 10: Printer.FontBold = False
        Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "R", "C")
    Else
        '-- 2002.07.29 By MGChoi 성바오로경우만 임상병리/핵의학 검사보고서로 함.

        Call Print_Setting("진단검사의학과 검사보고서", 0, 12, Twidth, "C", "C")
        Printer.FontSize = 10: Printer.FontBold = False
        Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "R", "C")
    End If
    '감염관리
    sICSString = ICSPatientString(mvarPtId, enICSNum.ResultReport)
    
    Printer.DrawStyle = vbSolid
    Printer.DrawWidth = 3
    Printer.Line (PrtLeft, lngCurYPos - 2)-(Twidth - PrtLeft, lngCurYPos - 2)
    
    Call Print_Setting("등록번호: ", PrtLeft, LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = True
    Call Print_Setting(mvarPtId, PrtLeft + Printer.TextWidth("등록번호: "), LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = False
    
    Call Print_Setting("환 자 명: ", Header1, LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = True
    Call Print_Setting(mvarPtNm & sICSString, Header1 + Printer.TextWidth("환 자 명: "), LineSpace, Twidth, "L", "C", False)
    Printer.FontBold = False
    
    Call Print_Setting("성별/나이: " & mvarPtSex & " / " & mvarPtAge, Header2, LineSpace, Twidth, "L", "C")
    
'    Call Print_Setting("의 뢰 과: " & mvarDept, PrtLeft, LineSpace, Twidth, "L", "C", False)
'    Call Print_Setting("주 치 의 : " & mvarDoct, Header1, LineSpace, Twidth, "L", "C", False)
'    Call Print_Setting("보 고 자 : " & mvarVfyNm, Header2, LineSpace, Twidth, "L", "C")
    
    '============================
    '환자별로 입원내역 긁어오자.
    '============================
    strTmp = GetIpWon(mvarPtId)
    '외래환자
    If strTmp = "" Then
        If mvarWard <> "" Then
            Call Print_Setting("병    동: " & mvarWard & "  " & GetPartName(mvarDept), PrtLeft, LineSpace, Twidth, "L", "C", False)
            'Call Print_Setting("병    동: " & medGetP(strTmp, 3, COL_DIV) & "-" & _
                                                      medGetP(strTmp, 2, COL_DIV) & " " & _
                                                      GetPartName(mvarDept), PrtLeft, LineSpace, Twidth, "L", "C", False)
        Else
            Call Print_Setting("의 뢰 과: " & mvarDept & "  " & GetPartName(mvarDept), PrtLeft, LineSpace, Twidth, "L", "C", False)
        End If
        
    Else
    '입원환자
        Call Print_Setting("병    동: " & medGetP(strTmp, 3, COL_DIV) & "-" & _
                                          medGetP(strTmp, 2, COL_DIV) & " " & _
                                          GetPartName(mvarDept), PrtLeft, LineSpace, Twidth, "L", "C", False)
    End If
    
    If mvarSpecial Then
        Call Print_Setting("접수번호: ", Header1, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting(CurWorkArea & "-" & Mid(CurAccDt, 3) & "-" & CurAccSeq & "   " & mvarDoct, _
                           Header1 + Printer.TextWidth("접수번호: "), LineSpace, Twidth, "L", "C", False)
    Else
        Call Print_Setting("처 방 의: " & mvarDoct, Header1, LineSpace, Twidth, "L", "C", False)
    End If
    
    If mvarSpecial Then
        Call Print_Setting("접수일시 : " & CurRcvDtTm, Header2, LineSpace, Twidth, "L", "C")
    Else
'        Call Print_Setting("보고일자 : " & mvarVfyDt, Header2, LineSpace, Twidth, "L", "C")
        Call Print_Setting("접수일시 : " & CurRcvDtTm, Header2, LineSpace, Twidth, "L", "C")
    End If
    
    Call Print_Setting("임상진단: " & mvarICD, PrtLeft, LineSpace, Twidth, "L", "C", False)
    
    'Call Print_Setting("수정일자 : " & mvarMdfDt, Header2, LineSpace, Twidth, "L", "C")
    
    If mvarSpecial Then
        'Call Print_Setting("검 체 명: " & mResults.Item(1).SpcNm, Header2, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("보고일시 : " & CurVfyDtTm, Header2, LineSpace, Twidth, "L", "C", False)
    Else
        Call Print_Setting("보고일자 : " & mvarVfyDt, Header2, LineSpace, Twidth, "L", "C", False)
    End If
    
    Call Print_Setting("", PrtLeft, LineSpace / 5, Twidth, "R", "C")
    Printer.Line (PrtLeft, lngCurYPos + LineSpace)-(Twidth - PrtLeft, lngCurYPos + LineSpace)
    
    Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")
    
    If mvarSpecial Then
        Call Print_Setting("", PrtLeft, LineSpace / 8, Twidth, "R", "C")
        Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
        Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")
    Else
        Printer.FontBold = True
        Call Print_Setting("채혈일시", iposOrdDt, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("    검체", iposSpcNm, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("        검  사  명", iposTestNm, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("    결과", iposRstCd, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting(" 단위", iposLastRst, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("     HL", iposUnit, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting(" DP", iposHL, LineSpace, Twidth, "L", "C", False)
        Call Print_Setting("      참 고 치", iposDP - iCm, LineSpace, Twidth, "L", "C")
        Printer.FontBold = False
        Printer.DrawWidth = 4
        
        Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
    End If
    Call Print_Setting("", PrtLeft, 2, Twidth, "L", "C")
    FirstFg = True
End Sub

Private Function GetIpWon(ByVal ptid As String) As String
    Dim sSQL As String
    Dim Rs   As Recordset
    

    sSQL = "SELECT a." & F_PTWARDID & " as Wardid ," & _
                 " a." & F_PTROOMID & " as hosilid," & _
                 " b." & F_DEPTNM & " as wardnm " & _
           " FROM " & T_HIS003 & " b," & T_HIS002 & " a" & _
           " WHERE " & _
                 DBW("a." & F_INPTID, ptid, 2) & _
           " AND " & "(" & F_BEDOUTDT2("a") & " is null)" & _
           " AND a." & F_PTWARDID & "=b." & F_DEPTCD
        
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        GetIpWon = Rs.Fields("Wardid").Value & "" & COL_DIV & _
                   Rs.Fields("hosilid").Value & "" & COL_DIV & _
                   Rs.Fields("wardnm").Value & ""
                 
    End If
    
    Set Rs = Nothing
End Function

Private Function GetPartName(ByVal PartCd As String) As String
    Dim sSQL As String
    Dim Rs   As Recordset
    
    PartCd = Trim(medGetP(PartCd, 1, "-"))
    
    
    
    sSQL = "SELECT  " & F_DEPTNM & " as deptnm FROM " & T_HIS003 & " WHERE " & DBW(F_DEPTCD, PartCd, 2)
        
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        GetPartName = Rs.Fields("deptnm").Value & ""
    Else
        GetPartName = PartCd
    End If
    
    Set Rs = Nothing
End Function


Private Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer, Optional Alignment As AlignmentConstants, Optional MaxLen As Integer)

    Printer.CurrentY = Y
    iNextY = Printer.CurrentY + iSpace
    
    Select Case Alignment
    Case vbCenter
       Printer.CurrentX = X + (MaxLen - Printer.TextWidth(str)) \ 2
    Case vbLeftJustify
        Printer.CurrentX = X
    Case vbRightJustify
        Printer.CurrentX = X + MaxLen - Printer.TextWidth(str)
    End Select
        
    If IsMissing(Alignment) Then Printer.CurrentX = X
   
    Printer.Print str

End Sub
Private Sub ChangeLine()
    Call Print_Setting("", iposSpcNm, LineSpace, Twidth, "C", "C")
    
    '추가함
    Call CheckNewPage
End Sub

Private Sub PrtBody()

    
    Dim tmpWorkArea     As String
    Dim tmpAccDt        As String
    Dim tmpAccSeq       As String
    Dim tmpWorkAreaNm   As String
    Dim SaveWorkArea    As String
    Dim strBuffer       As String
    Dim strTxt          As String
    Dim strNotice       As String
    Dim strTmp          As String
    
    Dim i       As Integer
    Dim ii      As Integer
    Dim jj      As Integer
    Dim kk      As Integer
    Dim lngCnt  As Integer

    
    Dim objRichText As RichTextBox
    Dim objimage    As Image
    Dim objESign    As clsLISElectronSign
    
    Dim tmpVfyNm    As String
    
On Error GoTo Err_Trap
    
    Printer.FontSize = 9
    
    tmpWorkArea = ""
    tmpAccDt = ""
    tmpAccSeq = ""
    
    CurWorkArea = mResults.Item(1).WorkArea
    CurAccDt = mResults.Item(1).AccDt
    CurAccSeq = mResults.Item(1).AccSeq
    CurRcvDtTm = mResults.Item(1).RcvDtTm
    CurVfyDtTm = mResults.Item(1).VfyDtTm
                
    Call PrtHeader
    
    Set objRichText = frmControls.rtfTextBox
    Set objimage = frmControls.Image1
    
    If P_ImageSystem = True And mvarImage = True Then
    
'        Call ChangeLine
        Call CheckNewPage
        Printer.FontBold = True
        Printer.FontSize = 12
        Call Print_Setting(mvarTestNm, 0, LineSpace, Twidth - Len(mvarTestNm), "C", "C", True)
'        Call Print_Setting(String(80, Chr(6)), iposOrdDt, LineSpace, Twidth, "L", "C", True)
        Printer.FontBold = False
        Printer.FontSize = 9
        Call CheckNewPage
        
        Call Print_Setting("", PrtLeft, LineSpace / 8, Twidth, "R", "C")
        '추가함
        Call CheckNewPage
        Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
        Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")
        '추가함
        Call CheckNewPage
        
        Call ChangeLine
        Call CheckNewPage
        objimage.Picture = LoadPicture(strSlidePath)
        Printer.PaintPicture objimage.Picture, iposOrdDt + 35, LineSpace * 14, 100, 100
        Call Print_Setting("", iposOrdDt, LineSpace * 20, Twidth, "L", "C", True)
        Call CheckNewPage
    End If
    
    For i = 1 To mResults.Count
        With mResults.Item(i)
            'Debug.Print mResults.Item(i).WorkArea & "," & .TestNm
            
            CurWorkArea = .WorkArea
            CurAccDt = .AccDt
            CurAccSeq = .AccSeq
            CurRcvDtTm = .RcvDtTm
            CurVfyDtTm = .VfyDtTm
                            
            CheckNewPage
            
            '상세 코드일경우 상세 항목들을 모두 출력한후에 소견결과를 출력한다.
'            Debug.Print Trim(Replace(.TestNm, ".", "")) & "----------> " & strBuffer & iCurY
            If strTxt <> "" And .DetailFg = "" Then
                Call Print_Setting("◈ 소견결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
                '추가함
                Call CheckNewPage
                While (strTxt <> "")
                    Call Print_Setting(medShift(strTxt, vbCr), iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 소견결과 :   "), 10, Twidth, "L", "C")
                    CheckNewPage
                Wend
                Call ChangeLine
                strTxt = ""
            End If
            '추가함
            Call CheckNewPage

'2003/10/08 수정전 원본
'            'WorkArea 이름 찍어주기...(기타검사는 제외)
'            If SaveWorkArea <> .WorkArea And Not mvarSpecial Then
'                '결과지 출력 시 중간보고 표시
'                If .StsCd = StsCd_LIS_MidRst Then
'                    tmpWorkAreaNm = GetWorkAreaName(.WorkArea) & "(중간보고)"
'                Else
'                    tmpWorkAreaNm = GetWorkAreaName(.WorkArea)
'                End If
'
'                Printer.FontBold = True
'                Printer.FontSize = 10
'                Call ChangeLine
'                Call CheckNewPage
'                Call Print_Setting("▶ " & tmpWorkAreaNm, iposOrdDt, LineSpace, Twidth, "L", "C", True)
'                Call CheckNewPage
'            '    Call ChangeLine
'                Printer.FontBold = False
'                Printer.FontSize = 9
'                SaveWorkArea = .WorkArea
'            End If

'Modify By Legends 2003/10/08
'-------------------------------------------------
'같은 workarea라도 상태가 바뀌면 학부명을 바까조야 됨.
'---------------------------------------------------
            Dim SaveStsCd As String
            
            'WorkArea 이름 찍어주기...(기타검사는 제외)
            If SaveWorkArea <> .WorkArea And Not mvarSpecial Then
                '결과지 출력 시 중간보고 표시
                If .StsCd = StsCd_LIS_MidRst Then
                    tmpWorkAreaNm = GetWorkAreaName(.WorkArea) & "(중간보고)"
                ElseIf .StsCd = StsCd_LIS_FinRst Then
                    tmpWorkAreaNm = GetWorkAreaName(.WorkArea)
                ElseIf .StsCd = StsCd_LIS_Modify Then
                    tmpWorkAreaNm = GetWorkAreaName(.WorkArea) & "(수정결과)"
                End If

                Call ChangeLine
                Call CheckNewPage
                
                Printer.FontBold = True
                Printer.FontSize = 10
                Call Print_Setting("▶ " & tmpWorkAreaNm, iposOrdDt, LineSpace, Twidth, "L", "C", True)
                Printer.FontBold = False
                Printer.FontSize = 9
                Call CheckNewPage
                
                SaveWorkArea = .WorkArea
                SaveStsCd = .StsCd
            ElseIf SaveWorkArea = .WorkArea And Not mvarSpecial Then
                '결과지 출력 시 중간보고 표시
                If SaveStsCd <> .StsCd Then
                    If .StsCd = StsCd_LIS_MidRst Then
                        tmpWorkAreaNm = GetWorkAreaName(.WorkArea) & "(중간보고)"
                    ElseIf .StsCd = StsCd_LIS_FinRst Then
                        tmpWorkAreaNm = GetWorkAreaName(.WorkArea)
                    ElseIf .StsCd = StsCd_LIS_Modify Then
                        tmpWorkAreaNm = GetWorkAreaName(.WorkArea) & "(수정결과)"
                    End If

                    Call ChangeLine
                    Call CheckNewPage
                    
                    Printer.FontBold = True
                    Printer.FontSize = 10
                    Call Print_Setting("▶ " & tmpWorkAreaNm, iposOrdDt, LineSpace, Twidth, "L", "C", True)
                    Printer.FontBold = False
                    Printer.FontSize = 9
                    Call CheckNewPage
                    
                    SaveStsCd = .StsCd
                End If
                
                SaveWorkArea = .WorkArea
            End If
            
            If FirstFg Then
                .OrdDt = .OrdDate
                .SpcNm = .SpcName
                If mvarSpecial Then
                    Printer.FontBold = True
                    Printer.FontSize = 11
                    'Debug.Print "▶ " & .TestNm & " ◀"
                    Call Print_Setting("▶ " & .TestNm & " ◀", iposOrdDt + iCm, LineSpace, Twidth, "L", "C", True)
                    '추가함
                    Call CheckNewPage

                    Printer.FontBold = False
                    Printer.FontSize = 9
                End If
                FirstFg = False
            End If
            
            Printer.FontSize = 9
            
            If mvarSpecial Then
                Call Print_Setting("", iposDP, LineSpace, Twidth - iposRefRng, "L", "C")
                '추가함
                Call CheckNewPage
            Else
                Call Print_Setting(.OrdDt, iposOrdDt, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                Printer.FontBold = True
                Call Print_Setting(.SpcNm, iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
                Printer.FontBold = False
                '추가함
                Call CheckNewPage
                
'                Dim strleft As String
'
'                Debug.Print mResults.Item(i).TestNm
'                If strleft = "R" Then
'
'                    Call Print_Setting(Space(40), iposUnit + iCm * 0.5 + 5, "-" & LineSpace, Twidth, "L", "C", False)
'                    Call Print_Setting("(" & Trim(.TestNm) & " " & .RstCd & ")", iposUnit + iCm * 0.5 + 5, "-" & LineSpace, Twidth, "L", "C", False)
'                    Call Print_Setting("", iposDP, LineSpace, Twidth - iposRefRng, "L", "C")
'                Else
                
                Call Print_Setting(.TestNm, iposTestNm, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                Call Print_Setting(.RstCd, iposRstCd, LineSpace, iposLastRst - iposRstCd, "C", "C", False)
                '추가함
                Call CheckNewPage
                
                '최근결과 빠진후 ....
                
                Call Print_Setting(.RstUnit, iposLastRst, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                Call Print_Setting(.HLDiv, iposUnit + iCm * 0.5 + 5, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                Call Print_Setting(.DPDiv, iposHL + iCm * 0.8 - 5, LineSpace, Twidth, "l", "C", False)
                '추가함
                Call CheckNewPage
                Call Print_Setting(.RefRng, iposDP, LineSpace, Twidth - iposRefRng, "L", "C")
                '추가함
                Call CheckNewPage
'                End If
                If Trim(.ClinicalNotice) <> "" Then
                    Call Print_Setting("☞ Clinical Notice", iposTestNm, LineSpace, Twidth, "L", "C")
                    '추가함
                    Call CheckNewPage
                    strNotice = Trim(.ClinicalNotice)
                    strNotice = Replace(strNotice, vbCr, "")
                    strTmp = medShift(strNotice, vbLf)
                    While strTmp <> ""
                        Call Print_Setting(strTmp, iposTestNm, LineSpace, Twidth, "L", "C")
                        strTmp = medShift(strNotice, vbLf)
                        CheckNewPage
                    Wend
                End If
                
            End If
            
            CheckNewPage
            
            '일반검사 Text결과 출력
'            If .TxtFg = "1" Or .TxtFg = "2" Then
'                'objRichText.TextRTF = GetRstText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
'                strBuffer = GetRstText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
'                strBuffer = Replace(strBuffer, vbLf, "")
'                Printer.FontBold = True
'                Call Print_Setting("◈ 소견결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
'                Printer.FontBold = False
'                While (strBuffer <> "")
'                    Call Print_Setting(medShift(strBuffer, vbCr), iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 소견결과 :   "), 10, Twidth, "L", "C")
'                    CheckNewPage
'                Wend
'                Call ChangeLine
'            End If

            '텍스트 결과를 출력하기 위해서 추가/수정
            If .TxtFg = "1" Or .TxtFg = "2" Then
                If strTxt <> "" Then
                    strTxt = strTxt & vbCrLf & " ( " & Trim(Replace(.TestNm, ".", "")) & " ) : " & _
                                GetRstText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
                Else
                    If .DetailFg = "" Then
                        strTxt = GetRstText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
                    Else
                        strTxt = " ( " & Trim(Replace(.TestNm, ".", "")) & " ) : " & _
                                    GetRstText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
                    End If
                End If
                strTxt = Replace(strTxt, vbLf, "")
                '단일항목에 텍스트결과가 있을경우는 단일항목밑에 출력
                If .DetailFg = "" Then
                    'Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "소견결과 : ", iCurY, 0)
                    Call Print_Setting("◈ 소견결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
                    '추가함
                    Call CheckNewPage
                    While (strTxt <> "")
                        Call Print_Setting(medShift(strTxt, vbCr), iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 소견결과 :   "), LineSpace, Twidth, "L", "C")
                        CheckNewPage
                    Wend
                    strTxt = ""
                    Call ChangeLine
                End If
            End If
            '상세항목만 처방난경우 마지막에 출력하기 위해 삽입.
            If i = mResults.Count And strTxt <> "" Then
                Call Print_Setting("◈ 소견결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
                '추가함
                Call CheckNewPage
                While (strTxt <> "")
                    Call Print_Setting(medShift(strTxt, vbCr), iposSpcNm + iCm * 0.4 + Printer.TextWidth("◈ 소견결과 :   "), LineSpace, Twidth, "L", "C")
                    CheckNewPage
                Wend
                Call ChangeLine
                strTxt = ""
            End If
            
            '기타검사 Text결과 출력 * 2014-11-10 PSK
            If .TxtFg = "Y" And .TestDiv = enTestDiv.TST_SpeTest Then
                objRichText.TextRTF = GetSpcText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd, .MfyFg)
                strBuffer = objRichText.Text
                strBuffer = vbCr & Replace(strBuffer, vbLf, "")
                Printer.FontBold = True
                'Call Print_Setting("◈ 소견결과 : ", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
                Printer.FontBold = False
                While (strBuffer <> "")
                    Call Print_Setting(medShift(strBuffer, vbCr), iposOrdDt + iCm, LineSpace, Twidth, "L", "C")
                    CheckNewPage
                Wend
                
                Call ChangeLine
                
'                Dim objEmp As clsBasisData
                
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                mvarVfyNm = GetEmpNm(.VfyId)
'                Set objEmp = Nothing
                
'                mvarVfyNm = GetEmpName(.VfyId)

                Call Print_Setting("보고자 : " & mvarVfyNm & " M.D", iposUnit - iCm, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                
                If Trim(strImgPath) <> "" Then
                    Printer.PaintPicture picESign.Picture, iposUnit + Printer.TextWidth("보고자 :" & mvarVfyNm & " M.D") - iCm / 2, Printer.CurrentY - 10, 30, 15
                Else
                    Set objESign = New clsLISElectronSign
                    If objESign.LoadElectronSign(.VfyId, InstallDir & "LIS\") = True Then
                        If objESign.ElectronSignPrintOk = True Then
                            strImgPath = objESign.ElectronSignPath & "\" & objESign.ElectronSignFileName
                            picESign.Picture = LoadPicture(strImgPath)
                            Printer.PaintPicture picESign.Picture, iposUnit + Printer.TextWidth("보고자 :" & mvarVfyNm & " M.D") - iCm / 2, Printer.CurrentY - 10, 30, 15
                        End If
                    End If
                    Set objESign = Nothing
                End If
'                picESign.Picture = LoadPicture()
    
                'Supplemental Report
'                strBuffer = GetSpcText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd, "1")
'                If strBuffer <> "" Then
'                    strBuffer = Replace(strBuffer, vbLf, "")
'                    Printer.FontBold = True
'                    Call Print_Setting("☞ Supplemental Report :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
'                    Printer.FontBold = False
'                    While (strBuffer <> "")
'                        Call Print_Setting(medShift(strBuffer, vbCr), iposSpcNm + iCm * 0.4 + Printer.TextWidth("Supplemental Report :   "), LineSpace, Twidth, "L", "C")
'                        CheckNewPage
'                    Wend
'                End If
                Call ChangeLine
                Call CheckNewPage
                Call ChangeLine
                Call CheckNewPage
            End If
            
'            If .WorkArea = "50" And .AccDt = "200303" And .AccSeq = "8552" Then
'                MsgBox .TestCd
'            End If
            
            '감수성 결과 출력
            If .SenFg = "Y" And .StsCd >= enStsCd.StsCd_LIS_FinRst Then
                strBuffer = GetSenResult_New(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
                strBuffer = Replace(strBuffer, vbLf, "")
                Printer.FontBold = True
                Call Print_Setting("▣ 감수성 결과 :", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                Printer.FontBold = False
                While (strBuffer <> "")
                    Call Print_Setting(medShift(strBuffer, vbCr), iposSpcNm + iCm * 0.4 + Printer.TextWidth("▣ 감수성 결과 :   "), LineSpace, Twidth, "L", "C")
                    CheckNewPage
                Wend
                'Call ChangeLine
                Call ChangeLine
                Call ChangeLine
                
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                mvarVfyNm = GetEmpNm(.VfyId)
'                Set objEmp = Nothing
                
'                mvarVfyNm = GetEmpName(.VfyId)
                
'                Call Print_Setting("보고자 : " & mvarVfyNm & " M.D", iposUnit - iCm / 2, LineSpace, Twidth, "L", "C", False)
                '추가함
                Call CheckNewPage
                If Trim(strImgPath) <> "" Then
                    Printer.PaintPicture picESign.Picture, iposUnit + Printer.TextWidth("보고자 :" & mvarVfyNm & " M.D"), Printer.CurrentY - 10, 30, 15
                    'Printer.PaintPicture picESign.Picture, iposUnit, _
                                         lngCurYPos - 10, 30, 15
                Else

                    Set objESign = New clsLISElectronSign
                    If objESign.LoadElectronSign(.VfyId, InstallDir & "LIS\") = True Then
                        If objESign.ElectronSignPrintOk = True Then
                            strImgPath = objESign.ElectronSignPath & "\" & objESign.ElectronSignFileName
                            picESign.Picture = LoadPicture(strImgPath)
                            Printer.PaintPicture picESign.Picture, iposUnit + Printer.TextWidth("보고자 :" & mvarVfyNm & " M.D"), Printer.CurrentY - 10, 30, 15
                        End If
                    End If
                    Set objESign = Nothing
                End If
                
                Call ChangeLine
                Call ChangeLine
            End If
            
            If i < mResults.Count Then
                tmpWorkArea = mResults.Item(i + 1).WorkArea
                tmpAccDt = mResults.Item(i + 1).AccDt
                tmpAccSeq = mResults.Item(i + 1).AccSeq
            Else
                tmpWorkArea = "End"
                tmpAccDt = "End"
                tmpAccSeq = "End"
            End If
            
            Dim aryFoot() As String
            
            
            
            If tmpWorkArea <> .WorkArea Or tmpAccDt <> .AccDt Or tmpAccSeq <> .AccSeq Then
                'Remark 출력
                If Trim(.RmkCd) <> "" Then
                    strBuffer = GetRemark(.RmkCd)
                    Printer.FontBold = True
                    Call Print_Setting("☞ Remark   : ", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
                    '추가함
                    Call CheckNewPage
                    Printer.FontBold = False
                    Call Print_Setting(strBuffer, iposSpcNm + iCm * 0.4 + Printer.TextWidth("☞ Remark   :   "), LineSpace, Twidth, "L", "C")
                    CheckNewPage
                End If
                'Footnote 출력
                If Trim(.FootNoteFg) <> "0" Then
                    strBuffer = GetFootNote_New(.WorkArea, .AccDt, Val(.AccSeq))
                    'strBuffer = Replace(strBuffer, vbLf, "")
                    aryFoot() = Split(strBuffer, vbCrLf)
                    Printer.FontBold = True
                    If strBuffer <> "" Then
                        Call Print_Setting("☞ FootNote : ", iposSpcNm + iCm * 0.4, LineSpace, Twidth, "L", "C", False)
                        '추가함
                        Call CheckNewPage
                        Printer.FontBold = False
                        For ii = LBound(aryFoot) To UBound(aryFoot)
                            If Trim(aryFoot(ii)) <> "" Then
                                If LenB(StrConv(aryFoot(ii), vbFromUnicode)) > 60 Then
                                    lngCnt = LenB(StrConv(aryFoot(ii), vbFromUnicode)) \ 60
                                    kk = 1
                                    For jj = 1 To lngCnt
                                        Call Print_Setting(Trim(Mid(aryFoot(ii), kk, 60)), iposSpcNm + iCm * 0.4 + Printer.TextWidth("FootNote :   "), LineSpace, Twidth, "L", "C")
                                        Call CheckNewPage
                                        kk = kk + 60
                                    Next
                                    If Trim(Mid(aryFoot(ii), kk)) <> "" Then
                                        Call Print_Setting(Trim(Mid(aryFoot(ii), kk)), iposSpcNm + iCm * 0.4 + Printer.TextWidth("☞ FootNote :   "), LineSpace, Twidth, "L", "C")
                                        '추가함
                                        Call CheckNewPage
                                    End If
                                Else
                                    Call Print_Setting(aryFoot(ii), iposSpcNm + iCm * 0.4 + Printer.TextWidth("☞ FootNote :   "), LineSpace, Twidth, "L", "C")
                                    Call CheckNewPage
                                End If
                            End If
                        Next
                        Call ChangeLine
                    End If
                End If
            End If
            
            If .TxtFg = "Y" And .TestDiv = enTestDiv.TST_SpeTest Then
                lngCurYPos = LastLineYpos
            End If
            
            '** 추가 인증문구 By M.G.Choi 2006.09.06
            Printer.FontSize = 9: Printer.FontBold = True
            
            '-- 특수검사는 제외 처리 함
            If .TestDiv = enTestDiv.TST_SpeTest Then
                
            Else
                If tmpVfyNm = "" Then
                    tmpVfyNm = GetEmpNm(.VfyId)
                    Call P_FIX("보고자 : " & tmpVfyNm & " " & "M.T.", PrtLeft, LastLineYpos - 15, Twidth - (PrtLeft * 2), "R", , "C")
                End If
                
                Call P_FIX("확인자 : " & P_HOSPITALCHIP & " " & "M.D.", PrtLeft, LastLineYpos - 12, Twidth - (PrtLeft * 2), "R", , "C")
            End If
            
            Call P_FIX(P_HOSPITALCOMMENT, PrtLeft, LastLineYpos - 6, Twidth - PrtLeft, "C", , "C")
            Call P_FIX(P_HOSPITALADDR_ENG, PrtLeft, LastLineYpos - 3, Twidth - PrtLeft, "C", , "C")
        
        End With
        
    Next
    
    
    If mvarImage = False Then
        For i = 1 To mResults.Count
            With mResults.Item(i)
                If PrintUpdate(.WorkArea, .AccDt, .AccSeq, .TestCd, .TestDiv) = False Then GoTo Err_Trap
            End With
        Next
    End If
    
    Call Print_Last
    Printer.EndDoc
    Exit Sub
    
Err_Trap:
    ErrorWrite (Err.Description)
On Error GoTo Err_Trap
    Resume Next
    Resume
End Sub


Private Sub PrtImageBody()
   Dim objimage   As Image
   Dim Rs As New Recordset
   Dim objSql As New clsLISSqlStatement
    
On Error GoTo Err_Trap
    
    Printer.FontSize = 9
                
    Call PrtHeader
   
    Set objimage = frmControls.Image1
    
    If P_ImageSystem = True And mvarImage = True Then
    
'        Call ChangeLine
        Call CheckNewPage
        Printer.FontBold = True
        Printer.FontSize = 12
        
        If mvarTestNm = "" Then
            Rs.Open objSql.SqlLAB001Read(mvarTestCd), DBConn
            If Rs.RecordCount > 0 Then
                mvarTestNm = Rs.Fields("testnm").Value & ""
            End If
        End If
        
        Call Print_Setting(mvarTestNm, 0, LineSpace, Twidth - Len(mvarTestNm), "C", "C", True)
'        Call Print_Setting(String(80, Chr(6)), iposOrdDt, LineSpace, Twidth, "L", "C", True)
        Printer.FontBold = False
        Printer.FontSize = 9
        Call CheckNewPage
        
        Call Print_Setting("", PrtLeft, LineSpace / 8, Twidth, "R", "C")
        '추가함
        Call CheckNewPage
        Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
        Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")
        '추가함
        Call CheckNewPage
        
        Call ChangeLine
        Call CheckNewPage
        objimage.Picture = LoadPicture(strSlidePath)
        Printer.PaintPicture objimage.Picture, iposOrdDt + 35, LineSpace * 14, 100, 100
        Call Print_Setting("", iposOrdDt, LineSpace * 20, Twidth, "L", "C", True)
        Call CheckNewPage
    End If
   
    
    Call Print_Last
    Printer.EndDoc
    
    Set Rs = Nothing
    Set objSql = Nothing
    Exit Sub
    
Err_Trap:
    ErrorWrite (Err.Description)
    Set Rs = Nothing
    Set objSql = Nothing
    
On Error GoTo Err_Trap
    Resume Next
    
End Sub

Private Function PrintUpdate(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal TestCd As String, ByVal TestDiv As String) As Boolean
    Dim strTblNm As String
    Dim sSQL     As String
    
    Select Case TestDiv
        Case "0": strTblNm = T_LAB302
        Case "1": strTblNm = T_LAB351
        Case "2": strTblNm = T_LAB404
    End Select
    
    On Error GoTo SAVE_ERROR
    
    DBConn.BeginTrans
    
    sSQL = " UPDATE " & strTblNm & _
           " set " & _
                    DBW("rptfg", "Y", 3) & _
                    DBW("rptdt", Format(Now, CS_DateDbFormat), 3) & _
                    DBW("rpttm", Format(Now, CS_TimeDbFormat), 3) & _
                    DBW("rptid", ObjSysInfo.EmpId, 2) & _
           " WHERE " & _
                    DBW("workarea=", WorkArea) & " AND " & _
                    DBW("accdt=", AccDt) & " AND " & _
                    DBW("accseq=", AccSeq) & " AND " & _
                    DBW("testcd=", TestCd)
                
    If TestDiv <> "0" Then
        sSQL = sSQL & _
                       "   AND " & DBW("stscd >= ", enStsCd.StsCd_LIS_FinRst)
    End If
    DBConn.Execute sSQL
    DBConn.CommitTrans
    PrintUpdate = True
    Exit Function
SAVE_ERROR:
    DBConn.RollbackTrans
    PrintUpdate = False
End Function


Private Sub CheckNewPage()
    
    If lngCurYPos > LastLineYpos - (1# * iCm) Then  ' newPage일 경우
        PageNumber = PageNumber + 1
        Printer.Line (PrtLeft, LastLineYpos)-(Twidth - PrtLeft, LastLineYpos)
        Call P_FIX(PageNumber, PrtLeft, LastLineYpos + 3, Twidth - PrtLeft, "C", , "C")
        Printer.NewPage
        Call PrtHeader
    End If
            
End Sub

Private Sub Print_Last()
    Dim strAddress As String
    Dim PageNumber As Integer
    Dim dwLen As Long
    Dim strDate As String
    Dim strIP As String
    Dim strString As String
    
    SocketsInitialize
    strIP = GetTheIP
    SocketsCleanup
 
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    strDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    '-- 예수병원 주소 추가 =================================================================
    strAddress = "전북 전주시 완산구 중화산동 1-300"
    '=======================================================================================
    
    PageNumber = PageNumber + 1
    Printer.Line (PrtLeft, LastLineYpos)-(Twidth - PrtLeft, LastLineYpos)
    Call P_FIX(PageNumber, PrtLeft, LastLineYpos + 7, Twidth - PrtLeft, "C", , "C")
    
    If Not mvarRouding Then '회진용은 Tail 찍지말자..
        Printer.FontSize = 13: Printer.FontBold = True
        If mvarRiPrint = "" Then
            Call P_FIX(P_HOSPITALNAME & " 진단검사의학과" & Space(5) & strAddress, PrtLeft, LastLineYpos + 2, Twidth - PrtLeft, "C", , "C")
        Else
            Call P_FIX(P_HOSPITALNAME & " 핵의학실", PrtLeft, LastLineYpos + 3, Twidth - PrtLeft, "C", , "C")
        End If
    End If
    
    Printer.FontSize = 9 'Printer.FontBold = True
    Printer.ForeColor = &HC0C0C0
    Call P_FIX("출력일시 : " & strDate & Space(5) & "출력IP : " & strIP & Space(5) & "출력ID : " & ObjMyUser.EmpLngNm & "(" & ObjMyUser.EmpId & ")", PrtLeft, LastLineYpos + 7, Twidth - PrtLeft, "", , "C")
    Printer.ForeColor = vbBlack

    Printer.FontSize = 9: Printer.FontBold = False
End Sub

Private Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select
         
    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    iCurY = Printer.CurrentY + iSpace
End Sub

Private Sub prtPageNum()
    
    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    
    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
    
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + iCm / 6
    Printer.Print "RUN-TIME : " & sTime
        
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
End Sub


Private Function GetRemark(ByVal RmkCd As String)
   
   Dim tmpSql As String
   Dim tmpRs As Recordset
   Dim MySql As New clsLISSqlReview
   
   GetRemark = ""
   tmpSql = MySql.SqlGetRemark(RmkCd)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   
   If tmpRs.EOF Then GoTo Nodata
   
   GetRemark = Trim("" & tmpRs.Fields("Remark").Value)

Nodata:
   Set tmpRs = Nothing
   Set MySql = Nothing
   
End Function


Private Function GetFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
   
   Dim i As Integer
   Dim tmpSql As String
   Dim tmpRs As Recordset
   Dim MySql As New clsLISSqlReview
   
   GetFootNote = ""
   tmpSql = MySql.SqlGetFootNote(pWorkArea, pAccDt, pAccSeq)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   
   If tmpRs.EOF Then GoTo Nodata
   
   While (Not tmpRs.EOF)
      GetFootNote = GetFootNote & Trim("" & tmpRs.Fields("FootNote").Value)
      tmpRs.MoveNext
   Wend
   
Nodata:
   Set tmpRs = Nothing
   Set MySql = Nothing
   
End Function

Private Function GetFootNote_New(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
   
   Dim i As Integer
   Dim tmpSql As String
   Dim tmpRs As Recordset
   Dim MySql As New clsLISSqlReview
   
   GetFootNote_New = ""
   tmpSql = MySql.SqlGetFootNote(pWorkArea, pAccDt, pAccSeq)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   
   If tmpRs.EOF Then GoTo Nodata
   
   While (Not tmpRs.EOF)
      GetFootNote_New = GetFootNote_New & Trim("" & tmpRs.Fields("FootNote").Value)
      tmpRs.MoveNext
   Wend
   
Nodata:
   Set tmpRs = Nothing
   Set MySql = Nothing
   
End Function

Public Function GetRstText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)
  
    Dim i        As Integer
    Dim tmpRs    As Recordset
    Dim MySql    As New clsLISSqlReview
    
    '텍스트 결과내역
    Set tmpRs = New Recordset
    tmpRs.Open MySql.SqlGetRstText(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
    
    If tmpRs.EOF Then
        GetRstText = ""
    Else
        GetRstText = Trim("" & tmpRs.Fields("RstTxt").Value)
    End If
    '텍스트 결과 수정내역
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open MySql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
    
    If Not tmpRs.EOF Then
        GetRstText = GetRstText & vbCrLf & Trim(tmpRs.Fields("rsttxt").Value & "")
    End If
    
    Set tmpRs = Nothing
    Set MySql = Nothing
End Function


Public Function GetSpcText(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                        ByVal pAccSeq As Integer, ByVal pTestCd As String, ByVal pMfySeq As String)

    Dim SqlStmt As String
    Dim tmpRs As Recordset
    
    SqlStmt = "Select txtrst as TextResult From " & T_LAB353 & "  " & _
              "Where " & _
                              DBW("workarea=", pWorkArea) & " " & _
                     "and " & DBW("accdt=", pAccDt) & " " & _
                     "and " & DBW("accseq=", pAccSeq) & " " & _
                     "and " & DBW("testcd=", pTestCd) & " " & _
                     "and " & DBW("mfyseq=", pMfySeq)
              
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
   If tmpRs.EOF Then
      GetSpcText = ""
   Else
      GetSpcText = Trim("" & tmpRs.Fields("TextResult").Value)
   End If
   Set tmpRs = Nothing

End Function


Public Function GetSenResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim i As Integer, j As Integer
    Dim AntiList As ListBox, AntiCnt As Integer
    Dim AntiSortList As ListBox
    Dim AntiRst As String, AntiCd As String, AntiNm As String
    Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
    Dim tmpTitle As String
    Dim tmpAntiCnt As Integer
    Dim MySql As New clsLISSqlReview
    Dim MyCommSql As New clsLISSqlStatement
    Dim SenRstBuf As clsResults
    Dim SenClipText As String
    Dim aryMicFg() As String
    
    Dim MySenResult() As SenResult
   
On Error GoTo Err_Trap

    Const MaxAntiCnt = 50
   
    GetSenResult = ""
    SenClipText = ""
   
    SqlStmt = MySql.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo Nodata
   
    MicroCnt = tmpRs.RecordCount  '균갯수
    ReDim MySenResult(MaxAntiCnt, MicroCnt)
    ReDim aryMicFg(MicroCnt)
   
    Set AntiList = frmControls.lstUnsortedList   '항생제 리스트(Unsorted)
    Set AntiSortList = frmControls.lstList   '항생제 리스트(Unsorted)
    AntiList.Clear
    AntiSortList.Clear
   
    tmpAntiCnt = 0
    tmpTitle = "Antibiotics  " & Space(26 - Len("Antibiotics  ")) '& vbTab
   
    '미생물 감수성 결과내역 Buffering (LAB405)
    For i = 1 To MicroCnt
       
        'tmpTitle = tmpTitle & " " & CStr(i) & " "
        AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
        aryMicFg(i) = "" & tmpRs.Fields("MicFg").Value  'MIC여부
        
        If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
'        If AntiCnt > 0 Then
'            tmpTitle = tmpTitle & " " & CStr(I) & " " & Space(11 - Len(CStr(I))) & vbTab
            If aryMicFg(i) = MRT_GenSen Then
                tmpTitle = tmpTitle & " " & CStr(i) & " "
            ElseIf aryMicFg(i) = MRT_MicSen Then
                tmpTitle = tmpTitle & " " & CStr(i) & Space(7) & " "
            End If
'        End If
        
        GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
        
        If Len(GrowthQty) > 15 Then
            SenClipText = SenClipText & Mid(GrowthQty, 1, 15) '& vbTab '배양정도
        Else
            SenClipText = SenClipText & GrowthQty & Space(15 - Len(GrowthQty)) '& vbTab '배양정도
        End If
        SenClipText = SenClipText & " " & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
        
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
   
    SenClipText = SenClipText & vbCrLf
   
    If tmpAntiCnt > 0 Then  '감수성결과가 있으면...
        SenClipText = SenClipText & "[ Susceptibility test ]" & vbCrLf
        SenClipText = SenClipText & tmpTitle & vbCrLf
    End If

    '감수성 결과 Buffering
    Dim mSenRst As New Collection
    Dim strRst1 As String, strRst2 As String
   
    For i = 1 To AntiList.ListCount
        AntiCd = AntiList.List(i - 1)
        AntiCd = Mid(AntiCd, 1, Len(AntiCd) - 1)
        tmpRs1.Open MyCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
        AntiNm = Trim("" & tmpRs1.Fields("text1").Value)     '항생제명
        If AntiNm = "" Then AntiNm = AntiCd
        Set tmpRs1 = Nothing
        
        Set SenRstBuf = New clsResults
        SenRstBuf.AntiCd = AntiCd
        SenRstBuf.AntiNm = AntiNm
        AntiSortList.AddItem AntiNm   '항생제명 순으로 정렬리스트에 추가 ...
        
        For j = 1 To MicroCnt
'            If MySenResult(i - 1, j).RstCd = "" Then
'                SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   "
'            Else
'                SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " " & MySenResult(i - 1, j).RstCd & " "  '균별 감수성결과
'            End If
            strRst1 = medGetP(MySenResult(i - 1, j).RstCd, 1, ";")  'SRI결과
            strRst2 = medGetP(MySenResult(i - 1, j).RstCd, 2, ";")  'MIC결과
            If strRst1 = "" Then
                If aryMicFg(j) = MRT_MicSen Then
                    SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   " & Space(7)
                Else
                    SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   "
                End If
            Else
'                If strRst1 = "N" Then
'                    strRst2 = "NEGATIVE"
'                ElseIf strRst1 = "P" Then
'                    strRst2 = "POSITIVE"
'                ElseIf strRst1 = "-" Then
'                    strRst2 = "SYN-R"
'                End If
                
                SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " " & strRst1  '균별 감수성결과
                If Trim(strRst2) <> "" Then
                    If Mid(strRst2, 1, 1) = "≤" Or Mid(strRst2, 1, 1) = "≥" Then
                        SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "(" & Trim(strRst2) & ")" & Space(6 - Len(strRst2) - 1)
                    Else
                        SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "(" & Trim(strRst2) & ")" & Space(6 - Len(strRst2))
                    End If
                Else
                    SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " "
                End If
            End If
        Next
        mSenRst.Add SenRstBuf, AntiNm
    Next
   
    '균코드
    For i = 1 To AntiSortList.ListCount
        AntiNm = AntiSortList.List(i - 1)
        If Len(AntiNm) > 26 Then
            SenClipText = SenClipText & Mid(AntiNm, 1, 26) '& vbTab '항생제명
        Else
            SenClipText = SenClipText & AntiNm & Space(26 - Len(AntiNm)) '& vbTab  '항생제명
        End If
        SenClipText = SenClipText & mSenRst.Item(AntiNm).SenRstBuf & vbCrLf
    Next
    
Nodata:
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    Set AntiList = Nothing
    Set MySql = Nothing
    Set mSenRst = Nothing

    GetSenResult = SenClipText
    Exit Function
'

Err_Trap:
    ErrorWrite (Err.Description)
On Error GoTo Err_Trap
    Resume Next

End Function

Public Function GetSenResult_New(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pTestCd As String)

    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim i As Integer, j As Integer
    Dim AntiList As ListBox, AntiCnt As Integer
    Dim AntiSortList As ListBox
    Dim AntiRst As String, AntiCd As String, AntiNm As String
    Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
    Dim tmpTitle As String
    Dim tmpAntiCnt As Integer
    Dim MySql As New clsLISSqlReview
    Dim MyCommSql As New clsLISSqlStatement
    Dim SenRstBuf As clsResults
    Dim SenClipText As String
    Dim aryMicFg() As String
    
    Dim MySenResult() As SenResult
   
On Error GoTo Err_Trap

    Const MaxAntiCnt = 50
   
    GetSenResult_New = ""
    SenClipText = ""
   
    SqlStmt = MySql.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
    tmpRs.Open SqlStmt, DBConn
   
    If tmpRs.EOF Then GoTo Nodata
   
    MicroCnt = tmpRs.RecordCount  '균갯수
    ReDim MySenResult(MaxAntiCnt, MicroCnt)
    ReDim aryMicFg(MicroCnt)
   
    Set AntiList = frmControls.lstUnsortedList   '항생제 리스트(Unsorted)
    Set AntiSortList = frmControls.lstList   '항생제 리스트(Unsorted)
    AntiList.Clear
    AntiSortList.Clear
   
    tmpAntiCnt = 0
    tmpTitle = "Antibiotics  " & Space(26 - Len("Antibiotics  ")) '& vbTab
   
    '미생물 감수성 결과내역 Buffering (LAB405)
    For i = 1 To MicroCnt
       
        'tmpTitle = tmpTitle & " " & CStr(i) & " "
        AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
        aryMicFg(i) = "" & tmpRs.Fields("MicFg").Value  'MIC여부
        
        If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
'        If AntiCnt > 0 Then
'            tmpTitle = tmpTitle & " " & CStr(I) & " " & Space(11 - Len(CStr(I))) & vbTab
            If aryMicFg(i) = MRT_GenSen Then
                tmpTitle = tmpTitle & " " & CStr(i) & " "
            ElseIf aryMicFg(i) = MRT_MicSen Then
                tmpTitle = tmpTitle & " " & CStr(i) & Space(7) & " "
            End If
'        End If
        
        GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
        SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
        
        If Len(GrowthQty) > 15 Then
            SenClipText = SenClipText & Mid(GrowthQty, 1, 15) '& vbTab '배양정도
        Else
            SenClipText = SenClipText & GrowthQty & Space(15 - Len(GrowthQty)) '& vbTab '배양정도
        End If
        SenClipText = SenClipText & " " & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
        
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
   
    SenClipText = SenClipText & vbCrLf
   
    If tmpAntiCnt > 0 Then  '감수성결과가 있으면...
        SenClipText = SenClipText & "[ Susceptibility test ]" & vbCrLf
        SenClipText = SenClipText & tmpTitle & vbCrLf
    End If

    '감수성 결과 Buffering
    Dim mSenRst As New Collection
    Dim strRst1 As String, strRst2 As String
   
    For i = 1 To AntiList.ListCount
        AntiCd = AntiList.List(i - 1)
        AntiCd = Mid(AntiCd, 1, Len(AntiCd) - 1)
        tmpRs1.Open MyCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
        AntiNm = Trim("" & tmpRs1.Fields("text1").Value)     '항생제명
        If AntiNm = "" Then AntiNm = AntiCd
        Set tmpRs1 = Nothing
        
        Set SenRstBuf = New clsResults
        SenRstBuf.AntiCd = AntiCd
        SenRstBuf.AntiNm = AntiNm
        AntiSortList.AddItem AntiNm   '항생제명 순으로 정렬리스트에 추가 ...
        
        For j = 1 To MicroCnt
'            If MySenResult(i - 1, j).RstCd = "" Then
'                SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   "
'            Else
'                SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " " & MySenResult(i - 1, j).RstCd & " "  '균별 감수성결과
'            End If
            strRst1 = medGetP(MySenResult(i - 1, j).RstCd, 1, ";")  'SRI결과
            strRst2 = medGetP(MySenResult(i - 1, j).RstCd, 2, ";")  'MIC결과
            If strRst1 = "" Then
                If aryMicFg(j) = MRT_MicSen Then
                    SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   " & Space(7)
                Else
                    SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   "
                End If
            Else
'                If strRst1 = "N" Then
'                    strRst2 = "NEGATIVE"
'                ElseIf strRst1 = "P" Then
'                    strRst2 = "POSITIVE"
'                ElseIf strRst1 = "-" Then
'                    strRst2 = "SYN-R"
'                End If
                
                SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " " & strRst1  '균별 감수성결과
                If Trim(strRst2) <> "" Then
                    If Mid(strRst2, 1, 1) = "≤" Or Mid(strRst2, 1, 1) = "≥" Then
                        SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "(" & Trim(strRst2) & ")" & Space(6 - Len(strRst2) - 1)
                    Else
                        SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "(" & Trim(strRst2) & ")" & Space(6 - Len(strRst2))
                    End If
                Else
                    SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " "
                End If
            End If
        Next
        mSenRst.Add SenRstBuf, AntiNm
    Next
   
    '균코드
    For i = 1 To AntiSortList.ListCount
        AntiNm = AntiSortList.List(i - 1)
        If Len(AntiNm) > 26 Then
            SenClipText = SenClipText & Mid(AntiNm, 1, 26) '& vbTab '항생제명
        Else
            SenClipText = SenClipText & AntiNm & Space(26 - Len(AntiNm)) '& vbTab  '항생제명
        End If
        SenClipText = SenClipText & mSenRst.Item(AntiNm).SenRstBuf & vbCrLf
    Next
    
Nodata:
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    Set AntiList = Nothing
    Set MySql = Nothing
    Set mSenRst = Nothing

    GetSenResult_New = SenClipText
    Exit Function
'

Err_Trap:
    ErrorWrite (Err.Description)
On Error GoTo Err_Trap
    Resume Next

End Function

Public Function GetLastReportDt() As String

    Dim objSql As New clsLISSqlReport
    Dim rs1 As New Recordset
    
    rs1.Open objSql.SqlGetLastReportDt, DBConn
    If rs1.EOF Then
        GetLastReportDt = Format(DateAdd("d", -1, Now), CS_DateDbFormat)
    Else
        GetLastReportDt = "" & rs1.Fields("lastdt").Value
    End If
    Set rs1 = Nothing

End Function


Public Function GetWorkAreaName(ByVal pWorkArea As String) As String

    Dim objSql As New clsLISSqlStatement
    Dim objRs As Recordset
    
    Set objRs = New Recordset
    objRs.Open objSql.SqlLAB032CodeList(LC3_WorkArea, "field1", pWorkArea), DBConn
    
    GetWorkAreaName = ""
    If Not objRs.EOF Then
        GetWorkAreaName = "" & objRs.Fields("field1").Value
    End If
    Set objRs = Nothing
    
End Function

Public Function SetNoReportRptUpdate(ByVal FRDT As String, ByVal TODT As String) As Boolean
    Dim sSQL As String
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    sSQL = " UPDATE " & T_LAB302 & " SET " & _
                                    DBW("rptfg", "Y", 3) & _
                                    DBW("rptdt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                                    DBW("rpttm", Format(GetSystemDate, "HHMMSS"), 3) & _
                                    DBW("rptid", ObjSysInfo.EmpId, 2) & _
           " WHERE " & _
                      DBW("vfydt>=", FRDT) & " and " & _
                      DBW("vfydt<=", TODT) & "and " & _
           " testcd IN (" & P_NoResultReport & ")"
    
    DBConn.Execute sSQL
    DBConn.CommitTrans
    SetNoReportRptUpdate = True
    Exit Function
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox " 결과지 미출력 아이템에 대한 UPDATE 시 문제 발생입니다.", vbInformation + vbOKOnly, "결과지출력"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 특수검사 결과지 출력 - 5.0.9: 이상대(2005-07-12)
'          특수검사 결과등록폼에서 결과지 출력시 사용
'          해당 접수번호의 결과지만 출력
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : AccDt
'       - pAccSeq   : AccSeq
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Sub SpecialReport(ByVal pWorkArea As String, pAccDt As String, ByVal pAccSeq As String, _
                         ByVal pTestCd As String, ByVal pPtId As String, ByVal pFromDate As String, _
                         ByVal pToDate As String, ByVal pTestDiv As String, _
                         ByVal pImgPath As String, ByVal pESign As Object, _
                         ByRef pProgBar As Object, ByRef pLastDt As String, _
                         ByRef pLastTm As String)

    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String, SvWorkArea As String
    Dim objRptSql As New clsLISSqlReview
    Dim MyData As New clsResults
    Dim blnFirst As Boolean
    Dim strDPDiv As String
    
    
    Dim objDic As New clsDictionary
    
    mvarRealPrinter = False
    
    objDic.Clear
    objDic.FieldInialize "seq", "workarea,accdt,accseq,testcd,testdiv"
    objDic.Sort = False
    
    MouseRunning
    
    mvarErrorFg = False
    
    strImgPath = pImgPath
    Set picESign = pESign
    
    '처방일/접수일 기준
    objRptSql.RiPrint = mvarRiPrint
    SqlStmt = objRptSql.SqlSpecialReport(pWorkArea, pAccDt, pAccSeq, pTestCd, pPtId, pFromDate, pToDate, pTestDiv, mvarReprint Or mvarRouding, _
                                          mvarBatchReprint, gUsingInWardMenu)
    
    'Query
    blnFirst = False
    tmpRs.Open SqlStmt, DBConn
    
    SvWorkArea = "": SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    If tmpRs.EOF Then
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    Dim ii As Integer
    Dim strDept As String
    
    Do Until tmpRs.EOF
        
        ii = ii + 1
    
        DoEvents
                
        If Not blnFirst Then
            strDept = GetDeptNm(tmpRs.Fields("deptcd").Value & "")
            
            If mvarWard <> "" And strDept <> "" Then
                mvarDept = Trim("" & tmpRs.Fields("DeptCd").Value)
                mvarDeptNm = strDept
                blnFirst = True
            End If
            
        End If
        
        mvarWardId = Trim("" & tmpRs.Fields("wardid").Value)
        
        '**------------------------------------------------------------------------------
        '** 결과구분이 Alternative이면서 결과가 없을경우는 출력에서 제외한다. 2001.6.20
        '**------------------------------------------------------------------------------
        
        If Trim("" & tmpRs.Fields("RstDiv").Value) = "A" And Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
            
            GoTo Skip
        End If
        
        If P_MicSensiReport = False Then    'MIC Sensi인경우 출력에서 제외
            If Trim("" & tmpRs.Fields("TestCd").Value) Like P_MicSensiTestCd & "*" Then
                
            objDic.AddNew ii, Join(Array(tmpRs.Fields("workarea") & "", tmpRs.Fields("accdt") & "", _
                                        tmpRs.Fields("accseq") & "", tmpRs.Fields("testcd") & "", _
                                        tmpRs.Fields("testdiv") & ""), COL_DIV)
                
                
                GoTo Skip
            End If
        End If
        '**------------------------------------------------------------------------------
        
    
        If pLastDt < Trim("" & tmpRs.Fields("VfyDt").Value) Then
            pLastDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        ElseIf pLastTm < Trim("" & tmpRs.Fields("Vfytm").Value) Then
            pLastTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        End If
        
        Dim svStsCd As String

        If SvWorkArea <> Trim("" & tmpRs.Fields("WorkArea").Value) Then
            SvKeyDt = ""
            SvSpcNm = ""
            SvWorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
            svStsCd = Trim("" & tmpRs.Fields("stscd").Value)
        ElseIf SvWorkArea = Trim("" & tmpRs.Fields("workarea").Value) Then
            If svStsCd <> Trim("" & tmpRs.Fields("stscd").Value) Then
                SvKeyDt = ""
                SvSpcNm = ""
                svStsCd = Trim("" & tmpRs.Fields("stscd").Value)
            End If
        End If
        
        'If SvKeyDt <> Trim("" & tmpRs.fields("KeyDate").value) Then
        If SvKeyDt <> Trim("" & tmpRs.Fields("ColDtTm").Value) Then
            MyData.OrdDt = Trim("" & tmpRs.Fields("ColDtTm").Value)     '채혈일
            MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            SvKeyDt = Trim("" & tmpRs.Fields("ColDtTm").Value)
            SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
        Else
            MyData.OrdDt = "":
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                MyData.SpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                MyData.SpcNm = ""
            End If
        End If
        
        'MyData.OrdDate = Trim("" & tmpRs.fields("KeyDate").value)       '처방일
        MyData.OrdDate = Trim("" & tmpRs.Fields("ColDtTm").Value)       '처방일
        MyData.SpcName = Trim("" & tmpRs.Fields("SpcNm").Value)         '검체명
        
        '-- 검사명
        If Trim("" & tmpRs.Fields("TestDiv").Value) = enTestDiv.TST_SpeTest Then
            MyData.TestNm = Trim("" & tmpRs.Fields("TestLongNm").Value)
        Else
            tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestShortNm").Value), 1, 24)
            If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                
                MyData.TestNm = tmpTestNm & " " & String(24 - Len(tmpTestNm), ".")
            Else
                MyData.TestNm = Space(4) & tmpTestNm & " " & String(24 - Len("  " & tmpTestNm), ".")
            End If
        End If
                    
        '-- 결과명(코드일 경우..)
        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
            MyData.RstCd = "미확"
        Else
            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                MyData.RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
            Else
                MyData.RstCd = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
            End If
            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                MyData.RstCd = "Growth"
            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                MyData.RstCd = Space(3)
            End If
        End If
                    
        '-- 결과단위
        MyData.RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
        
        '-- High / Low
        MyData.HLDiv = ""
        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then MyData.HLDiv = HLDIV_HIGH_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then MyData.HLDiv = HLDIV_LOW_FG
            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then MyData.HLDiv = "*" 'Abnormal_Flag
        End If
        
        '## 5.0.8: 이상대(2005-05-30)
        '   - Alpha결과 참고치의 판정을 "Abnormal"로 표시하도록 수정
        '-- Delta/Panic
        strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
        strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
        MyData.DPDiv = strDPDiv
        
        MyData.WorkArea = Trim("" & tmpRs.Fields("WorkArea").Value)
        MyData.AccDt = Trim("" & tmpRs.Fields("AccDt").Value):
        MyData.AccSeq = Trim("" & tmpRs.Fields("AccSeq").Value)
        MyData.LastRst = Trim("" & tmpRs.Fields("LastRst").Value)           '-- 최근결과
        MyData.TestCd = Trim("" & tmpRs.Fields("TestCd").Value):            '-- 검사코드
        MyData.SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)               '-- 검체코드
        MyData.VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)               '-- 보고일
        MyData.VfyId = Trim("" & tmpRs.Fields("VfyId").Value)               '-- 보고자
        MyData.TestDiv = Trim("" & tmpRs.Fields("TestDiv").Value)           '-- 검사구분
        MyData.TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)               '-- 소견결과여부
        MyData.FootNoteFg = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
        MyData.RmkCd = Trim("" & tmpRs.Fields("RmkCd").Value)               '-- Remark 코드
        MyData.SenFg = Trim("" & tmpRs.Fields("SenFg").Value)               '-- 감수성 여부
        MyData.PtSex = Trim("" & tmpRs.Fields("Sex").Value)                 '-- Sex
        MyData.AgeDay = Trim("" & tmpRs.Fields("AgeDay").Value)             '-- AgeDay
        MyData.RefRng = CS_QuestionMark
        MyData.MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)               '-- 수정SEQ
        MyData.StsCd = Trim("" & tmpRs.Fields("StsCd").Value)               '-- 상태
        MyData.VfyDtTm = Trim("" & tmpRs.Fields("VfyDtTm").Value)
        MyData.RcvDtTm = Trim("" & tmpRs.Fields("RcvDtTm").Value)
        MyData.ColDtTm = Trim("" & tmpRs.Fields("ColDtTm").Value)
        MyData.ClinicalNotice = Trim("" & tmpRs.Fields("Notice").Value)
        MyData.DetailFg = Trim("" & tmpRs.Fields("detailfg").Value)
        
        Call Me.Add(MyData)
Skip:
        tmpRs.MoveNext
    Loop
    objDic.Sort = True
      
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    
    DoEvents
    
    '결과지출력을 않해도 되는항목들.....UPDATE
    If objDic.RecordCount > 0 Then
        objDic.MoveFirst
        Do Until objDic.EOF
            Call PrintUpdate(objDic.Fields("Workarea"), objDic.Fields("accdt"), objDic.Fields("accseq"), objDic.Fields("testcd"), objDic.Fields("testdiv"))
            objDic.MoveNext
        Loop
    End If

    If Me.Count = 0 Then
        
        MouseDefault
        Set tmpRs = Nothing
        Exit Sub
    End If
    
    For i = 1 To Me.Count
        
        '참고치 검색
        If Me.Item(i).RefRng <> CS_QuestionMark Then GoTo RefSkip
        
        tmpSex = Me.Item(i).PtSex
        tmpAgeDay = Me.Item(i).AgeDay
        tmpTestCd = Me.Item(i).TestCd
        tmpSpcCd = Me.Item(i).SpcCd
        tmpVfyDt = Me.Item(i).VfyDt
        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
     
        SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open SqlStmt, DBConn
        
        If tmpRs1.EOF Then
            '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
            '--> 거의 Both로 등록됨.
            SqlStmt = objRptSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
        End If
        If tmpRs1.EOF Then
            tmpRefCd = Space(5)
        Else
            tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
            tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
            tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
            If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
               tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        Set tmpRs1 = Nothing
        For j = i To Me.Count
            If Me.Item(j).TestCd = tmpTestCd Then Me.Item(j).RefRng = tmpRefCd
        Next
     
        DoEvents
        
RefSkip:
    Next
      
ExitPos:
'    pProgBar.Value = pProgBar.Max
    DoEvents
      
Nodata:
    
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
    mvarRealPrinter = False
    
    lngErrFile = FreeFile
    Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_ERROR.log" For Append As lngErrFile
    
    Call Print_Report
    
    Close #lngErrFile

End Sub


