VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResultReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarPtId As String
Private mvarPtNm As String
Private mvarPtSex As String
Private mvarPtAge As String
Private mvarFromDt As String
Private mvarToDt As String
Private mvarDeptCd As String
Private mvarVfyDt   As String

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

Private lngCm As Integer

Const iCm = 567
Const iLineHeight = 10

Private iposOrdDt%, iposSpcNm%, iposTestNm%, iposRstCd%, iposLastRst%, _
    iposUnit%, iposHL, iposDP%, iposRefRng%, iposText%

Private mResults As Collection
Private mvarCount As Integer '지역 복사
Private FirstFg As Boolean

Private Type SenResult
   RstCd As String
   Row As String
   ForeColor As Long
End Type


Public Property Get Count() As Integer
    Count = mResults.Count
End Property

Public Property Get Item(ByVal KeyIndex As Variant) As clsResults
    Set Item = mResults(KeyIndex)
End Property


Private Sub Class_Initialize()
    
    Set mResults = New Collection
    mvarCount = 0
    'iposText = iposRefRng + 7.8 * lngCm
    
End Sub


Public Function Add(ByVal clsData As clsResults, Optional ByVal clsKey As Variant)
   
   Dim newItem As New clsResults
   
   With newItem
        .OrdDt = clsData.OrdDt
        .SpcNm = clsData.SpcNm
        .TestNm = clsData.TestNm
        .VfyDt = clsData.VfyDt
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
        .ColDtTm = clsData.ColDtTm
        .DetailFg = clsData.DetailFg
        .RstDiv = clsData.RstDiv
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

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

Public Property Let DeptCd(ByVal vNewValue As String)
    mvarDeptCd = vNewValue
End Property


Public Property Get VfyDt() As String
    VfyDt = mvarVfyDt
End Property

Public Property Let VfyDt(ByVal vNewValue As String)
    mvarVfyDt = vNewValue
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


Public Sub Print_Report()
    
    Call InitReport
    Call PrtBody
    
End Sub

Private Sub InitReport()
    Printer.Font = "굴림체"
    Printer.PaperSize = vbPRPSExecutive
    
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
    lngCm = CInt(iPageWidth / 20.15)
    
    iposOrdDt = lngCm * 1
    iposSpcNm = iposOrdDt + lngCm * 1.5
    iposTestNm = iposSpcNm + 2.2 * lngCm
    iposRstCd = iposTestNm + lngCm * 5
    
    iposLastRst = iposRstCd + lngCm * 3
    iposUnit = iposLastRst + lngCm * 3.5
    iposHL = iposUnit + lngCm * 1.6
    iposDP = iposHL + lngCm * 1.5
    iposRefRng = iposDP + lngCm * 1.2
    
    HeadString1 = "♧  환자별 검사결과"
    HeadString2 = "   등록번호 : " & mvarPtId & Space(3) & _
                  " 환자명 : " & mvarPtNm & ICSPatientString(mvarPtId, enICSNum.ResultReport) & Space(3) & _
                  " 성별/나이 : " & mvarPtSex & " / " & mvarPtAge
    HeadString3 = "   진 료 과 : " & mvarDeptCd & Space(3) & " 보고일자 : " & mvarVfyDt
   
End Sub

Private Sub PrtHeader()
   
    Printer.CurrentY = 0
    Printer.CurrentX = lngCm * 2
    iCurY = 0
    
    Call WriteStr(iCurY, lngCm * 1, HeadString1, iCurY, lngCm)
    Call WriteStr(iCurY, lngCm * 1, HeadString2, iCurY, lngCm / 2)
    Call WriteStr(iCurY, lngCm * 1, HeadString3, iCurY, lngCm * 0.7)
    
    Call DrawLine(lngCm * 1, iCurY, iPageWidth - lngCm / 4, iCurY, "solid", 2, lngCm / 6)
    
    Printer.FontBold = True
    Call WriteStr(iCurY, iposOrdDt, "처방일", iCurY, 0)
    Call WriteStr(iCurY, iposSpcNm, "    검체", iCurY, 0)
    Call WriteStr(iCurY, iposTestNm, "        검  사  명", iCurY, 0)
    Call WriteStr(iCurY, iposRstCd, "    결과", iCurY, 0)
    Call WriteStr(iCurY, iposLastRst, "  최근결과", iCurY, 0)
    Call WriteStr(iCurY, iposUnit, "   단위", iCurY, 0)
    Call WriteStr(iCurY, iposHL, "     HL", iCurY, 0)
    Call WriteStr(iCurY, iposDP, " DP", iCurY, lngCm / 2)   '0)
'    Call WriteStr(iCurY, iposRefRng, "       참 고 치", iCurY, lngCm / 2)
    
'    Call WriteStr(iCurY, iposText, "  ", iCurY, lngCm / 2)
    Printer.FontBold = False
    
    Call DrawLine(lngCm * 1, iCurY, iPageWidth - lngCm / 4, iCurY, "solid", 2, lngCm / 6)
    
    FirstFg = True
 
End Sub

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


Private Sub ChangeLine(iLineSpace As Integer)
    iCurY = iCurY + iLineSpace
    Printer.CurrentY = iCurY
    Printer.CurrentX = lngCm / 2
    
End Sub

Private Sub PrtBody()

    Dim i%
    Dim tmpWorkArea As String
    Dim tmpAccDt As String
    Dim tmpAccSeq As String
    Dim strBuffer As String
    Dim strTxt  As String
    Dim strOrdDt    As String
    Dim strColDtTm  As String
    Dim strSpcNm    As String
    
    Dim objRichText As RichTextBox
    
    Printer.FontSize = 9
    
    tmpWorkArea = ""
    tmpAccDt = ""
    tmpAccSeq = ""
    
    DoEvents
    Call PrtHeader
    
    Set objRichText = frmControls.rtfTextBox
    
    For i = 1 To mResults.Count
        With mResults.Item(i)
            
            
            If FirstFg Then
                .OrdDt = .OrdDate
                .SpcNm = .SpcName
                FirstFg = False
            End If
            '상세 코드일경우 상세 항목들을 모두 출력한후에 소견결과를 출력한다.
'            Debug.Print Trim(Replace(.TestNm, ".", "")) & "----------> " & strBuffer & iCurY
            If strTxt <> "" And .DetailFg = "" Then
                Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "소견결과 : ", iCurY, 0): Printer.FontBold = True: Printer.FontBold = False
                While (strTxt <> "")
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("소견결과 :   "), medShift(strTxt, vbCr), iCurY, lngCm / 2)
                    CheckNewPage
                Wend
                Call ChangeLine(lngCm / 2)
                strTxt = ""
            End If
            
            If strOrdDt <> .OrdDt Then
                Call WriteStr(iCurY, iposOrdDt, .OrdDt, iCurY, 0)
                Printer.FontBold = True
                strOrdDt = .OrdDt
            End If
            
'            Debug.Print .OrdDt; , "----->   "; .SpcNm
            If strSpcNm <> .SpcNm Then
                Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, .SpcNm, iCurY, 0)
                strSpcNm = .SpcNm
            End If
            
            If .ColDtTm <> strColDtTm Then
'                 Debug.Print .ColDtTm
'                 Call WriteStr(iCurY + 292, iposOrdDt, "", iCurY, 0)
'                 strColDtTm = .ColDtTm
            End If
           
           'Call WriteStr(iCurY, iposOrdDt, .ColDtTm, iCurY, 0)
            
            Printer.FontBold = False
            Call WriteStr(iCurY, iposTestNm, .TestNm, iCurY, 0)
            Call WriteStr(iCurY, iposRstCd, .RstCd, iCurY, 0, vbCenter, iposLastRst - iposRstCd)
            Call WriteStr(iCurY, iposLastRst, .LastRst, iCurY, 0, vbCenter, iposUnit - iposLastRst)
            'Call WriteStr(iCurY, iposRstCd + lngCm * 0.4, .RstCd, iCurY, 0, vbCenter, iposLastRst - iposRstCd)
            'Call WriteStr(iCurY, iposLastRst + lngCm * 0.2, "999.9", iCurY, 0, vbCenter, iposUnit - iposLastRst)
            Call WriteStr(iCurY, iposUnit + lngCm * 0.4, .RstUnit, iCurY, 0)
            
            Call WriteStr(iCurY, iposHL + lngCm * 0.9, .HLDiv, iCurY, 0): Printer.FontBold = True: Printer.FontBold = False
            Call WriteStr(iCurY, iposDP + lngCm * 0.3, .DPDiv, iCurY, lngCm / 2)    '0)
            'Call WriteStr(iCurY, iposRefRng, .RefRng, iCurY, lngCm / 2, vbCenter, iPageWidth - iposRefRng)  '가운데 정렬
            'Call WriteStr(iCurY, iposText, "  ", iCurY, lngCm / 2)
            
            CheckNewPage
            
            '일반검사 Text결과 출력
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
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "소견결과 : ", iCurY, 0): Printer.FontBold = True: Printer.FontBold = False
                    While (strTxt <> "")
                        Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("소견결과 :   "), medShift(strTxt, vbCr), iCurY, lngCm / 2)
                        CheckNewPage
                    Wend
                    strTxt = ""
                    Call ChangeLine(lngCm / 2)
                End If
            End If
            
            '상세항목만 처방난경우 마지막에 출력하기 위해 삽입.
            If i = mResults.Count And strTxt <> "" Then
                Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "소견결과 : ", iCurY, 0): Printer.FontBold = True: Printer.FontBold = False
                While (strTxt <> "")
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("소견결과 :   "), medShift(strTxt, vbCr), iCurY, lngCm / 2)
                    CheckNewPage
                Wend
                Call ChangeLine(lngCm / 2)
                strTxt = ""
            End If
            
            '기타검사 Text결과 출력
            If .TxtFg = "Y" Then
                objRichText.TextRTF = GetSpcText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd, .MfyFg)
                strBuffer = objRichText.Text
                'strBuffer = GetSpcText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd, "0")
                strBuffer = Replace(strBuffer, vbLf, "")
                Printer.FontBold = True
                Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "소견결과 : ", iCurY, lngCm / 2)
                Printer.FontBold = False
                While (strBuffer <> "")
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("소견결과 :   "), medShift(strBuffer, vbCr), iCurY, lngCm / 2)
                    CheckNewPage
                Wend
                'Supplemental Report
                strBuffer = GetSpcText(.WorkArea, .AccDt, Val(.AccSeq), .TestCd, "1")
                If strBuffer <> "" Then
                    strBuffer = Replace(strBuffer, vbLf, "")
                    Printer.FontBold = True
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "Supplemental Report : ", iCurY, lngCm / 2)
                    Printer.FontBold = False
                    While (strBuffer <> "")
                        Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("Supplemental Report :   "), medShift(strBuffer, vbCr), iCurY, lngCm / 2)
                        CheckNewPage
                    Wend
                    Call ChangeLine(lngCm / 2)
                End If
                
            End If
            
            '감수성 결과 출력
            If .SenFg = "Y" Then
                strBuffer = GetSenResult(.WorkArea, .AccDt, Val(.AccSeq), .TestCd)
                strBuffer = Replace(strBuffer, vbLf, "")
                If strBuffer <> "" Then
                    Printer.FontBold = True
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "감수성 결과 : ", iCurY, lngCm / 2)
                    Printer.FontBold = False
                    While (strBuffer <> "")
                        Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("감수성 결과 :   "), medShift(strBuffer, vbCr), iCurY, lngCm / 2)
                        CheckNewPage
                    Wend
                    Call ChangeLine(lngCm / 2)
                End If
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
            
            If tmpWorkArea <> .WorkArea Or tmpAccDt <> .AccDt Or tmpAccSeq <> .AccSeq Then
                'Remark 출력
                If Trim(.RmkCd) <> "" Then
                    strBuffer = GetRemark(.RmkCd)
                    Printer.FontBold = True
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "Remark : ", iCurY, 0)
                    Printer.FontBold = False
                    Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("Remark :   "), strBuffer, iCurY, lngCm / 2)
                    CheckNewPage
                    Call ChangeLine(lngCm / 2)
                End If
                'Footnote 출력
                If Trim(.FootNoteFg) <> "0" Then
                    strBuffer = GetFootNote(.WorkArea, .AccDt, Val(.AccSeq))
                    strBuffer = Replace(strBuffer, vbLf, "")
                    If strBuffer <> "" Then
                        Printer.FontBold = True
                        Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4, "Footnote : ", iCurY, 0)
                        Printer.FontBold = False
                        While (strBuffer <> "")
                            Call WriteStr(iCurY, iposSpcNm + lngCm * 0.4 + Printer.TextWidth("Footnote :   "), medShift(strBuffer, vbCr), iCurY, lngCm / 2)
                            CheckNewPage
                        Wend
                        Call ChangeLine(lngCm / 2)
                    End If
                    
                End If
            End If
            
        End With
    Next
        
    Printer.EndDoc
    
End Sub

Private Sub CheckNewPage()

    If iCurY > iPageHeight - 2 * lngCm Then ' newPage일 경우
        Printer.NewPage
        Call PrtHeader
    End If
            
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
    
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + lngCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + lngCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + lngCm / 6
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


Private Function GetFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
   
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


'% 감수성결과 조회
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
    
   Dim MySenResult() As SenResult
   
   Const MaxAntiCnt = 50
   
   GetSenResult = ""
   SenClipText = ""
   
   SqlStmt = MySql.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs.EOF Then GoTo Nodata
   
   MicroCnt = tmpRs.RecordCount  '균갯수
   ReDim MySenResult(MaxAntiCnt, MicroCnt)
   
   Set AntiList = frmControls.lstUnsortedList   '항생제 리스트(Unsorted)
   Set AntiSortList = frmControls.lstList   '항생제 리스트(Unsorted)
   AntiList.Clear
   AntiSortList.Clear
   
   tmpAntiCnt = 0
   tmpTitle = "Antibiotics  " & vbTab
   
   '미생물 감수성 결과내역 Buffering (LAB405)
   For i = 1 To MicroCnt
      
      tmpTitle = tmpTitle & " " & CStr(i) & " "
      GrowthQty = Trim("" & tmpRs.Fields("GrowthQty").Value)   '배양정도
      SenClipText = SenClipText & Trim("" & tmpRs.Fields("Seq").Value) & "  "
      If Len(GrowthQty) > 6 Then
        SenClipText = SenClipText & Mid(GrowthQty, 1, 6) & vbTab '배양정도
      Else
        SenClipText = SenClipText & GrowthQty & Space(6 - Len(GrowthQty)) & vbTab '배양정도
      End If
      SenClipText = SenClipText & Trim("" & tmpRs.Fields("MicroNm").Value) & vbCrLf   '균명
      
      AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '항생제 갯수
      If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
      
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
   
   SenClipText = SenClipText & vbCrLf
   
   If tmpAntiCnt > 0 Then  '감수성결과가 있으면...
      SenClipText = SenClipText & "[ Susceptibility test ]" & vbCrLf
      SenClipText = SenClipText & tmpTitle & vbCrLf
   End If

   '감수성 결과 Buffering
   Dim mSenRst As New Collection
   
   For i = 1 To AntiList.ListCount
      AntiCd = AntiList.List(i - 1)
      Set tmpRs1 = Nothing
      Set tmpRs1 = New Recordset
      tmpRs1.Open MyCommSql.SqlCommonCode(T_LAB032, LC3_AntiBiotic, AntiCd), DBConn
      AntiNm = Trim("" & tmpRs1.Fields("Field1").Value)     '항생제명
      
      Set SenRstBuf = New clsResults
      SenRstBuf.AntiCd = AntiCd
      SenRstBuf.AntiNm = AntiNm
      AntiSortList.AddItem AntiNm   '항생제명 순으로 정렬리스트에 추가 ...
'      tmpRs1.RsClose
      For j = 1 To MicroCnt
         If MySenResult(i - 1, j).RstCd = "" Then
            SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & "   "
         Else
            SenRstBuf.SenRstBuf = SenRstBuf.SenRstBuf & " " & MySenResult(i - 1, j).RstCd & " "  '균별 감수성결과
         End If
      Next
      mSenRst.Add SenRstBuf, AntiNm
   Next
   
   '균코드
   For i = 1 To AntiSortList.ListCount
      AntiNm = AntiSortList.List(i - 1)
      If Len(AntiNm) > 11 Then
        SenClipText = SenClipText & Mid(AntiNm, 1, 11) & vbTab '항생제명
      Else
        SenClipText = SenClipText & AntiNm & Space(11 - Len(AntiNm)) & vbTab  '항생제명
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
   
End Function
