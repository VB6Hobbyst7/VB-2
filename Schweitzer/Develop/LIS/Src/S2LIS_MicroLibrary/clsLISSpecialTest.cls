VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSpecialTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objSQL As New clsLISSqlETest


Public Sub LoadRstTemplate(ByVal strKey As String, ByRef lvwResult As Object)
    
    Dim objRs As Recordset
    Dim objComSql As New clsLISSqlStatement
    Dim itmX As Object
    
    lvwResult.ListItems.Clear
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeAddTemp, "*", strKey), DBConn
    
    With objRs
        While Not .EOF
            Set itmX = lvwResult.ListItems.Add(, , "" & .Fields("text1").Value)
            itmX.SubItems(1) = .Fields("cdval2").Value & ""
            .MoveNext
        Wend
    End With
    Set objRs = Nothing
    Set objComSql = Nothing

End Sub


Public Sub LoadRstFields(ByRef dicRstFields As Object)

    Dim objRs As Recordset
    Dim objComSql As New clsLISSqlStatement
    
    dicRstFields.Clear
    dicRstFields.DeleteAll
    dicRstFields.FieldInialize "stcd,tpcd", "valuecnt,rstfields"
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeTemp, "*"), DBConn
    
    With objRs
        While Not .EOF
            dicRstFields.AddNew "" & .Fields("cdval1").Value & COL_DIV & .Fields("cdval2").Value, _
                                "" & .Fields("field1").Value & COL_DIV & .Fields("text1").Value
            .MoveNext
        Wend
    End With
    Set objRs = Nothing
    Set objComSql = Nothing
End Sub


Public Function GetAppendText(ByVal pTestCd As String, ByVal pApdCd As String) As String

    Dim dsApd As Recordset
    
    Set dsApd = New Recordset
    dsApd.Open objSQL.SqlGetAppendText(pTestCd, pApdCd), DBConn
    
    If Not dsApd.EOF Then
        GetAppendText = Trim("" & dsApd.Fields("text1").Value & vbCrLf)
    Else
        GetAppendText = ""
    End If
    Set dsApd = Nothing
    
End Function

Public Function GetRemark(ByVal pRmkCd As String) As String
    
    Dim dsRemark As Recordset
    
    Set dsRemark = New Recordset
    dsRemark.Open objSQL.SqlGetRemark(pRmkCd), DBConn
    
    If Not dsRemark.EOF Then
        GetRemark = Trim("" & dsRemark.Fields("text1").Value)
    Else
        GetRemark = ""
    End If
    Set dsRemark = Nothing

End Function



Public Function GetTemplateRst(ByVal sRType As String, ByVal sTCode As String) As String

    Dim dsTxt As Recordset

    Set dsTxt = New Recordset
    dsTxt.Open objSQL.SqlGetTemplateRst(sRType, sTCode), DBConn
    
    If Not dsTxt.EOF Then
        GetTemplateRst = Trim("" & dsTxt.Fields("tpdata").Value)
    Else
        GetTemplateRst = ""
    End If
    Set dsTxt = Nothing

End Function


Public Sub LoadTstNo(ByRef lstTstNo As Object)
    
    Dim dsTN As New Recordset

    Set dsTN = New Recordset
    dsTN.Open objSQL.SqlLoadTestList, DBConn
    
    lstTstNo.Clear
    If Not dsTN.EOF Then
        While Not dsTN.EOF
           lstTstNo.AddItem "" & dsTN.Fields("testcd").Value & vbTab & _
                            "" & dsTN.Fields("testnm").Value & Space(20) & vbTab & _
                            "" & dsTN.Fields("rsttype").Value
           dsTN.MoveNext
        Wend
    End If
         
    Set dsTN = Nothing

End Sub

Public Sub LoadRemark(ByRef cboRemark As Object)
    
    Dim i As Integer
    Dim sqlRM As String, dsRM As Recordset

    ' 검체 Remark
    Set dsRM = New Recordset
    dsRM.Open objSQL.SqlGetRemark, DBConn

    cboRemark.Clear
    cboRemark.AddItem LIS_Nothing
    
    If Not dsRM.EOF Then
    
        For i = 1 To dsRM.RecordCount
            cboRemark.AddItem "" & dsRM.Fields("cdval1").Value
            dsRM.MoveNext
        Next i
    
    End If
    
    Set dsRM = Nothing

End Sub

Public Function LoadResultByLabNo(ByVal sTestCd As String, ByVal sWorkArea As String, _
                                  ByVal sAccDt As String, ByVal sAccSeq As String, _
                                  ByRef sRstFg As String, ByRef iMfySeq As Long) As String
    
    Dim dsTest As Recordset
    
    Set dsTest = New Recordset
    dsTest.Open objSQL.SqlLoadResultByLabNo(sTestCd, sWorkArea, sAccDt, sAccSeq), DBConn
    
    If Not dsTest.EOF Then
        sRstFg = "" & dsTest.Fields("stscd").Value
        iMfySeq = Val("" & dsTest.Fields("mfyseq").Value)
    End If
    Set dsTest = Nothing
    
End Function


Public Sub SetTemplate(ByVal pRstType As String, ByRef cboTemplate As Object)
    
    Dim sqlTemplate As String, dsTemplate As Recordset

    Set dsTemplate = New Recordset
    dsTemplate.Open objSQL.SqlGetTemplateRst(pRstType), DBConn
    
    cboTemplate.Clear
    If dsTemplate.EOF Then Set dsTemplate = Nothing: Exit Sub
    
    Do Until dsTemplate.EOF
        cboTemplate.AddItem "" & dsTemplate.Fields("tpcd").Value & vbTab & _
                            "" & dsTemplate.Fields("tpnm").Value & vbTab & vbTab & pRstType
        dsTemplate.MoveNext
    Loop
    
    Set dsTemplate = Nothing
    
End Sub


Public Sub SetAppend(ByVal pTestCd As String, ByRef cboAppend As Object)
    
    Dim sqlAppend As String, dsAppend As Recordset

    If Trim(pTestCd) = "" Then Exit Sub

    Set dsAppend = New Recordset
    dsAppend.Open objSQL.SqlGetAppendText(pTestCd), DBConn
    
    cboAppend.Clear
    If dsAppend.EOF Then Set dsAppend = Nothing: Exit Sub
    
    dsAppend.MoveFirst
    Do Until dsAppend.EOF
        cboAppend.AddItem "" & dsAppend.Fields("cdval2").Value & vbTab & "" & dsAppend.Fields("field1").Value
        dsAppend.MoveNext
    Loop
    
    Set dsAppend = Nothing
    
End Sub


Public Sub LoadAppendTemp(ByVal pRstType As String, ByRef cboAppend As Object)
    
    Dim objRs As Recordset
    Dim iCnt As Long
    Dim objComSql As New clsLISSqlStatement
    
    cboAppend.Clear
    
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeTemp, "*", pRstType), DBConn
    
    With objRs
        While Not .EOF
            For iCnt = 1 To Val(.Fields("field1").Value)
                cboAppend.AddItem medGetP("" & .Fields("text1").Value, iCnt, vbTab)
            Next
            .MoveNext
        Wend
    End With
    Set objRs = Nothing
End Sub

Public Function GetResultText(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              Optional ByVal pMfySeq As String = "") As String

    Dim rsRst As Recordset
    Dim objRichTemp As RichTextBox
    Dim objRichText As RichTextBox
    Dim blnFirst As Boolean
    
    Set objRichTemp = frmControls.rtfTempText
    Set objRichText = frmControls.rtfTextBox
    
    objRichTemp.Text = ""
    objRichText.Text = ""
    
    Set rsRst = New Recordset
    rsRst.Open objSQL.SqlGetResultText(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq), DBConn

    blnFirst = True
    GetResultText = ""
    If Not rsRst.EOF Then
        rsRst.MoveFirst
        Do Until rsRst.EOF
            If blnFirst Then
                objRichText.TextRTF = "" & rsRst.Fields("txtrst").Value
                blnFirst = False
            Else
                objRichTemp.TextRTF = "" & rsRst.Fields("txtrst").Value
                objRichText.Text = objRichText.Text & vbCrLf & objRichTemp.Text
            End If
            rsRst.MoveNext
        Loop
    End If

    Set rsRst = Nothing
    GetResultText = objRichText.TextRTF
    
End Function
   
Public Function GetSpeResultText(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              ByRef pRstCd1 As String, ByRef pRstCd2 As String, _
                              ByRef pRstCd3 As String, Optional ByVal pMfySeq As String = "") As String

    Dim rsRst As Recordset
    Dim objRichTemp As RichTextBox
    Dim objRichText As RichTextBox
    Dim blnFirst As Boolean
    
    Set objRichTemp = frmControls.rtfTempText
    Set objRichText = frmControls.rtfTextBox
    
    objRichTemp.Text = ""
    objRichText.Text = ""
    
    Set rsRst = New Recordset
    rsRst.Open objSQL.SqlGetResultText(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq), DBConn

    blnFirst = True
    GetSpeResultText = ""
    If Not rsRst.EOF Then
        rsRst.MoveFirst
        
        pRstCd1 = "" & rsRst.Fields("rstcd1").Value
        pRstCd2 = "" & rsRst.Fields("rstcd2").Value
        pRstCd3 = "" & rsRst.Fields("rstcd3").Value
        
        Do Until rsRst.EOF
            If blnFirst Then
                objRichText.TextRTF = "" & rsRst.Fields("txtrst").Value
                blnFirst = False
            Else
                objRichTemp.TextRTF = "" & rsRst.Fields("txtrst").Value
                objRichText.Text = objRichText.Text & vbCrLf & objRichTemp.Text
            End If
            rsRst.MoveNext
        Loop
    End If

    Set rsRst = Nothing
    GetSpeResultText = objRichText.TextRTF
    
End Function
   
   
Public Function DisplayRelTest(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByRef objSS As Object) As Boolean

    Dim SqlStmt As String

    Dim tmpRs As New Recordset
    Dim tmpRs1 As New Recordset
    Dim tmpHLDiv As String
    Dim i As Long
    Dim strRstDiv As String
    Dim strDetailFg As String
    
    Dim tmpTestCd As String, tmpSpcCd As String
    Dim tmpSex As String, tmpAgeDay As String, tmpVfyDt As String
    Dim strRefCd As String
    Dim dblRefFromVal As Double, dblRefToVal As Double
    
    Dim objRstSql As New clsLISSqlReview
   
    SqlStmt = objSQL.SqlGetRelTest(pWorkArea, pAccDt, pAccSeq)

    tmpRs.Open SqlStmt, DBConn
   
    objSS.MaxRows = 0
    If tmpRs.EOF Then
        DisplayRelTest = False
        GoTo NoData
    End If
   
    With objSS
        DisplayRelTest = True
        
        objSS.MaxRows = tmpRs.RecordCount
        For i = 1 To tmpRs.RecordCount
            .Row = i
            
            strRstDiv = "" & tmpRs.Fields("RstDiv").Value
            strDetailFg = "" & tmpRs.Fields("DetailFg").Value
            
            .Col = 1:
                If strRstDiv <> "*" And strDetailFg <> "" Then
                    .Value = "    " & tmpRs.Fields("TestNm").Value
                Else
                    .Value = "" & tmpRs.Fields("TestNm").Value
                End If
            .Col = 2: .Value = "" & tmpRs.Fields("RstCd").Value: .ForeColor = &H404080
            .Col = 3: .Value = "" & tmpRs.Fields("RstUnit").Value
            .Col = 4
                tmpHLDiv = "" & tmpRs.Fields("hldiv").Value
                If tmpHLDiv = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed  '&H7477EF '약간 붉은색
                If tmpHLDiv = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG: .ForeColor = DCM_LightBlue   '&HE48372 '약간 파란색
            .Col = 5: .Value = "" & tmpRs.Fields("dpdiv").Value
            
            '기준치 검색
            tmpTestCd = "" & tmpRs.Fields("TestCd").Value
            tmpSpcCd = "" & tmpRs.Fields("SpcCd").Value
            tmpSex = "" & tmpRs.Fields("Sex").Value
            tmpAgeDay = "" & tmpRs.Fields("AgeDay").Value
            tmpVfyDt = "" & tmpRs.Fields("VfyDt").Value
            
'            If tmpTestCd = "B109108" Then
'                Stop
'            End If
            SqlStmt = objRstSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
            If tmpRs1.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
                SqlStmt = objRstSql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                Set tmpRs1 = New Recordset
                tmpRs1.Open SqlStmt, DBConn
            End If
            If tmpRs1.EOF Then
                strRefCd = ""
            Else
                dblRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                dblRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                If dblRefFromVal <> 0 Or dblRefToVal <> 0 Then
                    strRefCd = dblRefFromVal & " - " & dblRefToVal
                Else
                    strRefCd = ""
                End If
            End If
            Set tmpRs1 = Nothing
            .Col = 6: .Value = strRefCd: .ForeColor = &H8000&
            
            .Col = 7: .Value = Format("" & tmpRs.Fields("ColDt").Value, CS_DateLongMask) & " " & _
                               Format("" & tmpRs.Fields("ColTm").Value, CS_TimeLongMask)
            .Col = 8: .Value = Format("" & tmpRs.Fields("VfyDt").Value, CS_DateLongMask) & " " & _
                               Format("" & tmpRs.Fields("VfyTm").Value, CS_TimeLongMask)
            
            tmpRs.MoveNext

        Next
    End With
   
NoData:
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Function

  

Public Function GetLabNoList(ByVal pTestCd As String, ByRef lstLabNo As Object)

    Dim dsLN As New Recordset
    Dim iLNCol As Integer
    Dim sAccDt As String
    
    dsLN.Open objSQL.SqlGetLabNoList(pTestCd), DBConn
    
    lstLabNo.Clear
    Do Until dsLN.EOF
        sAccDt = "" & dsLN.Fields("accdt").Value
        lstLabNo.AddItem dsLN.Fields("workarea").Value & "-" & Mid$(sAccDt, 3, Len(sAccDt) - 2) & "-" & _
                         dsLN.Fields("accseq").Value & " " & dsLN.Fields("ptid").Value
        dsLN.MoveNext
    Loop
      
    Set dsLN = Nothing
   
End Function


Public Function GetAccList(ByVal pTestCd As String, ByRef lstLabNo As Object)

    Dim dsLN As New Recordset
    Dim iLNCol As Integer
    Dim sAccDt As String
    
    dsLN.Open objSQL.SqlGetAccList(pTestCd), DBConn

    lstLabNo.Clear
    Do Until dsLN.EOF
        sAccDt = "" & dsLN.Fields("accdt").Value
        lstLabNo.AddItem dsLN.Fields("workarea").Value & "-" & Mid$(sAccDt, 3, Len(sAccDt) - 2) & "-" & _
                         dsLN.Fields("accseq").Value & " " & dsLN.Fields("ptid").Value
        dsLN.MoveNext
    Loop
      
    Set dsLN = Nothing
   
End Function


Public Function ReadFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim sqlFNote As String, dsFNote As New Recordset, iFNCol As Integer

    ReadFootNote = ""
    
    dsFNote.Open objSQL.SqlReadFootNote(pWorkArea, pAccDt, pAccSeq), DBConn
    
    Do Until dsFNote.EOF
        ReadFootNote = ReadFootNote & dsFNote.Fields("rsttxt").Value & "" & vbCrLf
        dsFNote.MoveNext
    Loop
    
    Set dsFNote = Nothing
End Function



Public Function SetHistory(ByVal pWorkArea As String, ByVal pAccDt As String, _
                           ByVal pAccSeq As String, ByVal pTestCd As String) As Boolean
    
    Dim sCurRst As String, dsCurRst As Recordset
    Dim sqlDeAccSeq54 As String
    Dim sHead As String, sData As String, sqlIns354 As String

    ' 현재 변경 대상 Data Read
    Set dsCurRst = New Recordset
    dsCurRst.Open objSQL.SqlGetCurrRst(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
    
    sqlIns354 = ""
    If Not dsCurRst.EOF Then
        sHead = "(workarea,accdt,accseq,testcd,mfyseq,valfg,txtfg,vfydt,vfytm,vfyid)"
        sData = "(" & DBV("workarea", pWorkArea, 1) & _
                      DBV("accdt", pAccDt, 1) & _
                      DBV("accseq", pAccSeq, 1) & _
                      DBV("testcd", pTestCd, 1) & _
                      DBV("mfyseq", "" & dsCurRst.Fields("mfyseq").Value, 1) & _
                      DBV("valfg", "" & dsCurRst.Fields("ValFg").Value, 1) & _
                      DBV("txtfg", "" & dsCurRst.Fields("TxtFg").Value, 1) & _
                      DBV("vfydt", "" & dsCurRst.Fields("VfyDt").Value, 1) & _
                      DBV("vfytm", "" & dsCurRst.Fields("VfyTm").Value, 1) & _
                      DBV("vfyid", "" & dsCurRst.Fields("VfyId").Value) & ")"

        sqlIns354 = "INSERT INTO " & T_LAB354 & " " & sHead & " VALUES " & sData
    End If
    Set dsCurRst = Nothing
    
On Error GoTo Err_Trap
    If Trim(sqlIns354) <> "" Then DBConn.Execute sqlIns354
    SetHistory = True
    
    Exit Function

Err_Trap:
    SetHistory = False

End Function


Public Function SaveStatus(ByVal pWorkArea As String, ByVal pAccDt As String, _
                           ByVal pAccSeq As String, ByVal pTestCd As String, _
                           ByVal pStatus As String, ByVal pMfyCnt As Integer, ByVal pValRst As String, _
                           ByVal pTxtRst As String, ByVal pDate As String, _
                           ByVal pTime As String, ByVal pEmpId As String) As Boolean

    Dim sqlUpRst As String, sqlUpOrd As String
    Dim objRs As Recordset
    Dim strPtId As String, strOrdDt As String, strOrdNo As String, strOrdSeq As String
    
    ' 등록과 수정 모두에 공통 적용
On Error GoTo Err_Trap

    sqlUpRst = objSQL.SqlUpdateETest(pWorkArea, pAccDt, pAccSeq, pTestCd, pStatus, _
                                     pMfyCnt, pValRst, pTxtRst, pDate, pTime, pEmpId)
    DBConn.Execute sqlUpRst
    
    If pStatus >= enStsCd.StsCd_LIS_FinRst Then
        'sqlUpOrd = objSql.SQLUpdateOrder(pWorkArea, pAccDt, pAccSeq, pTestCd, pStatus, pDate, pTime, pEmpId)
        sqlUpOrd = objSQL.SqlGetOrderSeq(pWorkArea, pAccDt, pAccSeq, pTestCd)
        Set objRs = New Recordset
        objRs.Open sqlUpOrd, DBConn
        
        If Not objRs.EOF Then
            strPtId = "" & objRs.Fields("ptid").Value
            strOrdDt = "" & objRs.Fields("orddt").Value
            strOrdNo = "" & objRs.Fields("ordno").Value
            strOrdSeq = "" & objRs.Fields("ordseq").Value
            sqlUpOrd = objSQL.SqlSetOrdStatus(strPtId, strOrdDt, strOrdNo, strOrdSeq, _
                                              pStatus, pDate, pTime, pEmpId)
            
            DBConn.Execute sqlUpOrd
        End If
        Set objRs = Nothing
    End If
    SaveStatus = True
    Exit Function
    
Err_Trap:
    SaveStatus = False
    
End Function

Public Function SaveAccStatus(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pStatus As String, _
                              ByVal pDate As String, ByVal pTime As String, _
                              ByVal pEmpId As String) As Boolean
    
    Dim dsAcc As Recordset
    Dim iTotalCount As Integer, iInputCount As Integer, sStatus As String
    Dim sqlUpdAcc As String
    
    SaveAccStatus = True
    
On Error GoTo Err_Trap:

    Set dsAcc = New Recordset
    dsAcc.Open objSQL.SqlGetTestCount(pWorkArea, pAccDt, pAccSeq), DBConn
               
    iTotalCount = Val("" & dsAcc.Fields("reqtotcnt").Value)
    iInputCount = Val("" & dsAcc.Fields("reqinputcnt").Value)
    sStatus = "" & dsAcc.Fields("stscd").Value
    
    Set dsAcc = Nothing
    
    If iTotalCount <= iInputCount + 1 Then sStatus = pStatus                    ' 먼저 체크
    If pStatus = enStsCd.StsCd_LIS_FinRst Then iInputCount = iInputCount + 1    ' 다음 체크 (순서 중요)
        
    ' 접수 내역에 Status 및 조회정보 반영
    If pStatus >= enStsCd.StsCd_LIS_FinRst Then
        If sStatus >= enStsCd.StsCd_LIS_FinRst Then
            sqlUpdAcc = objSQL.SqlUpdateAccStatus(pWorkArea, pAccDt, pAccSeq, pStatus, _
                                                  iInputCount, pDate, pTime, pEmpId)
        Else
            sqlUpdAcc = objSQL.SqlUpdateAccStatus(pWorkArea, pAccDt, pAccSeq, sStatus, _
                                                  iInputCount)
        End If
        DBConn.Execute sqlUpdAcc
    End If
    Exit Function
    
Err_Trap:
    SaveAccStatus = False

End Function



Public Function SaveValResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              ByVal pMfySeq As String, ByVal pResult As String) As Boolean

On Error GoTo Err_Trap
    
    ' 기존 수치 결과 삭제 (수정시 fMfySeq Add 1)
    DBConn.Execute objSQL.SqlSaveValResult(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq, pResult, "1")

    ' 저장할 수치결과 검색
    DBConn.Execute objSQL.SqlSaveValResult(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq, pResult, "2")
    
    SaveValResult = True
    Exit Function

Err_Trap:
    SaveValResult = False

End Function

Public Function SaveSpeResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              ByVal pMfySeq As String, ByVal pResult As String, _
                              ByVal pRstCd1 As String, ByVal pRstCd2 As String, _
                              ByVal pRstCd3 As String) As Boolean

On Error GoTo Err_Trap
    
    ' 기존 수치 결과 삭제 (수정시 fMfySeq Add 1)
    DBConn.Execute objSQL.SqlSaveSpeResult(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq, pResult, pRstCd1, pRstCd2, pRstCd3, "1")

    ' 저장할 수치결과 검색
    DBConn.Execute objSQL.SqlSaveSpeResult(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq, pResult, pRstCd1, pRstCd2, pRstCd3, "2")
    
    SaveSpeResult = True
    Exit Function

Err_Trap:
    SaveSpeResult = False

End Function

Public Function SaveFootnote(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pEmpId As String, _
                              ByVal pFootNote As String, ByVal pRmkCd As String, _
                              ByVal pStatus As String, ByVal fFNSeq As Long) As Boolean
    
    Dim dsFN As Recordset
    Dim pFSeq As Long
    Dim sMfyNo As Long
    
On Error GoTo Err_Trap

    SaveFootnote = True
    
    'Foot Note 결과 저장( 저장 버튼이거나 처음 footnote기록이면 기존 footnote 일단 제거 )
    If pStatus < enStsCd.StsCd_LIS_FinRst Or fFNSeq = 0 Then
        DBConn.Execute objSQL.SqlDeleteFootnote(pWorkArea, pAccDt, pAccSeq)
    End If

    If Trim(pFootNote) = "" Then Exit Function
    
    Set dsFN = New Recordset
    dsFN.Open objSQL.SqlGetMfySeq(pWorkArea, pAccDt, pAccSeq), DBConn
    
    If dsFN.EOF Then
       sMfyNo = 1
    Else
       sMfyNo = Val("" & dsFN.Fields("mfyno").Value) + 1
    End If

    Set dsFN = Nothing

    DBConn.Execute objSQL.SqlInsertFootnote(pWorkArea, pAccDt, pAccSeq, sMfyNo, pEmpId, pFootNote)

    If pRmkCd = LIS_Nothing Then pRmkCd = ""

    ' 접수 내역에 Footnote Seq 및 검체 Remark Code 등록
    If pStatus < enStsCd.StsCd_LIS_Modify Then pFSeq = 1
    If pStatus >= enStsCd.StsCd_LIS_Modify Then pFSeq = 2

    If pStatus >= enStsCd.StsCd_LIS_MidRst Then
        DBConn.Execute objSQL.SqlUpdateAccData(pWorkArea, pAccDt, pAccSeq, pFSeq, pRmkCd)
    End If
    Exit Function

Err_Trap:
    SaveFootnote = False

End Function


Public Function SubmitVerifyList(ByVal pDeptCd As String, ByVal pDate As String, ByVal pTime As String, _
                                 ByVal pPtId As String, ByVal pStatus As String, ByVal pEmpId As String, _
                                 ByVal pMajDoct As String, ByVal pBussDiv As String) As Boolean
    
    
    Dim sSubmit As String, dsSubmit As Recordset
    Dim sMfyFg As String, sExpire As String

On Error GoTo Err_Trap

    SubmitVerifyList = True
    
    If pStatus < enStsCd.StsCd_LIS_Modify Then sMfyFg = "0"
    If pStatus >= enStsCd.StsCd_LIS_Modify Then sMfyFg = "1"
    
    Set dsSubmit = New Recordset
    dsSubmit.Open objSQL.SqlReadVerifyList(pDeptCd, pDate, pPtId, sMfyFg), DBConn
     
    If dsSubmit.RecordCount >= 1 Then
    
        sExpire = "" & dsSubmit.Fields("donefg").Value
        
        Set dsSubmit = Nothing
        
        If sExpire = "1" Then
        ' 갱신 루틴
           DBConn.Execute objSQL.SqlUpdateVerifyList(pDeptCd, pDate, pTime, pPtId, sMfyFg, pEmpId)
        Else
        ' Skip 루틴
           Exit Function
        End If
      
    Else
         
        Set dsSubmit = Nothing
        
        ' 삽입 루틴
        DBConn.Execute objSQL.SqlInsertVerifyList(pDeptCd, pDate, pTime, pPtId, sMfyFg, pEmpId, pMajDoct, pBussDiv)
      
    End If
    Exit Function

Err_Trap:
    SubmitVerifyList = False

End Function

Public Function GetVfyDate(ByVal sWorkArea As String, ByVal sAccDt As String, ByVal sAccSeq As String) As String

    Dim dsTxt As Recordset

    Set dsTxt = New Recordset
    dsTxt.Open objSQL.SqlGetVfyDate(sWorkArea, sAccDt, sAccSeq), DBConn
    
    If Not dsTxt.EOF Then
        GetVfyDate = Trim("" & dsTxt.Fields("vfydt").Value) ' & Trim("" & dsTxt.Fields("vfytm").Value)
    Else
        GetVfyDate = ""
    End If
    Set dsTxt = Nothing

End Function



