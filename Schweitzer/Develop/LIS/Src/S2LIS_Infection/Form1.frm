VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'
'Private WithEvents clsTemplete  As frm230TempSearch
'Private WithEvents objCodeList  As clsCodeList
'
'Private MyOrder As New clsICSCom
'
'Dim WithEvents mnuPopup     As Menu
'Dim WithEvents mnuResult    As Menu
'Dim WithEvents mnuResult1   As Menu
'Dim WithEvents mnuRemark    As Menu
'Private blnChk              As Boolean
'
'Private Enum TblCol
'    tcNO = 1
'    tcLABNO
'    tcPTID
'    tcSpc
'    tcSTAT
'    tcHold
'End Enum
'
'
'Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
'
'    Dim strTitle As String
'
'    Set clsTemplete = New frm230TempSearch
'    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
'
'    With clsTemplete
'        .Show
'        If pintMode = 0 Then
'            .lblNAME = "Edit " & strTitle
'        Else
'            .lblNAME = "Modify " & strTitle
'        End If
'        .Caption = strTitle & " " & "Templete Editor"
'        .lblInfo.Caption = pintMode & "$" & pintPrg
'        .rtfText = rtfComment.Text
''
''        SELECT Case pintPrg
''            Case 1:
''                .lblCode.Caption = objPtInfo.RmkCd
''                .rtfText = rtfRemark.Text
''            Case 2:
''                .rtfText = rtfText.Text
''            Case 3:
''                .rtfText = rtfComment.Text
''        End SELECT
'    End With
'    gintTemplete = pintPrg
'
'End Sub
'Private Sub cmdCommentTemplete_Click()
'    Call CallTemplete(3, 0)
'End Sub
'Private Sub cmdRemarkTemplete_Click()
'
'    Dim SqlStmt As String
'
'    Set objCodeList = Nothing
'    Set objCodeList = New clsCodeList
'
'    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " where  " & DBW("cdindex =", LC4_Remark)
'    With objCodeList
'        Set .MyDB = DBConn
'        .ListCaption = "Remark"
'        .ListColHeader = "Code" & vbTab & "Remark"
'        .Top = 3700 'Me.cmdRemarkTemplete.Top + 5600
'        .Left = 2050
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< �� �� > ", 2, 1
'    End With
'
'End Sub
'
'Private Sub cmdOK_Click()
'    Dim sLabNo  As String
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblCol.tcLABNO: sLabNo = .Value
'        objResult.MoveFirst
'        Do Until objResult.EOF
'            If objResult.Fields("labno") = sLabNo Then
'                objResult.Fields("footnote") = rtfComment.Text
'                objResult.Fields("rmkcd") = lblRmkcd.Caption
'                objResult.Fields("rmknm") = rtfRemark.Text
'            End If
'            objResult.MoveNext
'        Loop
'        If rtfComment.Text <> "" Or lblRmkcd.Caption <> "" Then
'            .FontBold = True: .ForeColor = vbRed
'        Else
'            .FontBold = False: .ForeColor = vbBlack
'        End If
'    End With
'    rtfComment.Text = ""
'    rtfRemark.Text = ""
'    lblRmkcd.Caption = ""
'    fraComment.Visible = False
'End Sub
'Private Sub clsTemplete_CopyTemplete()
'   '
''    If ssRst.MaxRows < 1 Then Exit Sub
'
''    With objPtInfo
'        Select Case gintTemplete
'            Case 1:
''                If clsTemplete.rtfText.Text <> "" Then
''                    rtfRemark.Text = clsTemplete.rtfText.Text
''                    .RmkCd = frm230TempSearch.lblCode.Caption
''                    .RmkNm = rtfRemark.Text
''                Else
''                    rtfRemark.Text = ""
''                    .RmkCd = ""
''                    .RmkNm = ""
''                End If
''            Case 2:
''                rtfText.Text = clsTemplete.rtfText.Text
''                .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
''                rtfText.SetFocus
'            Case 3:
'                rtfComment.Text = clsTemplete.rtfText.Text
''                .FootNote = rtfComment.Text
'                rtfComment.SetFocus
'        End Select
''    End With
'    Set clsTemplete = Nothing
'End Sub
'
'Private Sub cmdClear_Click()
'    Call FormClear
'    txtEqpCd.Text = ""
'End Sub
'
'Private Sub lblColDt_Click()
'
'End Sub
'
'Private Sub mnuRemark_Click()
'    rtfComment.Text = ""
'    rtfRemark.Text = ""
'    fraComment.Visible = True
'    cmdCommentTemplete.Enabled = True
'    cmdRemarkTemplete.Enabled = True
'    lblRmkcd.Caption = ""
'    With tblResult
'        .Row = .ActiveRow: .Col = TblCol.tcLABNO
'        lblRemark.Caption = "Comment by Accession No (" & .Value & ")"
'        objResult.MoveFirst
'        Do Until objResult.EOF
'            If objResult.Fields("labno") = .Value Then
'                rtfComment.Text = objResult.Fields("footnote")
'                rtfRemark.Text = objResult.Fields("rmknm")
'                lblRmkcd.Caption = objResult.Fields("rmkcd")
'            End If
'            objResult.MoveNext
'        Loop
'
'    End With
'End Sub
'
'Private Sub mnuResult_Click()
'    Dim objTestCd   As New clsDictionary
'
'    Dim ii          As Integer
'
'    Dim sTestcd   As String
'    Dim sSpcCd    As String
'    Dim sLabNo      As String
'    Dim sPtid     As String
'
'    objTestCd.clear
'    objTestCd.FieldInialize "testcd", "spccd"
'
'    objTestCd.Sort = False
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblCol.tcLABNO: sLabNo = .Value
'        .Col = TblCol.tcPTID:  sPtid = .Value
'        objResult.MoveFirst
'        Do Until objResult.EOF
'            If objResult.Fields("labno") = sLabNo Then
'                If objTestCd.Exists(objResult.Fields("testcd")) = False Then
'                    objTestCd.AddNew objResult.Fields("testcd"), objResult.Fields("spccd")
'                End If
'            End If
'            objResult.MoveNext
'        Loop
'    End With
'
'    objTestCd.Sort = True
'
'    Set objTestCd = Nothing
'
'End Sub
'
'Private Sub mnuResult1_Click()
'    Dim objTestCd   As New clsDictionary
'    Dim ii          As Integer
'    Dim sTestcd     As String
'    Dim sSpcCd      As String
'    Dim sLabNo      As String
'    Dim sPtid       As String
'
'    objTestCd.clear
'    objTestCd.FieldInialize "testcd", "spccd"
'
'    objTestCd.Sort = False
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = .ActiveCol + 1: sTestcd = medGetP(.Value, 1, COL_DIV)
'        .Col = TblCol.tcLABNO: sLabNo = .Value
'        .Col = TblCol.tcPTID:  sPtid = .Value
'
'        objResult.KeyChange sLabNo & COL_DIV & sTestcd
'        sSpcCd = objResult.Fields("spccd")
'
'        objTestCd.AddNew sTestcd, sSpcCd
'
'    End With
'
'    objTestCd.Sort = True
'
'    Set objTestCd = Nothing
'
'End Sub
'
'Private Sub cmdEqp_Click()
'
'    If lstEQCode.ListCount = 0 Then
'        MsgBox "������ ��� �����ϴ�.", vbCritical
'        Exit Sub
'    End If
'
'    lstEQCode.Visible = True
'    Set objCodeList = Nothing
'    lstEQCode.ZOrder 0
'    lstEQCode.SetFocus
'
'End Sub
'
'
'Private Sub FormClear()
'    lblErr.Caption = ""
'    lblPtId.Caption = "": lblPtNm.Caption = "": lblSex.Caption = "": lblDoct.Caption = ""
'    lblWard.Caption = "": lblDept.Caption = "": lblOrdDt.Caption = "": lblColDt.Caption = "": lblRcvDT.Caption = ""
'
'    lblEqpCdNm.Caption = ""
'
'    chkStatFg.Value = 0
'    chkAuto.Value = 0
'
'    txtTransDt.Text = ""
'    lblEqpCdNm.Caption = ""
'    Call medClearTable(tblResult)
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'
'End Sub
'
'Private Sub cmdQuery_Click()
'    Dim objPrgBar       As clsProgressBar
'    Dim Rs              As DrRecordSet
'    Dim sSQL            As String
'    Dim sTransDt        As String
'    Dim sEqpCD          As String
'    Dim sStat           As String
'
'    Dim sValue          As String
'    Dim sRealVal        As String
'
'    Dim strTmp          As String
'
'
'    Dim jj              As Integer
'    Dim ii              As Integer
'    Dim KK              As Integer
'
'    Call medClearTable(tblResult)
'
'    With tblResult
'        .Row = 1: .Row2 = .MaxRows
'        .Col = 1: .Col2 = .MaxCols
'        .BlockMode = True
'        .ForeColor = vbBlack: .FontBold = False
'        .BlockMode = False
'    End With
'
'
'    Call ResultClear
'
'    If Trim(txtEqpCd.Text) = "" Then
'        MsgBox "��� �ڵ带 �Է��Ͻ��� ��ȸ�ϼ���.", vbInformation + vbOKOnly, "Info"
'        Exit Sub
'    End If
'
'    Me.MousePointer = 11
'
'    sEqpCD = Trim(txtEqpCd.Text)
'    sStat = chkStatFg.Value
'    sTransDt = Replace(txtTransDt.Text, "-", "")
'
'    Set Rs = OpenRecordSet(objRst.InstBatchResultQuery(sEqpCD, sTransDt, sStat))
'
'
'    ii = 0
'    If Not Rs.EOF Then
'        Set objPrgBar = New clsProgressBar
'        Set objPrgBar.StatusBar = medMain.stsBar
'        objPrgBar.Min = 1
'        objPrgBar.Max = Rs.RecordCount
'
'        With tblResult
''            .ReDraw = False
'            Do Until Rs.EOF
'                If strTmp <> Rs.Fields("labno").Value & "" And Rs.Fields("rstdiv").Value & "" <> "*" Then
'                    ii = ii + 1
'                    If ii > .MaxRows Then
'                        .MaxRows = .MaxRows + 1
'                        .Row = .MaxRows
'                    Else
'                        .Row = ii
'                    End If
'                    .RowHeight(.Row) = 12
'
'                    .Col = TblCol.tcNO:     .Value = .Row
'                    .Col = TblCol.tcLABNO:  .Value = Rs.Fields("labno").Value & ""
'                    .TypeHAlign = 2:        .TypeVAlign = 2 'TypeVAlignCenter'TypeHAlignCenter
'                    .Col = TblCol.tcPTID:   .Value = Rs.Fields("ptid").Value & ""
'                    .Col = TblCol.tcSTAT:   .Value = IIf(Rs.Fields("statfg").Value & "" = "1", "Y", ""): .ForeColor = DCM_LightRed
'                    .Col = TblCol.tcSpc:    .Value = Rs.Fields("spcnm").Value & ""
'                    .Col = TblCol.tcHold:   .Value = 1
'                    jj = 6
'                End If
'
'                If Rs.Fields("rstdiv").Value <> "*" Then
'                    strTmp = Rs.Fields("labno").Value & ""
'                    jj = jj + 3
'                    If .MaxCols < jj Then
'                        .MaxCols = jj
'                    End If
'
'                    sRealVal = ""
'                    .Row = 0:
'                    .Col = jj - 2: .Value = "�˻��": .ColWidth(jj - 1) = 5.63
'                    .Col = jj - 1: .ColHidden = True
'                    .Col = jj:     .Value = "���":   .ColWidth(jj) = 10
'                    .Row = ii
'                    .Col = jj - 2: .Value = Rs.Fields("abbrnm5").Value & ""
'                    .Col = jj - 1: .Value = Rs.Fields("testcd").Value & "" & COL_DIV & Rs.Fields("spccd").Value & ""
'                    .Col = jj:     .Value = Rs.Fields("rstcd").Value & ""
'                    sValue = Trim(.Value)
'
'                    If sValue <> "" Then
'                        Select Case Rs.Fields("hldiv").Value & ""
'                            Case "H":
'                                .Value = .Value: .ForeColor = DCM_LightRed:  '.FontBold = True
'                            Case "L":
'                                .Value = .Value: .ForeColor = DCM_LightBlue:   '.FontBold = True
'                        End Select
'                        If Rs.Fields("dpdiv").Value & "" <> "" Then
'                            .Value = .Value & " " & Rs.Fields("dpdiv").Value & "": .FontBold = True: .ForeColor = DCM_LightRed:
'                        End If
'                        sRealVal = .Value
'                    End If
'                    .TypeHAlign = 2: .TypeVAlign = 2 'TypeVAlignCenter'TypeHAlignCenter
'                End If
'
'                If objResult.Exists(Rs.Fields("labno").Value & "" & COL_DIV & Rs.Fields("testcd").Value & "") Then
'
'                Else
'                    Call objRst.GetRefVal(Rs.Fields("ptid").Value & "", Rs.Fields("testcd").Value & "", _
'                                          Rs.Fields("spccd").Value & "", Rs.Fields("rcvdt").Value & "")
'                    If Arlet_PanicChk = False Then
'                        objRst.PanicFrVal = Rs.Fields("panicfrval").Value & ""
'                        objRst.PanicToVal = Rs.Fields("panictoval").Value & ""
'                        objRst.ArletFrVal = Rs.Fields("arletfrval").Value & ""
'                        objRst.ArletToVal = Rs.Fields("arlettoval").Value & ""
'                    End If
'
'                    objResult.AddNew Rs.Fields("labno").Value & "" & COL_DIV & Rs.Fields("testcd").Value & "", _
'                                     Rs.Fields("spccd").Value & "" & COL_DIV & Rs.Fields("rstval").Value & "" & COL_DIV & _
'                                     Rs.Fields("rstcd").Value & "" & COL_DIV & Rs.Fields("rsttype").Value & "" & COL_DIV & _
'                                     Rs.Fields("hldiv").Value & "" & COL_DIV & Rs.Fields("dpdiv").Value & "" & COL_DIV & "0" & COL_DIV & _
'                                     Rs.Fields("avalval").Value & "" & COL_DIV & Rs.Fields("ptid").Value & "" & COL_DIV & _
'                                     Rs.Fields("lastrst").Value & "" & COL_DIV & Rs.Fields("lastvfydt").Value & "" & COL_DIV & Rs.Fields("deltafg").Value & "" & COL_DIV & _
'                                     Rs.Fields("deltaval1").Value & "" & COL_DIV & Rs.Fields("deltaval2").Value & "" & COL_DIV & _
'                                     Rs.Fields("panicfg").Value & "" & COL_DIV & objRst.PanicFrVal & COL_DIV & _
'                                     objRst.PanicToVal & COL_DIV & objRst.RefFrVal & COL_DIV & objRst.RefToVal & COL_DIV & sRealVal & COL_DIV & _
'                                     objRst.Sex & "/" & objRst.Age & COL_DIV & Rs.Fields("wardid").Value & "" & "-" & Rs.Fields("hosilid").Value & "" & COL_DIV & _
'                                     Rs.Fields("deptcd").Value & "" & COL_DIV & Rs.Fields("orddoct").Value & "" & COL_DIV & _
'                                     Format(Rs.Fields("coldt").Value & "", "0###-##-##") & " " & Format(Rs.Fields("coltm").Value, "0#:##:##") & COL_DIV & _
'                                     Format(Rs.Fields("rcvdt").Value & "", "0###-##-##") & " " & Format(Rs.Fields("rcvtm").Value, "0#:##:##") & COL_DIV & _
'                                     Format(Rs.Fields("orddt").Value & "", "0###-##-##") & COL_DIV & objRst.ptnm & COL_DIV & Rs.Fields("abbrnm5").Value & "" & COL_DIV & _
'                                     "" & COL_DIV & "" & COL_DIV & "" & COL_DIV & Rs.Fields("eqpcd") & "" & COL_DIV & Rs.Fields("txtfg").Value & "" & COL_DIV & "" & COL_DIV & _
'                                     Rs.Fields("detailfg").Value & "" & COL_DIV & Rs.Fields("ordno").Value & "" & COL_DIV & _
'                                     Rs.Fields("ordseq").Value & "" & COL_DIV & Rs.Fields("rstdiv").Value & "" & COL_DIV & Rs.Fields("rstcd").Value & "" & COL_DIV & "" & COL_DIV & _
'                                     objRst.ARefFrVal & COL_DIV & objRst.ARefToVal & COL_DIV & Rs.Fields("arletfg").Value & "" & COL_DIV & objRst.ArletFrVal & COL_DIV & objRst.ArletToVal
'
'
'
'                End If
'                KK = KK + 1
'                objPrgBar.Value = KK
'                Rs.MoveNext
'            Loop
'            .Row = 1: .Col = TblCol.tcHold + 3: .Action = ActionActiveCell
'            objResult.Sort = True
'            .ReDraw = True
'        End With
'        Call GetDisplay(8, 1)
'    End If
'
'    If objResult.RecordCount > 0 Then
'        Call objRst.SetRST(tblResult, objResult)
'    End If
'    Me.MousePointer = 0
'    Set Rs = Nothing
'    Set objPrgBar = Nothing
'
'
'
'
'
'End Sub
'Private Sub ResultClear()
'
'    Set objResult = Nothing
'    Set objResult = New clsDictionary
'    objResult.clear
'
'    objResult.FieldInialize "labno,testcd", "spccd,rstval,rstcd,rsttype,hldiv,dpdiv,hold,avalval," & _
'                                            "ptid,lastrst,lastvfydt,deltafg,deltaval1,deltaval2,panicfg," & _
'                                            "panicfrval,panictoval,refvalfrom,refvalto,rstcdval," & _
'                                            "sexage,wardid,deptcd,orddoct,coldate,rcvdate,orddt,ptnm,abbrnm5," & _
'                                            "footnote,rmkcd,rmknm,eqpcd,txtfg,rsttext,detailfg,ordno,ordseq,rstdiv," & _
'                                            "saverst,vfydt,afrval,atoval,arletfg,arletfrval,arlettoval"
'
'    objResult.Sort = False
'End Sub
'Private Sub cmdTrans_Click()
'
'    lstEQCode.Visible = False
'
'    '2001-11-07 �߰� : ���� ������۳��� ���� (�Ⱓ : 1����)
'    Screen.MousePointer = vbArrowHourglass
'    lblErr.Caption = "������ ������� ������ �����ϰ� �ֽ��ϴ�."
'    Call objRst.EqpHistoryDelete(txtEqpCd.Text, Format(DateAdd("d", -30, Now), CS_DateDbFormat))
'
'    lblErr.Caption = ""
'    Screen.MousePointer = vbDefault
'
'    TrasferListPop txtEqpCd.Text
'End Sub
'Private Sub TrasferListPop(ByVal EqpCd As String)
'   If EqpCd = "" Then Exit Sub
'   Dim sSQL As String
'
'   Set objCodeList = New clsCodeList
'   With objCodeList
'      Set .MyDB = DBConn
'      .ListCaption = "Instrument List"
'      .ListColHeader = "Name" & vbTab & "Code"
'      .Top = Me.cmdTrans.Top + 2000
'      .Left = Me.cmdTrans.Left - 1270
'      .Width = 3450
'      .Height = 3000
'      .Tag = "Transfer"
'      .CaptionOn = False
'      .MultiSel = False
'      sSQL = objRst.SqlStmtTransfer(EqpCd)
'      .PopupList sSQL, 2
'   End With
'End Sub
'
'Private Sub objCodeList_ListClick(ByVal SelList As String)
'
'   If Not IsNull(SelList) And SelList <> "" Then
'      Select Case objCodeList.Tag
'         Case "Transfer":
'            txtTransDt.Text = Format(medGetP(SelList, 1, vbTab), "0###-##-##")
'            cmdQuery.SetFocus
'        Case "Remark"
'            lblRmkcd.Caption = medGetP(SelList, 1, vbTab)
'            rtfRemark.Text = medGetP(SelList, 2, vbTab)
'    End Select
'   End If
'
'   Set objCodeList = Nothing
'   '
'End Sub
'
'Private Sub Form_Load()
'    '-- DB ����
''    If DBOpen(D0COM_SERVER) = False Then
''        MsgBox Err.Description
''        Set F_SQL = Nothing
''        Unload Me
''    End If
'
'    Call MyOrder.ItemList(lstTestList, barStatus)
'
'    Call FormClear
'    txtEqpCd.Text = ""
'End Sub
'
'Private Sub ShowEqpList()
'    Dim FNo As Long
'    Dim FName As String
'    Dim i As Long
'    Dim strData As String
'    Dim strTemp As String
'
'    FNo = FreeFile
'
'    On Error GoTo ErrList
'
'    Open App.Path & "\LIS.dat" For Input As #FNo
'
'    lstEQCode.clear
'    Do While Not EOF(1)
'        Line Input #FNo, strTemp
'
'        strData = DECrypt(strTemp)
'
'        lstEQCode.AddItem Trim(Mid(strData, 1, 10)) & vbTab & Trim(Mid(strData, 11)) & vbTab, i
'        i = i + 1
'    Loop
'    Close #FNo
'
'    If lstEQCode.ListCount = 0 Then
'        MsgBox "������ ��� �����ϴ�.", vbCritical
'    End If
'
'    Exit Sub
'ErrList:
'    MsgBox Err.Description, vbExclamation
'    Close #FNo
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'    Set objResult = Nothing
'    Set objRst = Nothing
'    Set clsTemplete = Nothing
'    Set objCodeList = Nothing
'    Set mnuPopup = Nothing
'    Set mnuResult = Nothing
'    Set mnuResult1 = Nothing
'    Set mnuRemark = Nothing
'End Sub
'
'Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim i As Integer
'
'    If Row = 0 And Col = TblCol.tcHold Then
'        With tblResult
'            .Col = TblCol.tcHold
'            For i = 1 To .MaxRows
'                .Row = i
'                .Col = TblCol.tcLABNO
'                If .Value = "" Then Exit For
'                If blnChk = False Then
'                    .Col = TblCol.tcHold: .Value = 0
'                Else
'                    .Col = TblCol.tcHold: .Value = 1
'                End If
'            Next
'        End With
'        If blnChk = False Then
'            blnChk = True
'        Else
'            blnChk = False
'        End If
'    End If
'End Sub
'
'
'Private Sub txtEqpCd_Change()
'    Call FormClear
'End Sub
'
'Private Sub txtEqpCd_GotFocus()
'   '
'    FocusMe Me.txtEqpCd
'   '
'End Sub
'
'Private Sub txtEqpCd_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If lstEQCode.ListCount = 0 Then Exit Sub
'    If KeyCode = vbKeyDown Then
'        lstEQCode.Visible = True
'        Set objCodeList = Nothing
'        lstEQCode.ListIndex = 0
'        lstEQCode.ZOrder 0
'        lstEQCode.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtEqpCd_KeyPress(KeyAscii As Integer)
'
'    Dim Char As String
'
'    Char = Chr(KeyAscii)
'    KeyAscii = Asc(UCase(Char))
'    If KeyAscii = vbKeyEscape Then Exit Sub
'    If KeyAscii = vbKeyReturn Then
'         Call lstEQCode_KeyDown(vbKeyReturn, 0)
'         lstEQCode.Visible = False
'         Exit Sub
'    End If
'
'    lstEQCode.Visible = True
'    Set objCodeList = Nothing
'    lstEQCode.ZOrder 0
'    Call medCodeHelp(KeyAscii, lstEQCode, txtEqpCd.Text, txtEqpCd, txtTransDt)
'
'End Sub
'
'Private Sub txtEqpCd_Validate(Cancel As Boolean)
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'    If ActiveControl.Name = cmdClear.Name Then Exit Sub
'    If ActiveControl.Name = cmdExit.Name Then Exit Sub
'
'
'    lblEqpCdNm.Caption = ""
'    If Trim(txtEqpCd.Text) = "" Then Exit Sub
'
'    Dim strEqpNm As String
'
'    strEqpNm = objRst.GetEqpNm(txtEqpCd.Text)
'    If Trim(strEqpNm) = "" Then
'        MsgBox "�ڵ� �Է� Error!", vbCritical
'        FocusMe Me.txtEqpCd
'        Exit Sub
'    End If
'    lblEqpCdNm.Caption = strEqpNm
'
'End Sub
'
'Private Sub lstEQCode_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyReturn Then
'        txtEqpCd.Text = medGetP(lstEQCode.Text, 1, vbTab)
'        lblEqpCdNm.Caption = medGetP(lstEQCode.Text, 2, vbTab)
'        lstEQCode.Visible = False
'    End If
'End Sub
'Private Sub lstEQCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then Call lstEQCode_KeyDown(vbKeyReturn, 0)
'End Sub
'
'Private Sub lstEQCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lstEQCode.SetFocus
'End Sub
'
'
'Private Sub GetDisplay(ByVal Col As Long, ByVal Row As Long)
'    Dim sTestcd As String
'    Dim sLabNo  As String
'
'
'    lblPtId.Caption = ""
'    lblSex.Caption = ""
'    lblWard.Caption = ""
'    lblDept.Caption = ""
'    lblRcvDT.Caption = ""
'    lblColDt.Caption = ""
'    lblOrdDt.Caption = ""
'    lblDoct.Caption = ""
'    lblPtNm.Caption = ""
'
'    With tblResult
'        .Row = Row
'        .Col = Col: sTestcd = medGetP(.Value, 1, COL_DIV)
'        .Col = TblCol.tcLABNO: sLabNo = .Value
'        If sLabNo = "" Then Exit Sub
'        objResult.KeyChange sLabNo & COL_DIV & sTestcd
'        lblPtId.Caption = objResult.Fields("ptid")
'        lblSex.Caption = objResult.Fields("sexage")
'        lblWard.Caption = objResult.Fields("wardid")
'        lblDept.Caption = objResult.Fields("deptcd")
'        lblRcvDT.Caption = objResult.Fields("rcvdate")
'        lblColDt.Caption = objResult.Fields("coldate")
'        lblOrdDt.Caption = objResult.Fields("orddt")
'        If ObjLISComCode.doct.Exists(objResult.Fields("orddoct")) Then
'            ObjLISComCode.doct.KeyChange objResult.Fields("orddoct")
'            lblDoct.Caption = ObjLISComCode.doct.Fields("doctnm")
'        Else
'            lblDoct.Caption = objResult.Fields("orddoct")
'        End If
'        lblPtNm.Caption = objResult.Fields("ptnm")
'
'    End With
'
'End Sub
'
'Private Sub tblResult_GotFocus()
'    With tblResult
'        If .MaxRows = 0 Then Exit Sub
'        .EditEnterAction = EditEnterActionDown
'    End With
'End Sub
'Private Sub tblResult_LostFocus()
'    Dim strTmp  As String
'    Dim sTestcd As String
'    Dim sLabNo  As String
'    Dim sValue  As String
'
'    If Screen.ActiveControl.Name <> tblResult.Name Then Exit Sub
'
'    With tblResult
'        If .ActiveRow < 1 Then Exit Sub
'        .Row = .ActiveRow
'        .Col = .ActiveCol:     strTmp = .Value                              '����Է°�:��������
'        If strTmp = "" Then Exit Sub
'        .Col = TblCol.tcLABNO: sLabNo = .Value
'        .Col = .ActiveCol - 1: sTestcd = medGetP(.Value, 1, COL_DIV)        '�˻��ڵ�
'        sValue = objRst.GetRstCd(sTestcd, UCase(strTmp))                    '�ڵ庰 �����
'        .Col = .ActiveCol:     .Value = sValue                              '���̴� �����
'        If objResult.Exists(sLabNo & COL_DIV & sTestcd) Then
'            objResult.KeyChange sLabNo & COL_DIV & sTestcd
'            objResult.Fields("rstcd") = strTmp
'        End If
'    End With
'End Sub
'Private Sub tblResult_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'    Dim sValue      As String
'    Dim sTestcd     As String
'    Dim sLabNo      As String
'    Dim ii          As Long
'
'    '���� ������Ʈ
'    objResult.MoveFirst
'    With tblResult
'        .Row = Row: .Col = TblCol.tcHold: sValue = .Value
'        .Col = TblCol.tcLABNO:    sLabNo = .Value
'        Do Until objResult.EOF
'
'            If objResult.Fields("labno") = sLabNo Then
'                objResult.Fields("hold") = sValue
'            End If
'            objResult.MoveNext
'        Loop
'    End With
'
'End Sub
'
'
'Private Sub tblResult_Advance(ByVal AdvanceNext As Boolean)
'    Dim strRstType  As String
'    Dim strErr      As String
'    Dim sLabNo      As String
'    Dim sTestcd     As String
'
'    Dim Col As Long
'    Dim Row As Long
'
'    Row = tblResult.ActiveRow
'    Col = tblResult.ActiveCol
'
'    If Row < 0 Then Exit Sub
'    If Col < 7 Then Exit Sub
'    If Col Mod 3 <> 0 Then Exit Sub
'
'    On Error GoTo ErrLevaeCell:
'   '
'    Col = tblResult.ActiveCol
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = Col - 1
'
'        If .Value = "" Then Exit Sub
'        .Col = TblCol.tcLABNO: sLabNo = .Value
'        .Col = Col - 1:        sTestcd = medGetP(.Value, 1, COL_DIV)
'        objResult.KeyChange sLabNo & COL_DIV & sTestcd
'
'        .Col = Col
'
'        If tblResult.Value = objResult.Fields("rstcdval") Then Exit Sub
'        objResult.Fields("rstcd") = tblResult.Value
'
'        strRstType = objResult.Fields("rsttype")
'
'        If strRstType = "N" Then
'            strErr = objResult.Fields("avalval")
'            If objRst.IsAvalVal(Col) = False Then
'                If strErr <> "0" Then
'                    strErr = "��ȿ���� �Է� ����. (" & objResult.Fields("avalval") & "�ڸ�)"
'                Else
'                    strErr = "��ȿ���� �Է� ����. (�������� �Է�)"
'                End If
'                GoTo ErrLevaeCell
'            Else
'                lblErr.Caption = ""
'                objRst.NumValCheck (Col)
'            End If
'        ElseIf strRstType = "A" Then
'            If objRst.IsAlphaCd(Col) = False Then
'                strErr = "��� �Է� ����!"
'                GoTo ErrLevaeCell
'            Else
'                lblErr.Caption = ""
'            End If
'        ElseIf strRstType = "R" Then
'            If objRst.IsRateCd(Col) = False Then
'                strErr = "������� �Է� ����!"
'                GoTo ErrLevaeCell
'            Else
'                lblErr.Caption = ""
'            End If
'        ElseIf strRstType = "F" Then
'            If objRst.IsFreeResult(Col) = False Then
'                strErr = "FREE��� �Է� ����! (10�ڸ��̳�)"
'                GoTo ErrLevaeCell
'            Else
'                objRst.NumValCheck (Col)
'                lblErr.Caption = ""
'            End If
'        End If
'
'    End With
'
'    objRst.ResultCheck (Col)
'    strRstType = objResult.Fields("rsttype")
'
'    If strRstType = "N" Then
'        strErr = objResult.Fields("avalval") 'objPtInfo.Result.Item(Row).AvalVal
'        If objRst.IsAvalVal(Col) = False Then
'            If strErr <> "0" Then
'                strErr = "��ȿ���� �Է� ����. (" & objResult.Fields("avalval") & "�ڸ�)"
'            Else
'                strErr = "��ȿ���� �Է� ����. (�������� �Է�)"
'            End If
'            GoTo ErrLevaeCell
'        Else
'            lblErr.Caption = ""
'            objRst.NumValCheck (Col)
'        End If
'    ElseIf strRstType = "A" Then
'        If objRst.IsAlphaCd(Col) = False Then
'            strErr = "��� �Է� ����!"
'            GoTo ErrLevaeCell
'        Else
'            lblErr.Caption = ""
'        End If
'    ElseIf strRstType = "R" Then
'        If objRst.IsRateCd(Col) = False Then
'            strErr = "������� �Է� ����!"
'            GoTo ErrLevaeCell
'        Else
'            lblErr.Caption = ""
'        End If
'    ElseIf strRstType = "F" Then
'        If objRst.IsFreeResult(Col) = False Then
'            strErr = "FREE��� �Է� ����! (10�ڸ��̳�)"
'            GoTo ErrLevaeCell
'        Else
'            objRst.NumValCheck (Col)
'            lblErr.Caption = ""
'        End If
'    End If
'
' '���� �Է��Ѱ��� ������ �ڵ尪�̸� �ڵ� ��������� ��ȯ�Ѵ�.
'
'    Dim strCodeValue As String
'
'    tblResult.Row = Row
'    tblResult.Col = Col
'    strCodeValue = UCase(Trim(tblResult.Value))
'
'    If strCodeValue <> "" Then
'        If objRst.GetRstCd(sTestcd, strCodeValue) <> tblResult.Value Then
'            tblResult.Value = objRst.GetRstCd(sTestcd, strCodeValue)
'        End If
'    End If
'
'    objResult.Fields("rstcdval") = tblResult.Value
'
'    Exit Sub
'
'ErrLevaeCell:
'    lblErr.Caption = strErr
'    tblResult.Value = ""
'    MsgBox strErr, vbCritical + vbOKOnly, "����Է� Ȯ��"
'    DoEvents
'   '
'    With tblResult
'        .Row = Row
'        .Col = Col
'        .Value = ""
'        .Action = ActionActiveCell
'        .SetFocus
'    End With
'    objRst.ResultCheck (Col)
'
'End Sub
'
'Private Sub tblResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'
'    Dim strRstType  As String
'    Dim strErr      As String
'    Dim sLabNo      As String
'    Dim sTestcd     As String
'
'
'
'    If Row < 1 Then Exit Sub
'    If Col < 7 Then Exit Sub
'    If Col Mod 3 <> 0 Then Exit Sub
'
'    With tblResult
'        .Row = Row: .Col = Col - 1
'        If .Value = "" Then Exit Sub
'    End With
'
'On Error GoTo ErrLevaeCell
'
'    If Row = tblResult.MaxRows Then
'        Call tblResult_LostFocus
'        Exit Sub
'    End If
'
'    lblErr.Caption = ""
'
'    tblResult.Row = Row
'    tblResult.Col = TblCol.tcLABNO: sLabNo = tblResult.Value
'    tblResult.Col = Col - 1:        sTestcd = medGetP(tblResult.Value, 1, COL_DIV)
'
'
'
'    objResult.KeyChange sLabNo & COL_DIV & sTestcd
'
'    tblResult.Col = Col
'
'    If tblResult.Value = objResult.Fields("rstcdval") Then
'        Call GetDisplay(8, NewRow)
'        Exit Sub
'    End If
'
'    objResult.Fields("rstcd") = tblResult.Value
'
'    objRst.ResultCheck (Col)
'    strRstType = objResult.Fields("rsttype")
'    If strRstType = "N" Then
'        strErr = objResult.Fields("avalval")
'        If objRst.IsAvalVal(Col) = False Then
'            If strErr <> "0" Then
'                strErr = "��ȿ���� �Է� ����. (" & objResult.Fields("avalval") & "�ڸ�)"
'            Else
'                strErr = "��ȿ���� �Է� ����. (�������� �Է�)"
'            End If
'            GoTo ErrLevaeCell
'        Else
'            objRst.NumValCheck (Col)
'        End If
'    ElseIf strRstType = "A" Then
'        If objRst.IsAlphaCd(Col) = False Then
'           strErr = "��� �Է� ����!"
'           GoTo ErrLevaeCell
'        End If
'    ElseIf strRstType = "R" Then
'        If objRst.IsRateCd(Col) = False Then
'           strErr = "������� �Է� ����!"
'           GoTo ErrLevaeCell
'        End If
'    ElseIf strRstType = "F" Then
'        If objRst.IsFreeResult(Col) = False Then
'           strErr = "FREE��� �Է� ����! (10�ڸ��̳�)"
'           GoTo ErrLevaeCell
'        End If
'        objRst.NumValCheck (Col)
'    End If
'    tblResult.EditEnterAction = EditEnterActionDown
'
''���� �Է��Ѱ��� ������ �ڵ尪�̸� �ڵ� ��������� ��ȯ�Ѵ�.
'    Dim strCodeValue As String
'
'    tblResult.Row = Row
'    tblResult.Col = Col
'    strCodeValue = UCase(Trim(tblResult.Value))
'
'    If strCodeValue <> "" Then
'        If objRst.GetRstCd(sTestcd, objResult.Fields("rstcd")) <> tblResult.Value Then
'            tblResult.Value = objRst.GetRstCd(sTestcd, strCodeValue)
'        End If
'    End If
'
'    objResult.Fields("rstcdval") = tblResult.Value
'
'
'
'
'    With tblResult
'        .Col = Col + 2
'        If .Value <> "" Then
'            .Col = Col + 3
'        Else
'            .Row = Row + 1
'
'            .Col = TblCol.tcNO
'            If .Value = "" Then Exit Sub
'
'            .Col = TblCol.tcHold + 3
'        End If
'
'        Call GetDisplay(8, .Row)
'        .Col = Col + 2
'        If .Value <> "" Then
'            .Col = Col + 3
'        Else
'            .Col = TblCol.tcHold + 3
'        End If
'
'        .Action = ActionActiveCell
'    End With
'
'
'    Exit Sub
'
'ErrLevaeCell:
'    lblErr.Caption = strErr
'    tblResult.EditEnterAction = EditEnterActionDown
'    objResult.Fields("rstcd") = ""
'    DoEvents
'    With tblResult
'        .Row = Row
'        .Col = Col
'        .Value = ""
'        .Action = ActionActiveCell
'    End With
'    objRst.ResultCheck (Col)
'   '
'    MsgBox strErr, vbCritical + vbOKOnly, "����Է� Ȯ��"
'    Cancel = True
'    tblResult.SetFocus
'
'End Sub
'Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'    If Row < 1 Then Exit Sub            '�������
'    If Col = 2 Then
'        tblResult.Row = Row
'        tblResult.Col = Col
'        tblResult.Action = ActionActiveCell
'        If tblResult.Value = "" Then Exit Sub
'        Set mnuPopup = frmControls.mnuPopup
'        Set mnuResult = frmControls.mnuSub
'        Set mnuResult1 = frmControls.mnuSub1
'        Set mnuRemark = frmControls.mnuSub2
'
'        mnuResult.Caption = "������ȣ�� �������"
'        mnuResult1.Caption = "�˻��׸� �������"
'        mnuRemark.Caption = "��� Remark ���"
'
'        mnuResult.Visible = True
'        mnuRemark.Visible = True
'        mnuResult1.Visible = False
'
'        PopupMenu mnuPopup
'        Set mnuPopup = Nothing
'        Set mnuResult = Nothing
'        Set mnuResult1 = Nothing
'        Set mnuRemark = Nothing
'
'    Else
'        If Col > 6 Then
'
'            If Col Mod 3 = 1 Then
'                tblResult.Col = Col + 1
'                If tblResult.Value = "" Then Exit Sub
'                tblResult.Row = Row
'                tblResult.Col = Col
'                tblResult.Action = ActionActiveCell
'
'                Set mnuPopup = frmControls.mnuPopup
'                Set mnuResult = frmControls.mnuSub
'                Set mnuResult1 = frmControls.mnuSub1
'                Set mnuRemark = frmControls.mnuSub2
'
'                mnuResult1.Caption = "�˻��׸� �������"
'
'                mnuResult1.Visible = True
'                mnuResult.Visible = False
'                mnuRemark.Visible = False
'                PopupMenu mnuPopup
'                Set mnuPopup = Nothing
'                Set mnuResult = Nothing
'                Set mnuResult1 = Nothing
'                Set mnuRemark = Nothing
'
'            Else
'                If Col < 9 Then Exit Sub            '�������� ����
'
'                If Col Mod 3 <> 0 Then Exit Sub     '�������� ���Ҷ�.
'
'                If Row <= 0 Then Exit Sub
'
'                objRst.SsTop = picRst.Top + 220
'                objRst.SsLeft = picRst.Left - 740
'                tblResult.Row = Row
'                tblResult.Col = Col
'                tblResult.Action = ActionActiveCell
'                Call objRst.PopUp(, Col)
'            End If
'
'        End If
'    End If
'
'End Sub
'
'Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'    If Row < 1 Then Exit Sub
'    If Col < 7 Then Exit Sub
'    If Col Mod 3 <> 1 Then Exit Sub
'
'    Dim strtip  As String
'    Dim sTestcd As String
'    Dim sLabNo  As String
'
'    With tblResult
'        Call .SetTextTipAppearance("����ü", 9, False, False, &HFFFFC0, vbBlack)
'        .Row = Row
'        .Col = Col + 1
'        sTestcd = medGetP(.Value, 1, COL_DIV)
'        .Col = TblCol.tcLABNO
'        sLabNo = .Value
'        If sTestcd = "" Then Exit Sub
'        objResult.KeyChange sLabNo & COL_DIV & sTestcd
'        .Col = Col
'
'        strtip = " �� �� ��    : " & .Value & vbNewLine
'        .Col = TblCol.tcSpc
'        strtip = strtip & " ��    ü    : " & .Value & vbNewLine
'        If objResult.Fields("refvalfrom") <> "" Then
'            strtip = strtip & " �� �� ġ    : " & objResult.Fields("refvalfrom") & " - " & objResult.Fields("refvalto") & vbNewLine
'        End If
'
'        If objResult.Fields("afrval") <> "" Then
'            strtip = strtip & " Auto ����ġ : " & objResult.Fields("afrval") & " - " & objResult.Fields("atoval") & vbNewLine
'        End If
'
'        If objResult.Fields("lastrst") <> "" Then
'            strtip = strtip & " �ֱٰ��    : " & vbNewLine & _
'                     "          [ " & objResult.Fields("lastrst") & " ]       �ֱٺ����Ͻ� : " & Format(objResult.Fields("lastvfydt"), "0###-##-##")
'        End If
'        TipText = vbNewLine & strtip & vbNewLine
'        TipWidth = 6000
'        MultiLine = 1
'
'        ShowTip = True
'
'    End With
'
'End Sub
'Private Sub cmdSave_Click()
'    If objResult Is Nothing Then Exit Sub
'    If objResult.RecordCount < 1 Then Exit Sub
'
'    Me.MousePointer = 11
'
'    Call ResultSAVE
'
'    If chkAuto.Value = 1 Then
'        Call cmdQuery_Click
'    Else
'        Call FormClear
'    End If
'
'    Me.MousePointer = 0
'End Sub
'
'
'Private Function ResultSAVE() As Boolean
'
'    Dim Rs          As DrRecordSet
'    Dim sWorkArea   As String
'    Dim sAccDt      As String
'    Dim sAccSeq     As String
'    Dim sTestcd     As String
'    Dim sDeptCD     As String
'    Dim sBussdiv    As String
'    Dim tmpSQL      As String
'    Dim sVfyDt      As String
'    Dim sVfyTm      As String '
'    Dim svfyID      As String
'    Dim SqlStmt     As String
'    Dim sRstCD      As String
'
'    Dim tmpInputCnt As Integer
'    Dim tmpTotCnt   As Integer
'    Dim lngCnt      As Long
'
'On Error GoTo DBExecError
'
'    Dim objPrgBar As New clsProgress
'
'    DoEvents
'    With objPrgBar
'        .Mode = 0
'        .CaptionOn = False
'        .Min = 0
'        .Max = objResult.RecordCount * 3
'        .Value = 0
'        .Visible = True
'    End With
'
'
'    sVfyDt = Format(DBConn.getsysdate, "YYYYMMDD")
'    sVfyTm = Format(DBConn.getsysdate, "HHMMSS")
'    svfyID = ObjSysInfo.EmpId
'
'    DBConn.BeginTrans
'
'    objResult.MoveFirst
'
'    With objResult
'        Do Until .EOF
'            sWorkArea = "": sAccDt = "": sAccSeq = "": sTestcd = ""
'            sWorkArea = medGetP(.Fields("labno"), 1, "-")
'            sAccDt = medGetP(.Fields("labno"), 2, "-")
'            sAccSeq = medGetP(.Fields("labno"), 3, "-")
'            sTestcd = .Fields("testcd")
'            sRstCD = Trim(.Fields("rstcd"))
'
'            '����üũ�� �ȵǾ������ �����Ѵ�.
'            If .Fields("hold") <> "1" Then
'                '-----------------------------------------------------------------
'                'AutoVerify�ΰ��
'                '����ġ�� Arefvalfrom,arefToval �� ������ HIGH/LOW�� �ٽ� üũ�Ѵ�.
'                '���࿡ �ɸ����� �ƹ����� ���Ѵ�.
'                '-----------------------------------------------------------------
'                If chkAuto.Value = 1 Then
'                    '---------------------------------------------------------
'                    '���ο� ����ġ�� ���ؼ� �����Ѵ�.'/* HIGH,LOW CHECK */
'                    '---------------------------------------------------------
'
'                    If Val(.Fields("afrval")) <> 0 Or Val(.Fields("atoval")) <> 0 Then
'                        If Val(sRstCD) < Val(.Fields("afrval")) Then
'                            'LOW �� �ɸ���.
'                            GoTo AutoVerifyChk_Ref1
'                        End If
'                        If Val(sRstCD) > Val(.Fields("atoval")) Then
'                            'HIGH �� �ɸ���.
'                            GoTo AutoVerifyChk_Ref1
'                        End If
'                    End If
'
'                    'HIGH/LOW �� üũ �߱⶧���� DELTA/PANIC �� üũ�Ѵ�.
'                    'Auto�� ����ġ�� ���°��� �Ϲ� ����ġ���Ѱ� ������ AutoVerify�� üũ�Ѵ�.
'
'                    If Val(.Fields("afrval")) = 0 And Val(.Fields("atoval")) = 0 Then
'                        If .Fields("hldiv") <> "" Then GoTo AutoVerifyChk
'                    End If
'                    If .Fields("dpdiv") <> "" Then GoTo AutoVerifyChk
'                End If
'
'                '����� ����ִ°��
'                If Trim(.Fields("rstcd")) <> "" And .Fields("rstcd") <> "ERR" Then
'                    's2lab302(�����������)
'                    SqlStmt = DBW("rstval", .Fields("rstval"), 3) & _
'                                      DBW("rstcd", sRstCD, 3) & _
'                                      DBW("hldiv", .Fields("hldiv"), 3) & _
'                                      DBW("dpdiv", .Fields("dpdiv"), 3) & _
'                                      DBW("eqpcd", .Fields("eqpcd"), 3) & _
'                                      DBW("txtfg", .Fields("txtfg"), 3) & _
'                                      DBW("rsttype", .Fields("rsttype"), 3) & _
'                                      DBW("vfydt", sVfyDt, 3) & _
'                                      DBW("vfytm", sVfyTm, 3) & _
'                                      DBW("vfyid", svfyID, 2)
'                    SqlStmt = " UPDATE " & T_LAB302 & " SET " & SqlStmt & _
'                                  " WHERE  " & DBW("workarea", sWorkArea, 2) & _
'                                  " AND    " & DBW("accdt", sAccDt, 2) & _
'                                  " AND    " & DBW("accseq", sAccSeq, 2) & _
'                                  " AND    " & DBW("testcd", sTestcd, 2) & _
'                                  " AND   (vfydt = ''  or  vfydt is null) "  '��������� ���Ȯ�� �� �������� �����Ѵ�(�������� ���� ������ ����)... 99.12.10 by KMK
'                    '��������
'                    .Fields("vfydt") = sVfyDt
'
'
'                    DBConn.Execute SqlStmt
'                    '�ؽ�Ʈ������
'                    If .Fields("txtfg") = "1" Then
'                        If objRst.UpdatableLAB303(sWorkArea, sAccDt, sAccSeq, sTestcd) = False Then
'                            SqlStmt = " INSERT INTO " & T_LAB303 & " (workarea,accdt,accseq,testcd, rsttxt)" & _
'                                              " VALUES(" & _
'                                              DBV("workarea", sWorkArea, 1) & _
'                                              DBV("accdt", sAccDt, 1) & _
'                                              DBV("accseq", sAccSeq, 1) & _
'                                              DBV("testcd", sTestcd, 1) & _
'                                              DBV("rsttxt", .Fields("rsttext")) & ")"
'                        Else
'                            SqlStmt = " UPDATE " & T_LAB303 & _
'                                              " SET    " & _
'                                                           DBW("rsttxt", .Fields("rsttext"), 2) & _
'                                              " WHERE  " & DBW("workarea", sWorkArea, 2) & _
'                                              " AND    " & DBW("accdt", sAccDt, 2) & _
'                                              " AND    " & DBW("accseq", sAccSeq, 2) & _
'                                              " AND    " & DBW("testcd", sTestcd, 2)
'                        End If
'
'                        DBConn.Execute SqlStmt
'                    End If
'
'                    '�ܺ��Ƿڳ��� Status �ݿ�
'                    If objRst.UpdatableLAB205(sWorkArea, sAccDt, sAccSeq, sTestcd) Then
'                        SqlStmt = " UPDATE " & T_LAB205 & _
'                                  " SET    " & _
'                                             DBW("stscd", enStsCd.StsCd_LIS_MidRst, 2) & _
'                                  " WHERE  " & DBW("workarea", sWorkArea, 2) & _
'                                  " AND    " & DBW("accdt", sAccDt, 2) & _
'                                  " AND    " & DBW("accseq", sAccSeq, 2) & _
'                                  " AND    " & DBW("testcd", sTestcd, 2)
'
'                        DBConn.Execute SqlStmt
'                    End If
'                    '������� ��⳻��
'                    If medGetP(.Fields("wardid"), 1, "-") = "" Then
'                        sDeptCD = .Fields("deptcd")
'                        sBussdiv = enBussDiv.BussDiv_OutPatient
'                    Else
'                        sDeptCD = .Fields("deptcd")
'                        sBussdiv = enBussDiv.BussDiv_InPatient
'                    End If
'
'                    tmpSQL = " SELECT * FROM " & T_LAB202 & _
'                             " WHERE  " & DBW("deptcd", sDeptCD, 2) & _
'                             " AND    " & DBW("vfydt", sVfyDt, 2) & _
'                             " AND    " & DBW("ptid", .Fields("ptid"), 2) & _
'                             " AND    " & DBW("mfyfg", "0", 2)
'
'                    Set Rs = OpenRecordSet(tmpSQL)
'                    If Rs.EOF Then
'                        SqlStmt = "INSERT INTO " & T_LAB202 & " (deptcd, vfydt, ptid, mfyfg, vfytm, vfyid, donefg, doneid, majdoct, bussdiv) " & _
'                                    " VALUES ( " & _
'                                    DBV("deptcd", sDeptCD, 1) & _
'                                    DBV("vfydt", sVfyDt, 1) & _
'                                    DBV("ptid", .Fields("ptid"), 1) & _
'                                    DBV("mfyfg", "0", 1) & _
'                                    DBV("vfytm", sVfyTm, 1) & _
'                                    DBV("vfyid", svfyID, 1) & _
'                                    DBV("donefg", "", 1) & _
'                                    DBV("doneid", "0", 1) & _
'                                    DBV("majdoct", .Fields("orddoct"), 1) & _
'                                    DBV("bussdiv", sBussdiv) & " ) "
'                    Else
'                        If "" & Rs.Fields("DoneFg").Value = "1" Then
'                            SqlStmt = " UPDATE " & T_LAB202 & _
'                                              " SET    " & _
'                                                           DBW("vfytm", sVfyTm, 3) & _
'                                                           DBW("vfyid", svfyID, 3) & _
'                                                           DBW("majdoct", .Fields("orddoct"), 3) & _
'                                                           DBW("donefg", "", 3) & _
'                                                           DBW("doneid", "0", 2) & _
'                                              " WHERE  " & DBW("deptcd", sDeptCD, 2) & _
'                                              " AND    " & DBW("vfydt", sVfyDt, 2) & _
'                                              " AND    " & DBW("ptid", .Fields("ptid"), 2) & _
'                                              " AND    " & DBW("mfyfg", "0", 2)
'                        End If
'                    End If
'
'                    DBConn.Execute SqlStmt
'                    Set Rs = Nothing
'
'                End If
'            Else
'AutoVerifyChk:
'                SqlStmt = DBW("rstval", .Fields("rstval"), 3) & _
'                                  DBW("rstcd", sRstCD, 3) & _
'                                  DBW("hldiv", .Fields("hldiv"), 3) & _
'                                  DBW("dpdiv", .Fields("dpdiv"), 3) & _
'                                  DBW("eqpcd", .Fields("eqpcd"), 3) & _
'                                  DBW("txtfg", .Fields("txtfg"), 2)
'                SqlStmt = " UPDATE " & T_LAB302 & " SET " & SqlStmt & _
'                              " WHERE  " & DBW("workarea", sWorkArea, 2) & _
'                              " AND    " & DBW("accdt", sAccDt, 2) & _
'                              " AND    " & DBW("accseq", sAccSeq, 2) & _
'                              " AND    " & DBW("testcd", sTestcd, 2) & _
'                              " AND   (vfydt = ''  or  vfydt is null) "  '��������� ���Ȯ�� �� �������� �����Ѵ�(�������� ���� ������ ����)... 99.12.10 by KMK
'
'                DBConn.Execute SqlStmt
'
'            End If
'
'AutoVerifyChk_Ref1:
'
'
'            DoEvents
'            lngCnt = lngCnt + 1
'            objPrgBar.MSG = objResult.Fields("ptnm") & " �� " & objResult.Fields("abbrnm5") & " �׸��� ����� �������Դϴ�."
'            objPrgBar.Value = lngCnt
'
'            .MoveNext
'        Loop
'
'    End With
'
'
'    '�����ڵ带 ������Ʈ ����� �Ѵ�.
'    'Rstdiv='*' �̰� Detailfg<>'' �ΰ� �����ڵ�
'    '�����ǿ��� Rstdiv�� 'R' �ΰſ� ���ؼ� ��� vfydt�� <>'' �ƴϸ� �����ڵ带 ������Ʈ ���ش�.
'
'    Dim sMotherCode As String
'    Dim sLabNo      As String
'    Dim intUpCnt    As Integer
'    Dim realUpCnt   As Integer
'
'
'    With objResult
'        .MoveFirst
'        Do Until .EOF
'            If .Fields("hold") <> "1" Then
'                'AutoVerify�ΰ��(hldiv="" AND dpdiv="") �ΰ�츸 �����Ѵ�.
'                If chkAuto.Value = 1 Then
'                    If Val(.Fields("afrval")) <> 0 Or Val(.Fields("atoval")) <> 0 Then
'                        If Val(.Fields("afrval")) < Val(.Fields("afrval")) Then
'                            'LOW �� �ɸ���.
'                            GoTo AutoVerifyChk1
'                        End If
'                        If Val(.Fields("afrval")) > Val(.Fields("atoval")) Then
'                            'HIGH �� �ɸ���.
'                            GoTo AutoVerifyChk1
'                        End If
'                    End If
'                    'HIGH/LOW �� üũ �߱⶧���� DELTA/PANIC �� üũ�Ѵ�.
'                    'Auto�� ����ġ�� ���°��� �Ϲ� ����ġ���Ѱ� ������ AutoVerify�� üũ�Ѵ�.
'                    If Val(.Fields("afrval")) = 0 And Val(.Fields("atoval")) = 0 Then
'                        If .Fields("hldiv") <> "" Then GoTo AutoVerifyChk1
'                    End If
'                    If .Fields("dpdiv") <> "" Then GoTo AutoVerifyChk1
'
'                    If sLabNo = .Fields("labno") Then
'                       If .Fields("rstdiv") = "R" Then
'                           intUpCnt = intUpCnt + 1
'                       End If
'                    Else
'                        GoTo AutoVerifyChk1
'                    End If
'
'                End If
'                If .Fields("rstdiv") = "*" And .Fields("detailfg") <> "" Then
'
'                    If sLabNo <> "" And (intUpCnt = realUpCnt) And intUpCnt <> 0 Then
'                        '�����ڵ� ������Ʈ
'                        SqlStmt = DBW("vfydt", sVfyDt, 3) & _
'                                  DBW("vfytm", sVfyTm, 3) & _
'                                  DBW("vfyid", svfyID, 2)
'                        SqlStmt = " UPDATE " & T_LAB302 & " SET " & SqlStmt & _
'                                  " WHERE  " & DBW("workarea", medGetP(sLabNo, 1, "-"), 2) & _
'                                  " AND    " & DBW("accdt", medGetP(sLabNo, 2, "-"), 2) & _
'                                  " AND    " & DBW("accseq", medGetP(sLabNo, 3, "-"), 2) & _
'                                  " AND    " & DBW("testcd", sMotherCode, 2) & _
'                                  " AND    (vfydt = ''  or vfydt is null) "   '��������� ���Ȯ�� �� �������� �����Ѵ�(�������� ���� ������ ����)... 99.12.10 by KMK
'
'                        DBConn.Execute SqlStmt
'                    End If
'
'                    intUpCnt = 0: realUpCnt = 0
'                    sLabNo = "": sMotherCode = ""
'                    sLabNo = .Fields("labno")
'                    sMotherCode = .Fields("testcd")
'                End If
'
'                If sLabNo = .Fields("labno") Then
'                    If .Fields("rstdiv") = "R" Then
'                        intUpCnt = intUpCnt + 1
'                        If .Fields("vfydt") <> "" Then
'                            realUpCnt = realUpCnt + 1
'                        End If
'                    End If
'                End If
'
'            End If
'
'
'AutoVerifyChk1:
'            DoEvents
'            lngCnt = lngCnt + 1
'            objPrgBar.MSG = objResult.Fields("ptnm") & " �� ������¸� ������Ʈ �մϴ�."
'            objPrgBar.Value = lngCnt
'
'            .MoveNext
'        Loop
'
'        If sLabNo <> "" And (intUpCnt = realUpCnt) And intUpCnt <> 0 And sMotherCode <> "" Then
'            '�����ڵ� ������Ʈ
'            SqlStmt = DBW("vfydt", sVfyDt, 3) & _
'                      DBW("vfytm", sVfyTm, 3) & _
'                      DBW("vfyid", svfyID, 2)
'            SqlStmt = " UPDATE " & T_LAB302 & " SET " & SqlStmt & _
'                      " WHERE  " & DBW("workarea", medGetP(sLabNo, 1, "-"), 2) & _
'                      " AND    " & DBW("accdt", medGetP(sLabNo, 2, "-"), 2) & _
'                      " AND    " & DBW("accseq", medGetP(sLabNo, 3, "-"), 2) & _
'                      " AND    " & DBW("testcd", sMotherCode, 2) & _
'                      " AND    (vfydt = ''  or vfydt is null) "   '��������� ���Ȯ�� �� �������� �����Ѵ�(�������� ���� ������ ����)... 99.12.10 by KMK
'
'            DBConn.Execute SqlStmt
'        End If
'
'    End With
'
'
'    Dim tmpStsCd As String
'    Dim tmpPtKey As String
'    Dim tmpAccDt As String
'
'    With objResult
'        .MoveFirst
'        Do Until .EOF
'            sWorkArea = "": sAccDt = "": sAccSeq = "": sTestcd = ""
'            sWorkArea = medGetP(.Fields("labno"), 1, "-")
'            sAccDt = medGetP(.Fields("labno"), 2, "-")
'            sAccSeq = medGetP(.Fields("labno"), 3, "-")
'            sTestcd = .Fields("testcd")
'
'            If .Fields("hold") <> "1" Then
'                'AutoVerify�ΰ��(hldiv="" AND dpdiv="") �ΰ�츸 �����Ѵ�.
'                If chkAuto.Value = 1 Then
'                    If Val(.Fields("afrval")) <> 0 Or Val(.Fields("atoval")) <> 0 Then
'                        If Val(.Fields("afrval")) < Val(.Fields("afrval")) Then
'                            'LOW �� �ɸ���.
'                            GoTo Skip
'                        End If
'                        If Val(.Fields("afrval")) > Val(.Fields("atoval")) Then
'                            'HIGH �� �ɸ���.
'                            GoTo Skip
'                        End If
'                    End If
'                    'HIGH/LOW �� üũ �߱⶧���� DELTA/PANIC �� üũ�Ѵ�.
'                    'Auto�� ����ġ�� ���°��� �Ϲ� ����ġ���Ѱ� ������ AutoVerify�� üũ�Ѵ�.
'                    If Val(.Fields("afrval")) = 0 And Val(.Fields("atoval")) = 0 Then
'                        If .Fields("hldiv") <> "" Then GoTo Skip
'                    End If
'                    If .Fields("dpdiv") <> "" Then GoTo Skip
'                End If
'                '------------------------------------------------------------
'                '�������� ������Ʈ
'                '------------------------------------------------------------
'                If tmpAccDt <> sWorkArea & sAccDt & sAccSeq Then
'                    tmpInputCnt = objRst.GetInputCnt(sWorkArea, sAccDt, sAccSeq)
'
'                    tmpTotCnt = objRst.GetTotCnt(sWorkArea, sAccDt, sAccSeq)
'
'                    SqlStmt = DBW("reqinputcnt", tmpInputCnt, 3) & _
'                              DBW("rmkcd", .Fields("rmkcd"), 3)
'
'                    If tmpTotCnt = tmpInputCnt Then
'                        '��ü Verify�� Update
'                        SqlStmt = SqlStmt & DBW("stscd", enStsCd.StsCd_LIS_FinRst, 3) & _
'                                          DBW("vfydt", sVfyDt, 3) & _
'                                          DBW("vfytm", sVfyTm, 3) & _
'                                          DBW("vfyid", svfyID, 3)
'                    End If
'
'                    If .Fields("footnote") <> "" Then
'                        SqlStmt = SqlStmt & DBW("footnotefg", "1", 2)
'                    Else
'                        SqlStmt = SqlStmt & DBW("footnotefg", "0", 2)
'                    End If
'
'                    SqlStmt = " UPDATE " & T_LAB201 & _
'                              " SET    " & SqlStmt & _
'                              " WHERE  " & DBW("workarea", sWorkArea, 2) & _
'                              " AND    " & DBW("accdt", sAccDt, 2) & _
'                              " AND    " & DBW("accseq", sAccSeq, 2)
'
'
'                    DBConn.Execute SqlStmt
'
'                    If .Fields("footnote") <> "" Then
'                        'FOOTNOTE ���
'                        SqlStmt = ""
'                        If objRst.UpdatableLAB304(sWorkArea, sAccDt, sAccSeq, "1") = False Then
'                            SqlStmt = " INSERT INTO " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt)" & _
'                                     " VALUES(" & _
'                                     DBV("workarea", sWorkArea, 1) & _
'                                     DBV("accdt", sAccDt, 1) & _
'                                     DBV("accseq", sAccSeq, 1) & _
'                                     DBV("seq", "1", 1) & _
'                                     DBV("vfyid", svfyID, 1) & _
'                                     DBV("rsttxt", .Fields("footnote")) & ")"
'                        Else
'                            SqlStmt = " UPDATE " & T_LAB304 & _
'                                     " SET    " & _
'                                                DBW("rsttxt", .Fields("footnote"), 3) & _
'                                                DBW("vfyid", svfyID, 2) & _
'                                     " WHERE  " & DBW("workarea", sWorkArea, 2) & _
'                                     " AND    " & DBW("accdt", sAccDt, 2) & _
'                                     " AND    " & DBW("accseq", sAccSeq, 2) & _
'                                     " AND    " & DBW("seq", "1", 2)
'                        End If
'                        DBConn.Execute SqlStmt
'                    End If
'
'                    tmpAccDt = sWorkArea & sAccDt & sAccSeq
'                End If
'
'                '------------------------------------------------------------
'                'ó��Body ���� ������Ʈ
'
'                If (.Fields("rstdiv") = "*" And Trim(.Fields("rstcd")) = "") Or Trim(.Fields("rstcd")) <> "" Then
'                    If tmpPtKey = .Fields("ptId") & Replace(.Fields("orddt"), "-", "") & .Fields("ordno") & .Fields("ordseq") & sTestcd Then GoTo Skip
'                    If .Fields("detailfg") <> "" Then
'                         '���׸��ΰ�� RstDiv�� "*"�� �� �������� ó��Body Update
'                        If .Fields("rstdiv") = "*" Then
'                            If objRst.IsOrderFollowUp(.Fields("ptId"), Replace(.Fields("orddt"), "-", ""), .Fields("ordno"), .Fields("ordseq")) = True Then
'                                tmpStsCd = enStsCd.StsCd_LIS_FinRst   '��ü Verify
'                            Else
'                                tmpStsCd = enStsCd.StsCd_LIS_MidRst  '�κ� Verify
'                            End If
'                             SqlStmt = " UPDATE " & T_LAB102 & _
'                                              " SET    " & _
'                                                           DBW("stscd", tmpStsCd, 3) & _
'                                                           DBW("examdt", sVfyDt, 3) & _
'                                                           DBW("examtm", sVfyTm, 3) & _
'                                                           DBW("examdoct", svfyID, 2) & _
'                                              " WHERE  " & DBW("ptid", .Fields("ptId"), 2) & _
'                                              " AND    " & DBW("orddt", Replace(.Fields("orddt"), "-", ""), 2) & _
'                                              " AND    " & DBW("ordno", .Fields("ordno"), 2) & _
'                                              " AND    " & DBW("ordseq", .Fields("ordseq"), 2)
'                            DBConn.Execute SqlStmt
'                        End If
'                    Else
'                        '�׷��ڵ� Ȥ�� ���� ITEM�ΰ��
'                        If objRst.IsOrderFollowUp(.Fields("ptId"), Replace(.Fields("orddt"), "-", ""), .Fields("ordno"), .Fields("ordseq")) = True Then
'                            tmpStsCd = enStsCd.StsCd_LIS_FinRst   '��ü Verify
'                        Else
'                            tmpStsCd = enStsCd.StsCd_LIS_MidRst  '�κ� Verify
'                        End If
'                        SqlStmt = " UPDATE " & T_LAB102 & _
'                                          " SET    " & _
'                                                       DBW("stscd", tmpStsCd, 3) & _
'                                                       DBW("examdt", sVfyDt, 3) & _
'                                                       DBW("examtm", sVfyTm, 3) & _
'                                                       DBW("examdoct", svfyID, 2) & _
'                                          " WHERE  " & DBW("ptid", .Fields("ptId"), 2) & _
'                                          " AND    " & DBW("orddt", Replace(.Fields("orddt"), "-", ""), 2) & _
'                                          " AND    " & DBW("ordno", .Fields("ordno"), 2) & _
'                                          " AND    " & DBW("ordseq", .Fields("ordseq"), 2)
'                        DBConn.Execute SqlStmt
'                    End If
'                    tmpPtKey = .Fields("ptId") & .Fields("orddt") & Replace(.Fields("orddt"), "-", "") & .Fields("ordseq") & sTestcd
'                End If
'            End If
'Skip:
'            DoEvents
'            lngCnt = lngCnt + 1
'            objPrgBar.MSG = objResult.Fields("ptnm") & " �� ó�泻���� ������Ʈ�մϴ�."
'            objPrgBar.Value = lngCnt
'
'            .MoveNext
'        Loop
'    End With
'
'    DBConn.CommitTrans
'    Set objPrgBar = Nothing
'
'    MsgBox "������ ó�� �Ǿ����ϴ�.", vbInformation + vbOKOnly, "Info"
'
''    If ICSResultChk = True Then
''         '��������
''         Dim objICS  As New clsICSResultChk
''         Dim strTmp  As String
''
''         strTmp = MsgBox("�������� ��� üũ�� �Ͻðڽ��ϱ�?", vbYesNo + vbInformation, "Info")
''         If strTmp = vbYes Then
''             Call objICS.ICSBatchResultCheck(objResult)
''         End If
''         Set objICS = Nothing
''     End If
'    Exit Function
'
'DBExecError:
'    DBConn.RollbackTrans
'    Set objPrgBar = Nothing
'    MsgBox "�������� ������ �߻��Ͽ����ϴ�.", vbInformation + vbOKOnly, "Info"
'
'End Function
''
'
'
