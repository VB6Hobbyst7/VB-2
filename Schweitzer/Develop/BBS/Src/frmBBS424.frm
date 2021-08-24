VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmBBS424 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈자 조회"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1665
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  헌혈자 등록 일자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  조회기간 설정"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   570
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  검색할 헌혈 종류"
      Appearance      =   0
   End
   Begin VB.Frame fraAcc1 
      BackColor       =   &H00DBE6E6&
      Height          =   870
      Left            =   8160
      TabIndex        =   7
      Top             =   795
      Width           =   4935
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   225
         Left            =   4020
         TabIndex        =   10
         Top             =   450
         Width           =   780
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "지정 헌혈"
         Height          =   435
         Index           =   0
         Left            =   600
         Style           =   1  '그래픽
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1245
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Pheresis"
         Height          =   435
         Index           =   1
         Left            =   2265
         Style           =   1  '그래픽
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   870
      Left            =   1440
      TabIndex        =   1
      Top             =   795
      Width           =   6705
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   330
         Left            =   1620
         TabIndex        =   3
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   64356355
         CurrentDate     =   36943
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   345
         Left            =   3675
         TabIndex        =   5
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   64356355
         CurrentDate     =   36943
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   360
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "등록일자"
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   3375
         TabIndex        =   4
         Top             =   420
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   495
      Left            =   11685
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "128"
      Top             =   8070
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   10365
      Style           =   1  '그래픽
      TabIndex        =   14
      Tag             =   "124"
      Top             =   8070
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   495
      Left            =   9045
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "15101"
      Top             =   8070
      Width           =   1215
   End
   Begin MSCommLib.MSComm MyComm 
      Left            =   13200
      Top             =   8070
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   5865
      Left            =   1440
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "10114"
      Top             =   1980
      Width           =   11655
      _Version        =   196608
      _ExtentX        =   20558
      _ExtentY        =   10345
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   14
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS424.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   20
   End
End
Attribute VB_Name = "frmBBS424"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RowHeight& = 12

'성명, 생년월일, 성별/나이, 혈액형, 등록일자, 헌혈종류, 지정환자, 혈액번호, 혈액제제, 헌혈량, 취소여부, 출력
Private Enum TblColumn
    tcName = 1
    tcDOB
    tcSEXAGE
    tcABO
    tcACCDT
    
    tcDONORTYPE
    tcSELPTID
    TcBLOODNO
    tcCOMP
    tcVOLUMN
    
    tcCANCEL
    tcCOMPOCD
    tcDONORID
    tcBAR
End Enum

Private Sub chkAll_Click()
    If chkAll.value = 1 Then
        optDonorCd(0).value = False
        optDonorCd(1).value = False
    Else
'        MsgBox "헌혈 종류를 선택하십시오.", vbExclamation
    End If
End Sub

Private Sub cmdClear_Click()
    Call InitForm
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS424 = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim objPro As clsProgress
    Dim RS As Recordset
    Dim rs1 As Recordset
    Dim strSQL As String
    Dim strDonorCd As String
    Dim strFDt As String
    Dim strTDt As String
    Dim strTmp As String
    Dim strPtid As String
    Dim strBldNo As String
    Dim i As Long
    
    Call InitTable
    
    If chkAll.value = 1 Then
        strDonorCd = "'1','3'"
    Else
        If optDonorCd(0).value Then
            strDonorCd = "'1'"
        ElseIf optDonorCd(1).value Then
            strDonorCd = "'3'"
        End If
    End If
    
    strFDt = Format(dtpFrDt.value, PRESENTDATE_FORMAT)
    strTDt = Format(dtpToDt.value, PRESENTDATE_FORMAT)
    
    strSQL = " SELECT a.donornm,a.dob,a.sex,a.abo,a.rh, b.donoraccdt,b.tmpid,b.donorcd,b.cancelfg,"
    strSQL = strSQL & " b.reservedid,b.donorid,b.bldsrc,b.bldyy,b.bldno,b.volumn,b.compocd, c.okdiv1,c.okdiv2,c.okdiv3,b.entfg,d.reserved,d.pherefg "
    strSQL = strSQL & " FROM " & T_BBS601 & " a, " & T_BBS602 & " b, " & T_BBS603 & " c, " & T_BBS401 & " d "
    strSQL = strSQL & " WHERE b.donoraccdt BETWEEN '" & strFDt & "' AND '" & strTDt & "' "
    strSQL = strSQL & " AND b.donorid=a.donorid  "
    strSQL = strSQL & " AND b.donorid=c.donorid "
    strSQL = strSQL & " AND b.donoraccdt=c.donoraccdt "
    strSQL = strSQL & " AND b.donorcd IN (" & strDonorCd & ")"
    strSQL = strSQL & " AND " & DBJ("b.bldsrc*=d.bldsrc")
    strSQL = strSQL & " AND " & DBJ("b.bldyy*=d.bldyy")
    strSQL = strSQL & " AND " & DBJ("b.bldno*=d.bldno")
    strSQL = strSQL & " AND " & DBJ("b.compocd*=d.compocd")
    strSQL = strSQL & " ORDER BY a.donornm,b.donoraccdt"
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        MsgBox "조건 기간 동안 등록된 헌혈자가 없습니다.", vbInformation
        Set RS = Nothing
        Exit Sub
    End If
    
    Set objPro = New clsProgress
    With objPro
        .Container = Me
        .Left = LisLabel3.Left
        .Top = LisLabel3.Top
        .Width = LisLabel3.Width
        .Height = LisLabel3.Height
        .Max = RS.RecordCount
    End With
    
    With tblList
        .ReDraw = False
        Do Until RS.EOF
            If .MaxRows < .DataRowCnt Then
                .MaxRows = .MaxRows + 1
            End If
            .Row = .DataRowCnt + 1
            
'성명, 생년월일, 성별/나이, 혈액형, 등록일자, 헌혈종류, 지정환자, 혈액번호, 혈액제제, 헌혈량, 취소여부, 출력
            
            If strTmp <> RS.Fields("donorid").value & "" Then
                .Col = TblColumn.tcName:   .value = RS.Fields("donornm").value & ""
                .Col = TblColumn.tcDOB:    .value = Format(RS.Fields("dob").value & "", "####-##-##")
                .Col = TblColumn.tcSEXAGE: .value = RS.Fields("sex").value & "" & "/"
                                           If Trim(RS.Fields("dob").value & "") <> "" Then
                                               .value = .value & medFindAge(RS.Fields("dob").value & "", "Y")
                                           End If
                .Col = TblColumn.tcABO: .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
            End If
            
            .Col = TblColumn.tcACCDT: .value = Format(RS.Fields("donoraccdt").value & "", "####/##/##")
            .Col = TblColumn.tcSELPTID
                If RS.Fields("reservedid").value & "" <> "" Then
                    strPtid = GetPtNm(RS.Fields("reservedid").value & "")
                    If strPtid <> "" Then
                        .value = strPtid & "(" & RS.Fields("reservedid").value & "" & ")"
                    End If
                Else
                    .value = ""
                End If
            .Col = TblColumn.TcBLOODNO
                strBldNo = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & Format(RS.Fields("bldno").value & "", "000000")
                If strBldNo = "--000000" Or strBldNo = "--" Then strBldNo = ""
                .value = strBldNo
                If strBldNo <> "" Then
                    .Col = 25: .CellType = CellTypeButton
                               .TypeButtonText = "BAR"
                Else
                    .Col = 25: .CellType = CellTypeStaticText
                               .Text = ""
                End If
            .Col = TblColumn.tcDONORTYPE
                Select Case RS.Fields("donorcd").value & ""
                    Case "1"
                        If RS.Fields("reserved").value & "" = "1" Then
                            .value = "지정헌혈"
                        ElseIf RS.Fields("reserved").value & "" = "0" Then
                            .value = "지정취소"
                            .Col = TblColumn.tcBAR: .CellType = CellTypeStaticText
                                                    .Text = ""
                        Else
                            .value = "지정헌혈"
                        End If
                    Case "3"
                        If RS.Fields("pherefg").value & "" = "1" Then
                            .value = "Pheresis"
                        ElseIf RS.Fields("pherefg").value & "" = "0" Then
                            .value = "Pheresis취소"
                            .Col = TblColumn.tcBAR: .CellType = CellTypeStaticText
                                                    .Text = ""
                        Else
                            .value = "Pheresis"
                        End If
                End Select
            .Col = TblColumn.tcCOMP: .value = medGetP(Get_CompNm(RS.Fields("compocd").value & ""), 1, COL_DIV)
            .Col = TblColumn.tcCOMPOCD: .value = RS.Fields("compocd").value & ""
            .Col = TblColumn.tcVOLUMN: .value = IIf(RS.Fields("volumn").value & "" = "0", "", RS.Fields("volumn").value & "")
            .Col = TblColumn.tcCANCEL: .value = IIf(RS.Fields("cancelfg").value & "" = "1", "Y", "")
            
            .Col = TblColumn.tcDONORID: .value = RS.Fields("donorid").value & ""
            .Col = TblColumn.tcSELPTID: strTmp = .value
            i = i + 1
            objPro.value = i
Skip:
            RS.MoveNext
        Loop
        
        .ReDraw = True
    End With
    
    Set RS = Nothing
    Set objPro = Nothing
End Sub

Private Sub Form_Load()
    Call InitForm
End Sub

Private Sub InitForm()
    dtpToDt.value = GetSystemDate
    dtpFrDt.value = DateAdd("d", -7, dtpToDt.value)
    optDonorCd(0).value = False
    optDonorCd(1).value = False
    chkAll.value = 1
    Call InitTable
End Sub

Private Sub InitTable()
    Call medClearTable(tblList)
    tblList.MaxRows = 20
    tblList.RowHeight(-1) = RowHeight
    tblList.Col = 1: tblList.COL2 = tblList.MaxCols
    tblList.Row = 1: tblList.Row2 = tblList.MaxRows
    tblList.BlockMode = True
    tblList.CellType = CellTypeStaticText
    tblList.TypeVAlign = TypeVAlignCenter
    tblList.TypeHAlign = TypeHAlignCenter
    tblList.BlockMode = False
End Sub

Private Sub optDonorCd_Click(Index As Integer)
    If optDonorCd(0).value Or optDonorCd(1).value Then
        chkAll.value = 0
    End If
End Sub

Private Sub tblList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    '혈액 Tag 출력
    Dim strBldNo As String
    Dim strCompo As String
    
    If Col <> TblColumn.tcBAR Then Exit Sub
    If Row < 1 Or Row > tblList.DataRowCnt Then Exit Sub
    
    With tblList
        .Row = Row
        .Col = TblColumn.TcBLOODNO: strBldNo = .value
        .Col = TblColumn.tcCOMPOCD: strCompo = .value
        
        Call TagPrint(strBldNo, strCompo)
    End With
End Sub

Private Sub TagPrint(ByVal vBldNo As String, ByVal vCompo As String)
'헌혈자 Tag 출력
'혈액번호, 제제로 쿼리해서 출력한다.
    Dim strSQL As String
    Dim RS As Recordset
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldNo As String
    Dim strCompo As String
    Dim aryData(1 To 11) As Variant
    
    strBldSrc = medGetP(vBldNo, 1, "-")
    strBldYY = medGetP(vBldNo, 2, "-")
    strBldNo = medGetP(vBldNo, 3, "-")
    strCompo = vCompo
    
    strSQL = " select a.bldsrc,a.bldyy,a.bldno,a.compocd,b.abbrnm,a.volumn,a.abo||a.rh bldabo,a.ptid,d." & F_PTNM & " ptnm, a.coldt, a.expdt,a.donorid,c.donornm,c.abo||c.rh donorabo   "
    strSQL = strSQL & " from " & T_BBS401 & " a, " & T_BBS006 & " b, " & T_BBS601 & " c, " & T_HIS001 & " d "
    strSQL = strSQL & " where " & DBW("a.bldsrc=", strBldSrc)
    strSQL = strSQL & " and " & DBW("a.bldyy=", strBldYY)
    strSQL = strSQL & " and " & DBW("a.bldno=", strBldNo)
    strSQL = strSQL & " and " & DBW("a.compocd=", strCompo)
    strSQL = strSQL & " and (a.reserved='1' or a.pherefg='1')"
    strSQL = strSQL & " and a.compocd=b.compocd"
    strSQL = strSQL & " and a.donorid=c.donorid"
    strSQL = strSQL & " and a.ptid=d.patno"
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        MsgBox "출력할 내역이 없습니다.", vbExclamation
        Set RS = Nothing
        Exit Sub
    End If

'aryData(1):혈액번호, aryData(2):혈액제제, aryData(3):용량
'aryData(4):혈액형, aryData(5):지정환자ID, aryData(6):환자명
'aryData(7):헌혈일, aryData(8):유효일, aryData(9):헌혈자
'aryData(10):헌혈자혈액형, aryData(11):바코드용 혈액번호
    aryData(1) = vBldNo
    aryData(2) = RS.Fields("abbrnm").value & ""
    aryData(3) = RS.Fields("volumn").value & ""
    aryData(4) = RS.Fields("bldabo").value & ""
    aryData(5) = RS.Fields("ptid").value & ""
    aryData(6) = RS.Fields("ptnm").value & ""
    aryData(7) = Mid(Format(RS.Fields("coldt").value & "", "####/##/##"), 3)
    aryData(8) = Mid(Format(RS.Fields("expdt").value & "", "####/##/##"), 3)
    aryData(9) = RS.Fields("donornm").value & ""
    aryData(10) = RS.Fields("donorabo").value & ""
    aryData(11) = RS.Fields("bldsrc").value & "" & RS.Fields("bldyy").value & "" & Format(RS.Fields("bldno").value & "", "000000")
    
    PrintDonorLabel aryData()
    
    Set RS = Nothing
End Sub

