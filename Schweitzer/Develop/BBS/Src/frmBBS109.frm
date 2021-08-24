VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS109 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액요청 전송"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   Icon            =   "frmBBS109.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14580
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  조회 조건"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1200
      Left            =   75
      TabIndex        =   15
      Top             =   285
      Width           =   14385
      Begin VB.TextBox txtReqId 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   9675
         TabIndex        =   23
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdReqId 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10710
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   270
         Width           =   350
      End
      Begin VB.ComboBox cboOrd 
         Height          =   300
         ItemData        =   "frmBBS109.frx":076A
         Left            =   1230
         List            =   "frmBBS109.frx":0774
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   660
         Width           =   3150
      End
      Begin VB.CheckBox chkStat 
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급처방만"
         Height          =   240
         Left            =   9675
         TabIndex        =   19
         Top             =   780
         Value           =   1  '확인
         Width           =   1230
      End
      Begin VB.CommandButton cmdPtId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Left            =   6765
         Style           =   1  '그래픽
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   675
         Width           =   360
      End
      Begin VB.TextBox txtPtId 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   5610
         TabIndex        =   4
         Text            =   "7123456"
         Top             =   675
         Width           =   1155
      End
      Begin VB.CommandButton cmdWardId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Left            =   6780
         Style           =   1  '그래픽
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   360
      End
      Begin VB.TextBox txtWardId 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   5595
         TabIndex        =   3
         Text            =   "7123456"
         Top             =   270
         Width           =   1170
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E4D5CD&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   12930
         Style           =   1  '그래픽
         TabIndex        =   5
         Tag             =   "15101"
         Top             =   555
         Width           =   1320
      End
      Begin VB.ComboBox cboInOut 
         Height          =   300
         ItemData        =   "frmBBS109.frx":0784
         Left            =   4545
         List            =   "frmBBS109.frx":0791
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   270
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   20774915
         CurrentDate     =   36838
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   315
         Left            =   2910
         TabIndex        =   1
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   20774915
         CurrentDate     =   36838
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   7155
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   7140
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   675
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.CheckBox chkDc 
         BackColor       =   &H00DBE6E6&
         Caption         =   "DC제외"
         Height          =   240
         Left            =   11025
         TabIndex        =   21
         Top             =   780
         Width           =   930
      End
      Begin MedControls1.LisLabel lblReqNm 
         Height          =   315
         Left            =   11070
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   255
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "예 정 일 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   135
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   645
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "혈액제제별"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   4545
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   675
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   8595
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "요청자"
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2715
         TabIndex        =   17
         Top             =   330
         Width           =   135
      End
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   870
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00C8CEDF&
      Caption         =   "전송(&O)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   14
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Enabled         =   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   6600
      Left            =   75
      TabIndex        =   6
      Top             =   1830
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   11642
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14737632
      MaxCols         =   45
      MaxRows         =   27
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS109.frx":07A7
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   75
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1500
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  수혈 처방 리스트"
      Appearance      =   0
   End
   Begin VB.Label lblAge 
      Caption         =   "lblAge"
      Height          =   195
      Left            =   1575
      TabIndex        =   26
      Top             =   8655
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblSex 
      Caption         =   "lblSex"
      Height          =   195
      Left            =   1560
      TabIndex        =   25
      Top             =   8370
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmBBS109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum TblColumn
    tcSEL = 1
    tcPTID
    tcPTNM
    TcABO
    tcORDNM

    tcORDDT
    tcUNITQTY
    tcREQQTY
    tcSENDQTY
    TcMESG
    tcSTATnm
    tcDCNM

    tcSTSNM
    tcWARD
    tcROOM
    tcDEPT
    tcSPCNO

    tcSTORE
    tcACCNO
    tcCENTERNM
    tcBUSSDIV
    tcORDDTDB

    tcORDNO
    tcORDSEQ
    tcSTATFG
    tcDCFG
    tcBEDINDT

    tCLegRowCol
    tcCENTERCD
    tcNOACCSSS
    tcPHERESIS
    tcSTSCD

    tcREASON
    tcDISEASE
    tcDISEASE2
    tcDISEASE3
    tcDISEASE4

    tcTime
    tcORDDIV
    tcDUPCHK
    tcREQDT
    tcDOCT
    tcTRANSDT

End Enum


Private WithEvents objMyList   As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private WithEvents objPtInfo    As frmPtInfo
Attribute objPtInfo.VB_VarHelpID = -1
'Private WithEvents mnuPopup     As Menu
'Private WithEvents mnuAddSpc    As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_REQ& = 1

Private aryLeg()
Private aryRow()
Private aryCol()

Private Sub cboDateDiv_Click()
    tblPtList.MaxRows = 0
End Sub

Private Sub cboInOut_Click()

    Select Case cboInOut.ListIndex
        Case "0"
            txtWardId = ""
            lblWardNm.Caption = ""
            txtWardId.Enabled = False
            cmdWardId.Enabled = False
    
            txtWardId.BackColor = Me.BackColor
        Case Else
           
            txtWardId = ""
            lblWardNm.Caption = ""
            txtWardId.Enabled = True
            cmdWardId.Enabled = True
        
            txtWardId.BackColor = RGB(255, 255, 255)
    End Select

End Sub

Private Sub cboInOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboTestDiv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboOrd_Click()
    tblPtList.MaxRows = 0
End Sub

Private Sub chkAccess_KeyPress(KeyAscii As Integer)
    'If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub chkQue_Click(Index As Integer)
'    If chkAccess.value = False Then
'        If Index < 2 Then
'            MsgBox "접수이전의 상태는 조회할수 없습니다.", vbInformation + vbOKOnly, "상태별조회"
'            Select Case chkQue(Index).value
'                Case 0: chkQue(Index).value = 1
'                Case 1:  chkQue(Index).value = 0
'            End Select
'        End If
'    End If
End Sub


Private Sub cmdClear_Click()
    ClearAll
    dtpFrDt.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPtId_Click()
    objPtInfo.Show vbModal
End Sub

Private Sub cmdQuery_Click()
    If cboInOut.ListIndex = 1 Then
    
        If txtWardId = "" Then
            MsgBox "병동을 선택하십시요.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Me.MousePointer = 11
    
    Call Query

    Me.MousePointer = 0

    If tblPtList.MaxRows > 0 Then
        '2001-12-17추가 : 환자별조회일 경우에만 출고전표 출력기능 제공
        If Trim(txtPtid.Text) <> "" Then
            cmdPrint.Enabled = True
        Else
            cmdPrint.Enabled = False
        End If

        tblPtList.SetFocus
    Else
        cmdPrint.Enabled = False
        MsgBox "해당자료가 없습니다", vbInformation, Me.Caption

    End If
End Sub

Private Sub cmdReqId_Click()

    Set objMyList = New clsPopUpList
    With objMyList
'        .BackColor = Me.BackColor
        .Connection = DBConn
        .FormCaption = "직원조회": .ColumnHeaderText = "사번;직원명"
'        .Width = .Width + 300: .ColSize(0) = 1000
        Call .LoadPopUp(GetSQLEmpList) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtReqId.Text = medGetP(.SelectedString, 1, ";")
            lblReqNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    Set objMyList = Nothing
End Sub

Private Sub cmdWardId_Click()

    Set objMyList = New clsPopUpList
    With objMyList
        .Connection = DBConn
        txtWardId.Text = "": lblWardNm.Caption = ""
        Select Case cboInOut.ListIndex
            Case "1"
                .FormCaption = "병동 조회": .ColumnHeaderText = "코드;코드명"
'                .Width = .Width + 700
                Call .LoadPopUp(GetSQLWardList) ', 2850, 7650) ', ObjBBSComCode.wardid)
            Case "2"
                .FormCaption = "진료과조회": .ColumnHeaderText = "코드, 코드명"
'                .Width = .Width + 300
'                .ColSize(0) = 1000
                Call .LoadPopUp(GetSQLDeptList) ', 2350, 7650) ', ObjBBSComCode.DeptCd)
        End Select
        If .SelectedString <> "" Then
            If txtWardId.Text <> medGetP(.SelectedString, 1, ";") Then
                tblPtList.MaxRows = 0
            End If
            txtWardId = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
    
            dtpFrDt.SetFocus
        Else
            txtWardId.SetFocus
        End If
    End With
    Set objMyList = Nothing
End Sub


Private Sub dtpFrDt_Change()
    tblPtList.MaxRows = 0
End Sub

Private Sub dtpFrDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpToDt_Change()
    tblPtList.MaxRows = 0
End Sub

Private Sub dtpToDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objAccess   As clsBBSAccess
    Dim objBBSsql   As clsGetSqlStatement
    Dim RS          As Recordset
    Dim Rsord       As Recordset
    Dim ii          As Long

    Set objPtInfo = New frmPtInfo
    Set objAccess = New clsBBSAccess
    Set objBBSsql = New clsGetSqlStatement
    Set Rsord = objBBSsql.Get_CompoRecordSet

    '검사항목
    With Rsord
        cboOrd.Clear
        cboOrd.AddItem "전체혈액제제"
        For ii = 1 To .RecordCount
             cboOrd.AddItem .Fields("compocd").value & "" & Space(2) & .Fields("abbrnm").value & ""
            .MoveNext
        Next ii
    End With
    cmdPrint.Visible = False
    dtpFrDt = DateAdd("d", -1, GetSystemDate)
    dtpToDt = DateAdd("d", 2, GetSystemDate)

    cboInOut.ListIndex = 0
    chkStat.value = False
    Call ClearAll
    cmdPrint.Visible = True
    Me.Show

    Set RS = Nothing
    Set Rsord = Nothing
    Set objAccess = Nothing
    Set objBBSsql = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objPtInfo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub


Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_REQ
            With tblPtList
                .Row = .ActiveRow
                .Col = TblColumn.tcACCNO
                frmBBS204.txtAccNo = .value
                frmBBS204.Show 1
            End With
    End Select
End Sub

'Private Sub mnuAddSpc_Click()
'
'    With tblPtList
'        .Row = .ActiveRow
'        .Col = TblColumn.tcACCNO
'        frmBBS204.txtAccNo = .value
'        frmBBS204.Show 1
'    End With
'End Sub

Private Sub objPtInfo_Click(ByVal isSELECT As Boolean, ByVal ptInfo As clsPtInformation)
    If isSELECT = False Then
        'txtPtId.SetFocus
    Else
        If txtPtid <> ptInfo.PtId Then tblPtList.MaxRows = 0
        txtPtid = ptInfo.PtId
        lblPtNm.Caption = ptInfo.ptnm
    End If
End Sub

Private Function CanSelect(ByVal Col As Long, ByVal Row As Long) As Boolean

    Dim objSQL   As clsQueryOrder
    Dim CenterCd As String
    Dim noaccess As String
    Dim pheresis As String
    Dim sel      As String
    Dim spcno    As String
    Dim KeepOur  As Long
    Dim i        As Long

    '중간에 나가면 불가능한 것이다.....
    CanSelect = False

    Set objSQL = New clsQueryOrder
    KeepOur = objSQL.GetKeepHour
    Set objSQL = Nothing



    With tblPtList
        '검체번호가 있는 것만 대상
        '접수번호가 없는 것(처방미접수)만 대상
        '보관장소가 없는 것(검체미접수)만 대상
        'D/C처방은 제외
        '검체보관시간 지나지 않은것만 대상
        'irradiation 처방이 아닌 처방만 대상

        .Row = Row

        '건물코드가 다르면 접수할수 없다.
'        .Col = TblColumn.tcCENTERCD: centercd = .value
'        If centercd <> ObjSysInfo.BuildingCd Then Exit Function

        'D/C발생한 처방에 대해서는 접수할수 없다.
        .Col = TblColumn.tcDCFG
        If .value = "1" Then Exit Function

        '검체번호가 없으면 접수할수 없다.
        .Col = TblColumn.tcSPCNO
        If .value = "" Then Exit Function

        '접수번호가 있으면 접수할수 없다.
'        .Col = TblColumn.tcACCNO
'        If .value <> "" Then Exit Function

        '상태가 처방인것은 접수할수 없다.
        .Col = TblColumn.tcSTSNM
        If .value = "완결" Then Exit Function

        '72시간이 지난 검체는 접수할수 없다.
'        .Col = TblColumn.tcTime
'        If Val(.value) > KeepOur Then Exit Function

        'IRRAdiation 처방은 접수할수 없다.
        .Col = TblColumn.tcORDDIV
        If .value = "Z" Then Exit Function
    End With

    CanSelect = True
End Function

Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    Static BfRow As Long
    Dim clrBackOdd As Long
    Dim clrForeOdd As Long
    Dim clrBackEven As Long
    Dim clrForeEven As Long

    Dim CenterCd As String
    Dim noaccess As String
    Dim pheresis As String
    Dim sel      As String
    Dim spcno    As String
    Dim i        As Long

    If Row < 1 Then Exit Sub
    If Row > tblPtList.MaxRows Then Exit Sub
'    If fraStore.Visible = True Then Exit Sub


    With tblPtList

        Call .GetOddEvenRowColor(clrBackOdd, clrForeOdd, clrBackEven, clrForeEven)

        If BfRow <> Row Then
            .Row = BfRow: .Row2 = BfRow
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            If (BfRow Mod 2) = 0 Then
                .BackColor = clrBackEven
            Else
                .BackColor = clrBackOdd
            End If
            .BlockMode = False
        End If

        .Row = Row: .Row2 = Row
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .BackColor = .SelBackColor
        .BlockMode = False

        BfRow = Row
    End With


    With tblPtList
        Select Case Col
            Case TblColumn.tcSTORE
                .Row = Row
                .Col = TblColumn.tcNOACCSSS: noaccess = .value
                .Col = TblColumn.tcCENTERCD: CenterCd = .value

                '-------------------아직 검체접수가 안된 것만 처리.
                If noaccess = "0" Then Exit Sub
                '---------------------우리 센터에서 처리할 수 없다.
                If CenterCd <> ObjSysInfo.BuildingCd Then Exit Sub

            Case TblColumn.tcSEL
                .Col = Col
                .Row = Row
                If .CellType <> CellTypeCheckBox Then Exit Sub

                If CanSelect(Col, Row) = False Then
                    .Col = Col
                    .Row = Row
                    .value = 0
                    Exit Sub
                End If

                'pheresis 처방일경우는 처방한건당체크가 가능하다.....
                .Row = Row
                .Col = TblColumn.tcSPCNO: spcno = .value
                .Col = TblColumn.tcSEL:   sel = .value
'                .value = IIf(sel = 1, 0, 1)
'                If pheresis <> "1" Then

                For i = 1 To .MaxRows
                    If i <> Row Then
                        .Row = i
                        .Col = TblColumn.tcORDDIV
                        'irradiation처방인것을 구분하기 위해서......
                        If .value = C_WORKAREA Then
                            .Col = TblColumn.tcSPCNO
                            '같은 채혈번호를 가질때...
                            If spcno = .value Then
                                '접수번호가 ""(접수않된거만)....
                                .Col = TblColumn.tcACCNO
                                If .value = "" Then
                                    '접수처리가 가능한 혈액에 대해서만....
                                    .Col = Col
                                    If .CellType = CellTypeCheckBox Then
                                        .Col = TblColumn.tcSEL
                                        .value = IIf(sel = 1, 0, 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
'                End If
        End Select
    End With
End Sub

Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim objDisease   As New S2BBS_Library.clsDisease
    Dim objSQL       As New clsQueryOrder
    Dim RS           As Recordset
    Dim strAccNo     As String  '접수번호
    Dim strSpcNo     As String  '검체번호
    Dim strStore     As String
    Dim strRack      As String
    Dim strRow       As String
    Dim strCol       As String
    Dim strCenter    As String
    Dim StrWARD      As String  '병동
    Dim STRDEPT      As String  '진료과
    Dim strReason    As String  '수혈사유
    
    Dim strDisea1    As String  '진단명
    Dim strDisea2    As String  '진단명2
    Dim strDisea3    As String  '진단명3
    Dim strDisea4    As String  '진단명4
    
    Dim strTime      As String
    Dim coldttm      As String
    Dim strDiseaDisp As String
    Dim strReqDt     As String
    Dim strMesg      As String

    'IRRADIATION처방인경우..
    Dim strPtid      As String
    Dim strOrdDt     As String
    Dim strOrdNo     As String
    
    Dim i            As Long
    Dim strtip       As String
    
    '감염여부
    Dim sICSStr     As String
    
    
    If Row < 1 Then Exit Sub

    With tblPtList
        Call .SetTextTipAppearance("굴림체", 9, False, False, &HFFFFC0, vbBlack)
        .Row = Row
        .Col = TblColumn.tcACCNO:       strAccNo = .value
        .Col = TblColumn.tcWARD:        StrWARD = .value
        .Col = TblColumn.tcDEPT:        STRDEPT = .value
        .Col = TblColumn.tcTime:        strTime = .value
        .Col = TblColumn.tcREQDT:       strReqDt = .value
        .Col = TblColumn.TcMESG:        strMesg = .value
        .Col = TblColumn.tcPTID:        strPtid = .value
        .Col = TblColumn.tcORDDT:       strOrdDt = .value
        .Col = TblColumn.tcORDNO:       strOrdNo = .value
        
        '진단명을 구한다.
        objDisease.Clear
        objDisease.PtId = strPtid
        objDisease.orddt = strOrdDt
        objDisease.ordno = strOrdNo
        
        If objDisease.GetDisease Then
            i = 0
            Do
                If objDisease.EOF Then Exit Do

                If objDisease.DiseaseCd <> "" Then
                    i = i + 1
                    Select Case i
                        Case 1: strDisea1 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 2: strDisea2 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 3: strDisea3 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 4: strDisea4 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End Select
                End If
                objDisease.MoveNext
            Loop
        End If
        
        strDiseaDisp = "  진 단 명 : " & strDisea1
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea2
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea3
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea4
        '수혈사유
        strReason = objSQL.GetTransReason(strPtid, strOrdDt, strOrdNo): If strReason = "" Then strReason = "(없음)"
        '검체정보(위치,번호)
        Call objSQL.GetSpcNoAndStore(strPtid, strSpcNo, strRack, strRow, strCol, strCenter)
        
        If strRow = "0" Then strRow = ""
        If strCol = "0" Then strCol = ""

        If strSpcNo = "" Then
            strStore = ""
        Else
            If strRack = "" Then
                strStore = ""
            Else
                strStore = strRack & "(" & strRow & "," & strCol & ")"
            End If
        End If
        '경과시간
        If strSpcNo <> "" Then
            Set RS = New Recordset
            RS.Open objSQL.Get_spcTime(medGetP(strSpcNo, 1, "-"), medGetP(strSpcNo, 2, "-")), DBConn
            If Not RS.EOF Then
                If Len(RS.Fields("coltm").value & "") = 4 Then
                    coldttm = RS.Fields("coltm").value & "" & "00"
                    coldttm = Format(RS.Fields("coldt").value & "", "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                Else
                    coldttm = Format(RS.Fields("coldt").value & "", "0###-##-##") & " " & Format(RS.Fields("coltm").value & "", "0#:##:##")
                End If
                strTime = DateDiff("h", coldttm, GetSystemDate) & "시간"
            End If
        End If
        
        sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
        
        strtip = "  접수번호 : [" & strAccNo & "], 검체번호 : [" & strSpcNo & "]," & _
                 "  보관장소 : [" & strStore & "]" & vbNewLine & _
                 "  경과시간 : " & strTime & vbNewLine & _
                 "  병동/과  : " & StrWARD & "/" & STRDEPT & vbNewLine & _
                 "  수혈사유 : " & strReason & vbNewLine & _
                 "  예정일시 : " & strReqDt & vbNewLine & _
                 "  처방비고 : " & strMesg & vbNewLine & _
                 strDiseaDisp
        
        If sICSStr <> "" Then
            strtip = strtip & vbNewLine & " 감염여부 : " & sICSStr
        End If
        
        
        TipWidth = 6200
        MultiLine = 1
        TipText = vbNewLine & strtip & vbNewLine
        ShowTip = True
    End With

    Set RS = Nothing
    Set objSQL = Nothing
    Set objDisease = Nothing

End Sub

Private Sub txtPtId_GotFocus()
    txtPtid.tag = txtPtid
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = cmdClear.name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = cmdExit.name Then Exit Sub

    If txtPtid.tag = txtPtid Then Exit Sub
    If SearchPTINFO = False Then txtPtid.SetFocus
End Sub

Private Function SearchPTINFO() As Boolean
    SearchPTINFO = Search_PtInfo
    tblPtList.MaxRows = 0
End Function

Private Sub txtWardId_GotFocus()
    txtWardId.tag = txtWardId
    txtWardId.SelStart = 0
    txtWardId.SelLength = Len(txtWardId)
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If SearchWard = True Then
            txtWardId.tag = txtWardId
            SendKeys "{TAB}"
        Else
            txtWardId.SelStart = 0
            txtWardId.SelLength = Len(txtWardId)
        End If
    End If
End Sub

Private Sub txtWardId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = cmdClear.name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = cmdExit.name Then Exit Sub

    If txtWardId.tag = txtWardId Then Exit Sub
    If SearchWard = False Then txtWardId.SetFocus
End Sub

Private Function SearchWard() As Boolean

    SearchWard = Search_Ward

    tblPtList.MaxRows = 0
End Function

Private Sub ClearAll()

    txtWardId = ""
    lblWardNm.Caption = ""
    txtPtid = ""
    lblPtNm.Caption = ""
    tblPtList.MaxRows = 0
    cboOrd.ListIndex = 0
    Call ICSPatientMark
    
End Sub

Private Function Search_PtInfo() As Boolean
    Dim objPtInfo As clsPtInformation
    Dim DrRS      As Recordset
    Dim ii        As Long
    Dim strLng    As String

    If txtPtid = "" Then
        lblPtNm.Caption = ""
        Search_PtInfo = True
    Else
        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
            strLng = strLng & "0"
        Next ii
        If Len(Trim(txtPtid.Text)) <> BBS_PTID_LENGTH Then
            txtPtid.Text = Format(txtPtid.Text, strLng & "#")
        End If
        
        Set objPtInfo = New clsPtInformation
        Set DrRS = New Recordset
        DrRS.Open objPtInfo.Get_Ptid(txtPtid), DBConn
        If DrRS.EOF = False Then
            With objPtInfo
                .BedPt_Chk txtPtid.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
                If .PtDiv = "BED" Then
                    txtPtid = .PtId
                    lblPtNm.Caption = .ptnm
                    lblSex.Caption = .Sex
                    lblAge.Caption = .Age
                Else
                    txtPtid = .PtId
                    lblPtNm.Caption = .ptnm
                    lblSex.Caption = .Sex
                    lblAge.Caption = .Age
                End If
            End With
            Search_PtInfo = True
        Else
            MsgBox "해당되는 환자가 없습니다. 확인후 조회하세요.", vbInformation + vbOKOnly, Me.Caption
            txtPtid = ""
            lblPtNm.Caption = ""
            Search_PtInfo = False
        End If
        Set DrRS = Nothing
        Set objPtInfo = Nothing
    End If
    Call ICSPatientMark(txtPtid.Text, enICSNum.BBS_ALL)
    
End Function

Private Function Search_Ward() As Boolean
    If txtWardId = "" Then
        lblWardNm.Caption = ""
        Search_Ward = True
    Else
        txtWardId.Text = UCase(txtWardId.Text)
        lblWardNm.Caption = GetWardNm(txtWardId.Text)
        If lblWardNm.Caption = "" Then
            MsgBox "해당되는 자료가 없습니다. 확인후 입력하세요.", vbInformation + vbOKOnly, "병동입력"
            lblWardNm.Caption = ""
            Search_Ward = False
        End If
    End If
End Function

Private Function CompleteOrderChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Boolean
    Dim objXM As New clsCrossMatching
    Dim A_Cnt As Long   'Assign수량
    Dim C_Cnt As Long   'Assign Cancel 수량
    Dim O_Cnt As Long   '출고수량
    Dim R_Cnt As Long   '반환수량
    Dim X_Cnt As Long   '폐기수량
    Dim T_Cnt As Long   '총Assign 수량


    'CompleteOrderChk=True이면 완결처방
    'CompleteOrderChk=미완결처방
    CompleteOrderChk = False
    If accdt <> "" Then
        
        With objXM
            .Assign_Cnt accdt, Val(accseq)
            A_Cnt = .AssignCnt
            C_Cnt = .CancelCnt
            O_Cnt = .OutCnt
            R_Cnt = .RetCnt
            X_Cnt = .ExpCnt
        End With
        Set objXM = Nothing
                
        T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt
        
        If unitqty = T_Cnt Then
            CompleteOrderChk = True
        End If
    End If
End Function

Private Function CompleteRequestChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Long
    
    Dim objXM As New clsCrossMatching

    'CompleteOrderChk=True이면 완결처방
    'CompleteOrderChk=미완결처방
    CompleteRequestChk = 0
    If accdt <> "" Then

        With objXM
            .Request_Cnt accdt, Val(accseq)
            CompleteRequestChk = unitqty - .RequestCnt
        End With
        Set objXM = Nothing

    End If
    
End Function

Private Function IRR_DUPchk(ByVal PtId As String, ByVal orddt As String) As Boolean
    Dim ii As Integer
    Dim strTmp As String

    strTmp = PtId & COL_DIV & orddt
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcDUPCHK
            If .value = strTmp Then
                IRR_DUPchk = True
                Exit Function
            End If
        Next
    End With
End Function

Private Function GetABO(ByVal PtId As String) As String
'혈액형,부작용,감염정보,상병코드,상병을 조회한다.
    Dim ObjABO As New clsABO


    With ObjABO
        .PtId = PtId
        If .GetABO = True Then
            GetABO = .ABO & .Rh
        Else
            GetABO = ""
        End If
    End With
    Set ObjABO = Nothing

End Function

Private Sub Query()
    Dim i           As Long
    Dim RS        As Recordset
    Dim QueryOrder  As clsQueryOrder

    Dim accno       As String
    Dim status      As String
    Dim spcno       As String
    Dim storeleg    As String
    Dim storerow    As String
    Dim storecol    As String
    Dim center      As String

    Dim inout       As String
    Dim MaxRowCnt   As Long
    Dim TestDiv     As String
    Dim lngReqCnt   As Long
    Dim objPrgBar   As clsProgress


    '윗줄과 같은내용이면 글자를 감추기 위한변수들
    Dim bkPtId      As String
    Dim bkReqDt     As String
    Dim bkOrdDt     As String
    Dim bkRoomid    As String
    Dim bkWard      As String
    Dim bkDept      As String

    Dim strDc       As String

    tblPtList.MaxRows = 0

    Set QueryOrder = New clsQueryOrder

    If cboOrd.ListIndex <> 0 Then TestDiv = medGetP(cboOrd.Text, 1, " ")
    
    '상태별 조회
    QueryOrder.stscd = "'" & BBSOrdStatus.stsACCESS & "','" & _
                             BBSOrdStatus.stsREQUEST & "','" & _
                             BBSOrdStatus.stsINPROCESS & "','" & _
                             BBSOrdStatus.stsEnd & "'"   ' "'2'"
    Select Case cboInOut.ListIndex
        Case 0: inout = ""
        Case 1: inout = "2"
        Case 2: inout = "1"
    End Select
    If chkDc.value = "1" Then strDc = "1"

    Set RS = QueryOrder.QueryOrder(Format(dtpFrDt, PRESENTDATE_FORMAT), Format(dtpToDt, PRESENTDATE_FORMAT), chkStat.value, txtPtid.Text, inout, strDc, txtWardId, TestDiv)

    If RS Is Nothing Then
        Set RS = Nothing
        Set QueryOrder = Nothing
        Exit Sub
    End If
    
    Set objPrgBar = New clsProgress

    With objPrgBar
        .Container = Me
        .Width = LisLabel3.Width
        .Left = LisLabel3.Left
        .Top = LisLabel3.Top
        .Height = 280
        .Message = "수혈처방내역을 검색중입니다..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = LisLabel3.Width
'        .XPos = LisLabel3.Left
'        .YPos = LisLabel3.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "수혈처방내역을 검색중입니다..."
'        .value = 1
    End With

    objPrgBar.Min = 1
    objPrgBar.Max = RS.RecordCount


    With tblPtList
        bkPtId = ""
        .ReDraw = False
        For i = 1 To RS.RecordCount

            objPrgBar.value = i
            MaxRowCnt = MaxRowCnt + 1: .MaxRows = MaxRowCnt: .Row = MaxRowCnt
            accno = Trim(RS.Fields("accdt").value & "") & "-" & _
                    Val(Trim(RS.Fields("accseq").value & ""))
            If accno = "-0" Then accno = ""
            '2001-11-29 추가 : 요청가능수량 구하기. 0 이면 Skip
            lngReqCnt = CompleteRequestChk(RS.Fields("accdt").value & "", _
                                           RS.Fields("accseq").value & "", _
                                           RS.Fields("unitqty").value & "")
            .Col = TblColumn.tcACCNO:       .value = accno
            .Col = TblColumn.tcPTID:        .value = RS.Fields("ptid").value & ""
            .Col = TblColumn.tcPTNM:        .value = GetPtNm(RS.Fields("ptid").value & "")
            .Col = TblColumn.tcORDNM:       .value = RS.Fields("testnm").value & ""
            .Col = TblColumn.tcORDDT:       .value = Format(RS.Fields("orddt").value & "", "####-##-##")
            .Col = TblColumn.tcUNITQTY:     .value = RS.Fields("unitqty").value & "": .ForeColor = DCM_Black
            .Col = TblColumn.tcREQQTY:      .value = lngReqCnt: .ForeColor = DCM_Blue
            .Col = TblColumn.tcSENDQTY:     .value = lngReqCnt: .TypeIntegerMax = lngReqCnt: .TypeIntegerMin = 0: .ForeColor = DCM_Red
            .Col = TblColumn.tcREQDT:       .value = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
            .Col = TblColumn.tcDOCT:        .value = RS.Fields("majdoct").value & ""
            .Col = TblColumn.tcWARD:        .value = RS.Fields("wardid").value & ""
            .Col = TblColumn.tcROOM:        .value = RS.Fields("hosilid").value & ""
            .Col = TblColumn.tcDEPT:        .value = RS.Fields("deptcd").value & ""
            .Col = TblColumn.tcBUSSDIV:     .value = RS.Fields("bussdiv").value & ""
            .Col = TblColumn.tcORDDTDB:     .value = RS.Fields("orddt").value & ""
            .Col = TblColumn.tcORDNO:       .value = Val(RS.Fields("ordno").value & "")
            .Col = TblColumn.tcORDSEQ:      .value = Val(RS.Fields("ordseq").value & "")
            .Col = TblColumn.tcSTATFG:      .value = RS.Fields("statfg").value & ""
            .Col = TblColumn.tcSTATnm:      .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", ""): .ForeColor = vbRed: .FontBold = True
            .Col = TblColumn.tcBEDINDT:     .value = RS.Fields("bedindt").value & "" & ""
            .Col = TblColumn.tcDCFG:        .value = RS.Fields("dcfg").value & ""
            .Col = TblColumn.tcDCNM:        .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", ""): .ForeColor = vbBlue: .FontBold = True
            .Col = TblColumn.tcPHERESIS:    .value = RS.Fields("testdiv").value & ""
            .Col = TblColumn.tcSTSCD:       .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.TcMESG: .value = RS.Fields("mesg").value & ""
            'dup check
            .Col = TblColumn.tcDUPCHK: .value = RS.Fields("ptid").value & "" & COL_DIV & RS.Fields("orddt").value & ""
            '중복되는 값은 안보이게...
            If bkPtId <> RS.Fields("ptid").value & "" Then
                bkPtId = RS.Fields("ptid").value & ""
                bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                bkRoomid = RS.Fields("hosilid").value & ""
                bkWard = RS.Fields("wardid").value & ""
                bkDept = RS.Fields("deptcd").value & ""
            Else
                .Row = i - 1
                .Col = TblColumn.tcWARD: bkWard = .value
                .Col = TblColumn.tcDEPT: bkDept = .value
                .Row = i
                .Col = TblColumn.tcPTID: .ForeColor = .BackColor
                .Col = TblColumn.tcPTNM: .ForeColor = .BackColor
                If bkWard = RS.Fields("wardid").value & "" Then .Col = TblColumn.tcWARD: .ForeColor = .BackColor
                If bkDept = RS.Fields("deptcd").value & "" Then .Col = TblColumn.tcDEPT: .ForeColor = .BackColor
                
                If bkRoomid = RS.Fields("hosilid").value & "" Then
                    .Col = TblColumn.tcROOM: .ForeColor = .BackColor
                Else
                    bkRoomid = RS.Fields("hosilid").value & ""
                End If
                
                If bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00") Then
                    .Col = TblColumn.tcREQDT: .ForeColor = .BackColor
                Else
                    bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                End If
                
                If bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##") Then
                    .Col = TblColumn.tcORDDT: .ForeColor = .BackColor
                Else
                    bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                End If
            End If
            '요청할 수 있는 건인지
            If lngReqCnt > 0 Then
                .Row = MaxRowCnt: .Col = TblColumn.tcSEL:
                .CellType = CellTypeCheckBox: .TypeCheckCenter = True
            Else
                .Row = MaxRowCnt: .Col = TblColumn.tcSEL
                .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignCenter
                .Col = TblColumn.tcSTSNM: .Col = TblColumn.tcSEL: .Text = "√": .ForeColor = vbRed
            End If
            RS.MoveNext
        Next i
        Set objPrgBar = Nothing
        If .DataRowCnt > 0 Then GetBatchABO
        .ReDraw = True
    End With

    Set RS = Nothing
    Set QueryOrder = Nothing
End Sub

Private Function SaveCheckNotAuto() As Boolean
'보관장소의 입력이 되었는지 체크한다.
    Dim SavePos    As String
    Dim SaveTF     As String
    Dim DcFg       As String
    Dim strRowCol  As String
    Dim strCol     As String
    Dim ii As Integer

    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                .Col = TblColumn.tcSTORE: SavePos = .value
                If SavePos <> "" Then
                    SaveCheckNotAuto = True
                Else
                    SaveCheckNotAuto = False
                    Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function Save_Check() As Boolean

    If Collect_Cnt = 0 Then
        '접수하고자 하는 건수를 구한다
        MsgBox "접수대상항목이 없습니다.", vbCritical + vbOKOnly, Me.Caption
        Exit Function
    End If
    Save_Check = True

End Function

Private Sub cmdCollect_Click()
    Dim objBg          As clsBloodRequest
    Dim RS             As Recordset
    Dim strColDt       As String
    Dim strColTm       As String
    Dim strAccDt       As String
    Dim lngAccNo       As Long
    Dim ii             As Integer

'    요청전송을 위한 변수들
    Dim strCenterCd As String
    Dim strPtid     As String
    Dim strOrdDt    As String
    Dim strTmp      As String
    Dim strSpcYYR   As String
    Dim strFullSpc  As String
    Dim strLeg      As String
    Dim pheresis    As String
    Dim store_cnt   As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngSpcNoR   As Long
    Dim lngOrdseq   As Long
    Dim lngOrdNo    As Long
    Dim blnSave     As Boolean


    Dim SSQL        As String

    If Save_Check = False Then Exit Sub
    
    If Trim(txtReqId.Text) = "" Then
        MsgBox "요청자 ID를 반드시 입력하세요.", vbInformation
        Exit Sub
    End If

    Set objBg = New clsBloodRequest

    Me.MousePointer = 11
    strCenterCd = ObjSysInfo.BuildingCd         '센터코드
    strColDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strColTm = Format(GetSystemDate, PRESENTTIME_FORMAT)


On Error GoTo Save_Spc_Error

    DBConn.BeginTrans

    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL

            If Val(.value) = 1 Then
                .Col = TblColumn.tcSENDQTY
                If Val(.value) > 0 Then
                    .Col = TblColumn.tcDCFG

                    .Col = TblColumn.tcPTID:     strPtid = .value
                    .Col = TblColumn.tcSPCNO:    strSpcYYR = Mid(.value, 1, 2)
                                                 lngSpcNoR = Val(Mid(.value, 4))
                                                 strFullSpc = strSpcYYR & CStr(lngSpcNoR)
                    .Col = TblColumn.tcPHERESIS: pheresis = IIf(.value = "1", "1", "0")

                    .Col = TblColumn.tcORDDT:    strOrdDt = Mid(.value, 1, 4) & Mid(.value, 6, 2) & Mid(.value, 9, 2)
                    .Col = TblColumn.tcORDNO:    lngOrdNo = Val(.value)
                    .Col = TblColumn.tcORDSEQ:   lngOrdseq = Val(.value)

                    .Col = TblColumn.tcACCNO:    strAccDt = medGetP(.value, 1, "-")
                    .Col = TblColumn.tcACCNO:    lngAccNo = Val(medGetP(.value, 2, "-"))
                
                    SSQL = objBg.Set_UpdateL101(strPtid, strOrdDt, CStr(lngOrdNo))
                    DBConn.Execute SSQL

                    SSQL = objBg.Set_UpdateL102(strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq), strAccDt, CStr(lngAccNo))
                    DBConn.Execute SSQL

                    SSQL = objBg.Set_BBS202(strAccDt, lngAccNo, txtReqId.Text)
                    DBConn.Execute SSQL

                    .Col = TblColumn.tcSENDQTY
                    SSQL = objBg.Insert_BBS204(strAccDt, lngAccNo, txtReqId.Text, Val(.value))
                    DBConn.Execute SSQL

                    '검체번호는 있는 데 처방 접수가 없는 경우
                    '성분헌혈이 아닌경우는 검체 해당자료는 저장하지 않는다.
                    '이미 검체가 지정되어있는 경우는 검체보관장소를 update 해주지 않는다.

                    lngAccNo = lngAccNo + 1
                    blnSave = True
                End If
            End If

        Next ii
    End With

    DBConn.CommitTrans
    Call Query

    Me.MousePointer = 0
    MsgBox "수혈요청이 전송되었습니다.", vbInformation, "전송"
    Set objBg = Nothing
    Exit Sub

Save_Spc_Error:

    DBConn.RollbackTrans
    Me.MousePointer = 0
    MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "수혈요청오류"
    Set objBg = Nothing
End Sub


Private Function Collect_Cnt() As Long
    Dim strTmp As String
    Dim strCollect As String        '접수여부...
    Dim strGather As String         '채혈여부...
    Dim store_cnt As Integer
    Dim ii As Integer

    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                Collect_Cnt = Collect_Cnt + 1
                .Col = TblColumn.tcPTID
                If .value <> strTmp Then
                    store_cnt = store_cnt + 1
                End If
                strTmp = .value
            End If
        Next
    End With

End Function

Private Sub cmdPrint_Click()
    If tblPtList.MaxRows <= 0 Then
        MsgBox "먼저 처방내역 또는 요청내역을 조회한 후 출력하세요.", vbInformation, "출고전표 재발행"
        Exit Sub
    End If
    Call PrintDeliveryList(True)

End Sub
Private Sub PrintDeliveryList(Optional ByVal blnReprint As Boolean = False)

'출력하자.....크리스탈
    Dim strPtid As String, strPtNm As String, strABO As String, STRUNIT As String, strReqDt As String
    Dim StrWARD As String, STRDEPT As String, strOrdNm As String, STRDISEA As String
    Dim strTmp  As String, strDoct As String, strTransDt As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim intFNum    As Integer
    Dim ii         As Integer
    Dim jj         As Integer
    Dim lngCnt     As Long

    If tblPtList.MaxRows = 0 Then Exit Sub
    Me.MousePointer = 11
    lngCnt = 0
    
    STRDISEA = ""
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            If ii = 1 Then
                .Col = TblColumn.tcPTID:    strPtid = Trim(.value)
                .Col = TblColumn.tcPTNM:    strPtNm = Trim(.value)
                .Col = TblColumn.TcABO:     strABO = Trim(.value)
                .Col = TblColumn.tcREQDT:   strReqDt = Trim(.value)
                .Col = TblColumn.tcDISEASE + 2: STRDISEA = Trim(.value)
                .Col = TblColumn.tcWARD:    StrWARD = Trim(.value)
                .Col = TblColumn.tcDEPT:    STRDEPT = Trim(.value)
                
                .Col = TblColumn.tcDOCT:    strDoct = Trim(.value)
                
                .Col = TblColumn.tcTRANSDT: strTransDt = Trim(.value)
                

                '진단명
                
                If STRDISEA <> "" Then
                    .Col = TblColumn.tcDISEASE2
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                    .Col = TblColumn.tcDISEASE3
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                    .Col = TblColumn.tcDISEASE4
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                End If
            End If
            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
        Next ii
    End With
    
    strDoct = GetEmpNm(strDoct)
    StrWARD = GetWardNm(StrWARD)
    STRDEPT = GetDeptNm(STRDEPT)
    
    Call PrintIntionlize
    PrintHeader_Trans strPtNm, StrWARD, strPtid, lblSex.Caption & "/" & lblAge.Caption, STRDISEA, strABO, _
                      "", "", strDoct, STRDEPT
    Me.MousePointer = 0
End Sub

Private Sub PrintIntionlize()
    PrtLeft = 5
    LineSpace = 6
    lngCurYPos = 10
    
    
    Printer.Font = "굴림체"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    

    Twidth = Printer.ScaleWidth

    LastLineYpos = Printer.ScaleHeight             '마지막라인Y위치

End Sub


Private Sub PrintHeader_Trans(ByVal pPtNm As String, ByVal pWard As String, ByVal pPtId As String, _
                        ByVal pSex As String, ByVal pDise As String, ByVal pABO As String, _
                        ByVal pTrans As String, ByVal pIM As String, ByVal pDoct As String, ByVal pDept As String)
    Dim lngX1 As Long
    Dim lngX2 As Long
    Dim lngX3 As Long
    
    
    
    lngX1 = 10
    lngX2 = lngX1 + Printer.TextWidth("성    명 : ")
    lngX3 = lngX1 + 70
    
    Printer.FontSize = 16: Printer.FontBold = True
    Call Print_Setting("수혈 요청 및 출고 전표", PrtLeft, lngCurYPos, Twidth, "C", "C", False)
    Printer.FontSize = 13: Printer.FontBold = False
    
    lngCurYPos = lngCurYPos + 20
'     Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
    
    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos + 70), , B
    
    lngCurYPos = lngCurYPos + LineSpace
    Call Print_Setting("성    명 : " & pPtNm, lngX1, LineSpace, , , "C", False)
    Call Print_Setting("병    동 : " & pWard, lngX3, LineSpace, , , "C", False)
    
    Call Print_Setting("   혈액형 ", 130, LineSpace, , "L", "C", False)
    
    
    
    
    
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("등록번호 : " & pPtId, lngX1, LineSpace, , , "C", False)
    Printer.FontBold = True: Printer.FontSize = 40
    Call Print_Setting(pABO, 135, LineSpace, , , "C", False)
    Printer.FontBold = False: Printer.FontSize = 13
    
    Call Print_Setting("성별/나이 : " & pSex, lngX3, 10, , , "C", False)
    lngCurYPos = lngCurYPos + 10
    
    Call Print_Setting("진 단 명 : " & pDise, lngX1, 10, , , "C", False)
    
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("수 혈 력 :     □ 무      □ 유 " & pTrans, lngX1, 10, , , "C", False)
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("임 신 력 :     □ 무      □ 유  (     주)" & pIM, lngX1, 10, , , "C", False)
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("담당의사 : " & pDoct, lngX1, 10, , , "C", False)
    Call Print_Setting("진 료 과 : " & pDept, lngX3, 10, , , "C", False)
    
    
    
    lngCurYPos = lngCurYPos + 20
    
    
    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
    Dim ii As Integer
    
    lngCurYPos = lngCurYPos + 2
    
    For ii = 1 To 12
        Printer.Line (PrtLeft, lngCurYPos + 8 * ii)-(Twidth - PrtLeft, lngCurYPos + 8 * ii)
    Next


'혈액불출
    Printer.Line (PrtLeft, lngCurYPos - 2)-(PrtLeft, lngCurYPos + 8 * 12)
    
    '혈액번호
    Printer.Line (lngX2, lngCurYPos + 8)-(lngX2, lngCurYPos + 8 * 12)
    '혈액종류
    Printer.Line (lngX2 + 30, lngCurYPos + 8)-(lngX2 + 30, lngCurYPos + 8 * 12)
    '혈액형
    Printer.Line (lngX2 + 45, lngCurYPos + 8)-(lngX2 + 45, lngCurYPos + 8 * 12)
    '채혈자
    Printer.Line (lngX2 + 60, lngCurYPos + 8)-(lngX2 + 60, lngCurYPos + 8 * 12)
    
    '수혈시작시간
    
    Printer.Line (lngX2 + 75, lngCurYPos + 8)-(lngX2 + 75, lngCurYPos + 8 * 12)
    
    Printer.Line (lngX2 + 90, lngCurYPos - 2)-(lngX2 + 90, lngCurYPos + 8 * 12)
    
    
    
    
    Printer.Line (lngX2 + 105, lngCurYPos + 8)-(lngX2 + 105, lngCurYPos + 8 * 12)
    
    
    
    '수혈끝시간
    Printer.Line (lngX2 + 120, lngCurYPos + 8)-(lngX2 + 120, lngCurYPos + 8 * 12)
    'Dr
    Printer.Line (lngX2 + 130, lngCurYPos + 8)-(lngX2 + 130, lngCurYPos + 8 * 12)
    'Nr
    Printer.Line (lngX2 + 140, lngCurYPos + 8)-(lngX2 + 140, lngCurYPos + 8 * 12)
    '수혈부작용
    'Printer.Line (lngX2 + 165, lngCurYPos + 8)-(lngX2 + 142, lngCurYPos + 8 * 12)
    
    '마지막
    Printer.Line (Twidth - PrtLeft, lngCurYPos - 2)-(Twidth - PrtLeft, lngCurYPos + 8 * 12)
    
    Printer.FontSize = 10
    
    Call Print_Setting("혈액불출기록", PrtLeft, 8, , , "C", False)
    Call Print_Setting("수혈기록", lngX2 + 90, 8, , , "C", False)
    
    lngCurYPos = lngCurYPos + LineSpace
    
    Call Print_Setting("혈액불출시간", PrtLeft, 12, lngX2 - PrtLeft, "C", "C", False)
    Call Print_Setting("혈액번호", lngX2, 12, 30, "C", "C", False)
    Call Print_Setting("혈액종류", lngX2 + 30, 12, 15, "C", "C", False)
    Call Print_Setting("혈액형", lngX2 + 45, 12, 15, "C", "C", False)
    Call Print_Setting("채혈일", lngX2 + 60, 12, 15, "L", "C", False)
    Call Print_Setting("출고자", lngX2 + 75, 12, 27, "L", "C", False)
    Call Print_Setting("수혈시간", lngX2 + 90, 12, 20, "L", "C", False)
    Call Print_Setting("수혈끝", lngX2 + 105, 12, 20, "L", "C", False)
    Call Print_Setting("Dr.", lngX2 + 120, 12, 10, "C", "C", False)

    Call Print_Setting("Nr.", lngX2 + 130, 12, 10, "C", "C", False)
    Call Print_Setting("수혈부작용", lngX2 + 140, 12, 20, "C", "C", False)

    lngCurYPos = lngCurYPos + 8 * 12
    Printer.FontBold = True
    Call Print_Setting("Memo (Special v/s 및 환자상태기록)", PrtLeft, LineSpace, , , "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos + 50), , B
    
    
    Printer.Line (PrtLeft, lngCurYPos + 55)-(Twidth - PrtLeft, lngCurYPos + 55)
    
    lngCurYPos = lngCurYPos + 60
    
    Call Print_Setting(HOSPITAL_NAME, PrtLeft, LineSpace, Twidth, "C", "C", False)
    Printer.FontBold = False
    
    Printer.EndDoc
    
End Sub
Private Sub GetBatchABO()
    Dim ObjABO      As New clsABO
    Dim objPrgBar   As New clsProgress
    Dim QueryOrder  As New clsQueryOrder
    Dim ii          As Integer
    Dim tmpptid     As String
    Dim sPtid       As String
    Dim sORDDT      As String
    Dim sLastDt     As String
    
    With objPrgBar
        .Container = Me
        .Width = LisLabel3.Width
        .Left = LisLabel3.Left
        .Top = LisLabel3.Top
        .Height = 300
        
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = LisLabel3.Width
'        .XPos = LisLabel3.Left
'        .YPos = LisLabel3.Top
'        .YHeight = 300
'        .ForeColor = &H864B24
'        .Msg = "수혈처방내역을 검색중입니다..."
'        .value = 1
    End With
    
    With tblPtList
        objPrgBar.Max = .DataRowCnt
        .ReDraw = False
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = TblColumn.tcPTID
            If tmpptid <> Trim(.value) Then
                sPtid = Trim(.value)
                ObjABO.PtId = sPtid
                If ObjABO.GetABO = False Then
                    .Col = TblColumn.TcABO:     .value = ""
                Else
                    .Col = TblColumn.TcABO:     .value = ObjABO.ABO & ObjABO.Rh
                End If
                sLastDt = QueryOrder.GetLatestTrandDt(sPtid)
                .Col = TblColumn.tcTRANSDT:  .value = sLastDt
            Else
                .Col = TblColumn.TcABO:      .value = ObjABO.ABO & ObjABO.Rh
                .Col = TblColumn.tcTRANSDT:  .value = sLastDt
            End If
            .Col = TblColumn.tcPTID: tmpptid = Trim(.value)
            objPrgBar.value = ii: objPrgBar.Message = tmpptid & " 의 혈액형을 검색정입니다."
        Next
        .ReDraw = True
    End With
    
    Set ObjABO = Nothing
    Set QueryOrder = Nothing
    Set objPrgBar = Nothing
End Sub

