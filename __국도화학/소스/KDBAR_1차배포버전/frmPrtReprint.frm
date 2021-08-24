VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPrtReprint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "라벨 재출력"
   ClientHeight    =   9930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   19845
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   8205
      Left            =   90
      TabIndex        =   13
      Top             =   1530
      Width           =   20000
      Begin FPSpread.vaSpread spdPrtReprint 
         Height          =   7665
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   18975
         _Version        =   393216
         _ExtentX        =   33470
         _ExtentY        =   13520
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   25
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmPrtReprint.frx":0000
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   20000
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   17970
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtPToNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7470
         MaxLength       =   5
         TabIndex        =   6
         Top             =   840
         Width           =   1080
      End
      Begin VB.TextBox txtPFrNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         MaxLength       =   5
         TabIndex        =   5
         Top             =   840
         Width           =   1080
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Caption         =   "출력"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   16830
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   780
         Width           =   1095
      End
      Begin VB.ComboBox cboLabelType 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmPrtReprint.frx":13FC
         Left            =   10590
         List            =   "frmPrtReprint.frx":13FE
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   330
         Width           =   2595
      End
      Begin VB.ComboBox cboSlittingNo 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmPrtReprint.frx":1400
         Left            =   2460
         List            =   "frmPrtReprint.frx":1402
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   840
         Width           =   1845
      End
      Begin VB.ComboBox cboProd 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   330
         Width           =   3105
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   14550
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   780
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   15690
         Style           =   1  '그래픽
         TabIndex        =   8
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   780
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   1590
         TabIndex        =   1
         Top             =   330
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131006464
         CurrentDate     =   43884
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   17790
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   18450
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":1404
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":199E
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":1F38
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":24D2
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":2D64
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":2EBE
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":3018
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":3172
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReprint.frx":3A4C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7140
         TabIndex        =   19
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "▶ P No"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4770
         TabIndex        =   18
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "▶ Slitting 작업번호"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   330
         TabIndex        =   17
         Top             =   900
         Width           =   2025
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "▶ 구분"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9570
         TabIndex        =   16
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "▶ 제품명"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4770
         TabIndex        =   15
         Top             =   390
         Width           =   1065
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   14940
         Picture         =   "frmPrtReprint.frx":4326
         Top             =   330
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         BackStyle       =   0  '투명
         Caption         =   "Com1 연결성공"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15360
         TabIndex        =   14
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 생산일자 "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   330
         TabIndex        =   12
         Top             =   390
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmPrtReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmPrtReprint.frm
'   작성자  : 오세원
'   내  용  : 라벨 재출력
'   작성일  : 2020-03-03
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cboProd_Click()
    Dim strCompCd    As String
    
'    txtProdCd.Text = Trim(mGetP(cboProd.Text, 2, "|"))
'    strCompCd = Trim(mGetP(cboComp.Text, 2, "|"))
    
'    Call GetComp_CodeName(txtProd.Text)
    
'    spdRegOrderDetail.MaxRows = 0
    
End Sub

'
'Private Sub cmdAdd_Click()
'    Dim pAdoRS      As ADODB.Recordset
'    Dim intRow      As Integer
'    Dim intNum      As Integer
'    Dim intMaxNum   As Integer
'
'    With spdRegOrderDetail
'        .MaxRows = .MaxRows + 1
'        Call SetText(spdRegOrderDetail, dtpProdOrderDt.Value, .MaxRows, 1)
'        Call SetText(spdRegOrderDetail, cboSlittingNo.Text, .MaxRows, 2)
'        Call SetText(spdRegOrderDetail, txtProdCd.Text, .MaxRows, 3)
'        Call SetText(spdRegOrderDetail, cboSlittingNo.Text, .MaxRows, 4)
'        Call SetText(spdRegOrderDetail, CStr(.MaxRows), .MaxRows, 5)
'
'        Call SetText(spdRegOrderDetail, "", .MaxRows, 6)
'        .Row = .MaxRows
'        .Col = 6
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 300
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'
''        Call SetText(spdRegOrderDetail, "P" & CStr(.MaxRows), .MaxRows, 7)
''        .Row = .MaxRows
''        .Col = 7
''        .CellType = CellTypeEdit
''        .TypeMaxEditLen = 2
''        .TypeHAlign = TypeHAlignCenter
''        .TypeVAlign = TypeVAlignCenter
'
'        Call SetText(spdRegOrderDetail, "", .MaxRows, 7)
'        .Row = .MaxRows
'        .Col = 7
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 4
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'
'        Call SetText(spdRegOrderDetail, "", .MaxRows, 8)
'        .Row = .MaxRows
'        .Col = 8
'        .CellType = CellTypeEdit
'        .TypeMaxEditLen = 4
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'    End With
'
'
'
'
'End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReprint.MaxRows = 0

    dtpDate.Value = Now

    cboSlittingNo.Clear
    For i = 1 To 10
        cboSlittingNo.AddItem CStr(i)
    Next
    cboSlittingNo.ListIndex = 0
    
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub



'제품 리스트 가져오기(조회용)
Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset

    Set pAdoRS = New ADODB.Recordset

    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)

    cboProd.Clear

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboProd.AddItem "전체" & Space(50) & "|전체"

        Do Until pAdoRS.EOF
            cboProd.AddItem pAdoRS.Fields("PROD_PRT_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop

        If pAdoRS.RecordCount > 0 Then
            cboProd.ListIndex = 0
        End If
    End If

    pAdoRS.Close

End Sub
    
    
''제품 리스트 가져오기(등록용)
'Private Sub GetProdList_CodeName_Reg(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
'
'    cboProdCd.Clear
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until pAdoRS.EOF
'            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
'            pAdoRS.MoveNext
'        Loop
'
'        If pAdoRS.RecordCount > 0 Then
'            cboProdCd.ListIndex = 0
'        End If
'    End If
'
'    pAdoRS.Close
'
'End Sub
    
'' 작업지시서 리스트 가져옴
Private Sub GetPrintList(ByVal pOrderDate As String, ByVal pProdCd As String, ByVal pLabelType As String, ByVal pSltNo As String, ByVal pPFrNo As String, ByVal pPToNo As String)

    Dim strLabelType    As String

    Set AdoRs = Get_PrintList(pOrderDate, pProdCd, pLabelType, pSltNo, pPFrNo, pPToNo)

    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReprint
                .MaxRows = .MaxRows + 1
                Call SetText(spdPrtReprint, "1", .MaxRows, 1)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_LOT_NO").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReprint, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 3)
                'Call SetText(spdPrtReprint, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 4)
                'Call SetText(spdPrtReprint, AdoRs.Fields("COMP_NAME").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReprint, AdoRs.Fields("PROD_REEL_BAR").Value & "", .MaxRows, 6)
                'Call SetText(spdPrtReprint, AdoRs.Fields("PNO").Value & "", .MaxRows, 7)
                'Call SetText(spdPrtReprint, AdoRs.Fields("USER_NAME").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReprint, AdoRs.Fields("REGIST_DT_R").Value & "", .MaxRows, 9)
            End With
            AdoRs.MoveNext
        Loop
    End If
    AdoRs.Close

End Sub


Private Sub cmdSearch_Click()
    Dim strDate         As String
    Dim strProdCd       As String
    Dim strLabelType    As String
    Dim strSltNo        As String
    Dim strFrNo         As String
    Dim strToNo         As String
    
    strDate = Format(dtpDate.Value, "yyyymmdd")
    strProdCd = mGetP(cboProd.Text, 2, "|")
    strLabelType = Mid(cboLabelType, 1, 1)
    strSltNo = cboSlittingNo.Text
    strFrNo = txtPFrNo.Text
    strToNo = txtPToNo.Text
    
    ' 포장코드 리스트 가져오기
    Call GetPrintList(strDate, strProdCd, strLabelType, strSltNo, strFrNo, strToNo)

    
End Sub


Private Sub Form_Load()

    Call CtlInitializing
    
    '-- 통신열기
    Call OpenCommunication
    
    '고객사 리스트 가져오기
'    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    ' 포장코드 리스트 가져오기
'    Call GetPackList
    
    
End Sub

Private Sub OpenCommunication()

On Error GoTo ErrHandle
    
'    If frmPrtReel.comEqp.PortOpen = True Then
'        frmPrtReel.comEqp.PortOpen = False
'    End If
'    If frmPrtPP.comEqp.PortOpen = True Then
'        frmPrtPP.comEqp.PortOpen = False
'    End If
'    If frmPrtICE.comEqp.PortOpen = True Then
'        frmPrtICE.comEqp.PortOpen = False
'    End If

    comEqp.CommPort = gComm.COMPORT
    comEqp.RTSEnable = gComm.RTSEnable
    comEqp.DTREnable = gComm.DTREnable
    comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

    If comEqp.PortOpen Then
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
    End If

    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            
            Resume Next
        Else
            
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "위    치 : " & "Public Sub OpenCommunication()" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If


End Sub

'Private Sub GetPackList()
'    Dim pAdoRS      As ADODB.Recordset
'    Dim strPackInfo As String
'
'    Set pAdoRS = Get_PackList
'
'    cboPackCd.Clear
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until pAdoRS.EOF
'            ' PACK_CAT_WIDTH,PACK_PRO_WIDTH,PACK_PRO_LENGTH
'            strPackInfo = pAdoRS.Fields("PACK_CORE").Value & "x" & pAdoRS.Fields("PACK_DIA").Value & " " & pAdoRS.Fields("PACK_CAT_WIDTH").Value & " " & pAdoRS.Fields("PACK_PRO_WIDTH").Value
'
'            cboPackCd.AddItem pAdoRS.Fields("PACK_NAME").Value & Space(3) & strPackInfo & Space(20) & "|" & pAdoRS.Fields("PACK_CD").Value & Space(3)
'            pAdoRS.MoveNext
'        Loop
'
'    End If
'
'    If pAdoRS.RecordCount > 0 Then
'        cboPackCd.ListIndex = 0
'    End If
'
'    pAdoRS.Close
'
'End Sub



Private Function GetCompList_Name(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_Name(pCompCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            GetCompList_Name = pAdoRS.Fields("COMP_NAME").Value & ""

            pAdoRS.MoveNext
        Loop

    End If

    pAdoRS.Close

End Function

'-- 상단 고객사리스트 가져오기
'Private Sub GetCompList_CodeName()
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_CompList_CodeName
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        cboCompCd.Clear
'
'        Do Until pAdoRS.EOF
'            cboCompCd.AddItem pAdoRS.Fields("COMP_VIEW").Value & Space(1) & ":" & Space(1) & pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
'
'            pAdoRS.MoveNext
'        Loop
'
'        If pAdoRS.RecordCount > 0 Then
'            cboCompCd.ListIndex = 0
'        End If
'    End If
'
'    pAdoRS.Close
'
'End Sub

'-- 제품선택했을때 해당 고객사 가져오기
'Private Sub GetComp_CodeName(ByVal pProdCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'    Dim i           As Integer
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_Comp_CodeName(pProdCd)
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        txtProdLen.Text = ""
'
'        Do Until pAdoRS.EOF
'            txtProdLen.Text = pAdoRS.Fields("PROD_LENGTH").Value & ""
'            For i = 0 To cboCompCd.ListCount
'                If pAdoRS.Fields("COMP_CD").Value & "" = mGetP(cboCompCd.List(i), 2, "|") Then
'                    cboCompCd.ListIndex = i
'                    Exit For
'                End If
'            Next
'            pAdoRS.MoveNext
'        Loop
'
'        pAdoRS.Close
'
'    End If
'
'
'End Sub


'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim i           As Integer
    
    With spdPrtReprint
        Call SetText(spdPrtReprint, "선택", 0, 1):              .ColWidth(1) = 6
        Call SetText(spdPrtReprint, "생산LotNo", 0, 2):         .ColWidth(2) = 20
        Call SetText(spdPrtReprint, "생산일자", 0, 3):          .ColWidth(3) = 15
        Call SetText(spdPrtReprint, "제품명", 0, 4):            .ColWidth(4) = 20
        Call SetText(spdPrtReprint, "고객사", 0, 5):            .ColWidth(5) = 11
        Call SetText(spdPrtReprint, "Reel바코드", 0, 6):        .ColWidth(6) = 30
        Call SetText(spdPrtReprint, "P No", 0, 7):              .ColWidth(7) = 10
        Call SetText(spdPrtReprint, "출력자", 0, 8):            .ColWidth(8) = 10
        Call SetText(spdPrtReprint, "출력시간", 0, 9):          .ColWidth(9) = 25
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 0
        .ColWidth(23) = 0
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .MaxRows = 0
    End With
    
    dtpDate.Value = Now
    cboSlittingNo.Clear
    For i = 1 To 10
        cboSlittingNo.AddItem CStr(i)
    Next
    cboSlittingNo.ListIndex = 0
    
    cboLabelType.Clear
    cboLabelType.AddItem "Reel"
    cboLabelType.AddItem "PP Box"
    cboLabelType.AddItem "ICE Box"
    cboLabelType.ListIndex = 0
    
    txtPFrNo.Text = "P101"
    txtPToNo.Text = "P102"
    
    gSORT = 0

End Sub

'Private Sub spdRegOrder_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim i               As Integer
'    Dim strDate         As String
'    Dim strLotNo        As String
'    Dim strProdPosNo    As String
'    Dim strProdCd       As String
'    Dim strSltNo        As String
'
'    If Row = 0 Then
'        Call SetSpreadSort(spdRegOrder)
'        Exit Sub
'    End If
'
'    spdRegOrderDetail.MaxRows = 0
'
'    txtLotNo.Text = GetText(spdRegOrder, Row, 1)
'    strDate = GetText(spdRegOrder, Row, 2)
'    dtpProdOrderDt.Value = Format(strDate, "####-##-##")
'    'cboProdPosNo.Text = GetText(spdRegOrder, Row, 3)
''    strProdPosNo = cboProdPosNo.Text
'
'    For i = 0 To cboCompCd.ListCount
'        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegOrder, Row, 4) Then
'            cboProdCd.ListIndex = i
'            strProdCd = Trim(mGetP(cboProdCd.List(i), 2, "|"))
'            Exit For
'        End If
'    Next
'    For i = 0 To cboPackCd.ListCount
'        If Mid(cboPackCd.List(i), 1, 2) = GetText(spdRegOrder, Row, 6) Then
'            cboPackCd.ListIndex = i
'            Exit For
'        End If
'    Next
'    txtOrderMemo.Text = GetText(spdRegOrder, Row, 7)
'    txtProdLen.Text = GetText(spdRegOrder, Row, 9)
'    cboSlittingNo.Text = GetText(spdRegOrder, Row, 10)
'    strSltNo = cboSlittingNo.Text
'    txtReelQTY.Text = GetText(spdRegOrder, Row, 11)
'
'    '고객사
'    For i = 0 To cboCompCd.ListCount
'        If Trim(mGetP(cboCompCd.List(i), 2, "|")) = Trim(mGetP(GetText(spdRegOrder, Row, 12), 2, "|")) Then
'            cboCompCd.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    Call GetOrderDetail(strDate, strProdCd, strSltNo)
'
'
'End Sub


' 작업지시서 리스트 가져옴 'strDate, cboProdPosNo.Text, cboProdCd.Text, cboSlittingNo.Text
'Private Sub GetOrderDetail(ByVal pDate As String, ByVal pProCd As String, ByVal pSltNo As String)
'
'    Set AdoRs = Get_OrderDetail(pDate, pProCd, pSltNo)
'
'    If AdoRs Is Nothing Then
'        '등록된 정보 없음
'    Else
'        spdRegOrderDetail.MaxRows = 0
'
'        Do Until AdoRs.EOF
'            With spdRegOrderDetail
'                .MaxRows = .MaxRows + 1
'
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
'                'Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 2)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 3)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 4)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("SEQ_NO").Value & "", .MaxRows, 5)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_INFO").Value & "", .MaxRows, 6)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_F").Value & "", .MaxRows, 7)
'                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_T").Value & "", .MaxRows, 8)
'
'            End With
'
'            AdoRs.MoveNext
'        Loop
'        AdoRs.Close
'    End If
'
'End Sub

