VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmBBS912 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "환자별수혈내역"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS912.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00F4F0F2&
      Caption         =   "전체목록"
      Height          =   325
      Left            =   9540
      Style           =   1  '그래픽
      TabIndex        =   25
      Tag             =   "15101"
      Top             =   3630
      Width           =   1125
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   480
      Left            =   3360
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "124"
      Top             =   2880
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   9480
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "128"
      Top             =   8040
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   480
      Left            =   8100
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "15101"
      Top             =   8040
      Width           =   1245
   End
   Begin VB.TextBox txtRmk 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   240
      MaxLength       =   10
      TabIndex        =   14
      Top             =   6900
      Width           =   9465
   End
   Begin VB.TextBox txtPtid 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   1
      Top             =   660
      Width           =   1545
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   660
      Width           =   300
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   661
      BackColor       =   8421504
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
      Caption         =   "환 자 정 보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel10 
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   180
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   661
      BackColor       =   8421504
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
      Caption         =   "처 방 정 보"
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvwOrder 
      Height          =   2895
      Left            =   4860
      TabIndex        =   11
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "처방명 "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "수량"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "수혈사유"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "수혈예정일"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "변수집합"
         Object.Width           =   0
      EndProperty
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   661
      BackColor       =   8421504
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
      Caption         =   "수 혈 정 보"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   2355
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   10335
      _Version        =   196608
      _ExtentX        =   18230
      _ExtentY        =   4154
      _StockProps     =   64
      BackColorStyle  =   1
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
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   9
      MaxRows         =   5
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS912.frx":076A
      TextTip         =   4
   End
   Begin MedControls1.LisLabel lblPtNm 
      Height          =   315
      Left            =   1140
      TabIndex        =   19
      Top             =   1020
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MedControls1.LisLabel lblSexAge 
      Height          =   315
      Left            =   1140
      TabIndex        =   20
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MedControls1.LisLabel lblWard 
      Height          =   315
      Left            =   1140
      TabIndex        =   21
      Top             =   1740
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MedControls1.LisLabel lblDept 
      Height          =   315
      Left            =   1140
      TabIndex        =   22
      Top             =   2100
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MedControls1.LisLabel lbldisCd 
      Height          =   315
      Left            =   1140
      TabIndex        =   23
      Top             =   2460
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MedControls1.LisLabel lblDisNm 
      Height          =   315
      Left            =   2340
      TabIndex        =   24
      Top             =   2460
      Width           =   2235
      _ExtentX        =   3942
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
   Begin MedControls1.LisLabel lblReaction 
      Height          =   315
      Left            =   1560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Reaction"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblInfection 
      Height          =   315
      Left            =   1140
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "@"
      Appearance      =   0
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "부 작 용 :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   240
      TabIndex        =   15
      Top             =   6600
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3915
      Index           =   2
      Left            =   135
      Top             =   3960
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3015
      Index           =   1
      Left            =   4815
      Top             =   540
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3015
      Index           =   0
      Left            =   135
      Top             =   540
      Width           =   4635
   End
   Begin VB.Label lblABO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "AB+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1140
      Left            =   2820
      TabIndex        =   8
      Top             =   1020
      Width           =   1920
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "환자명:"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   7
      Tag             =   "40304"
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "환자ID:"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Tag             =   "103"
      Top             =   720
      Width           =   630
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "성별/나이"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   285
      TabIndex        =   5
      Tag             =   "108"
      Top             =   1425
      Width           =   810
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "병동:"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   300
      TabIndex        =   4
      Tag             =   "40304"
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "진료과:"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   300
      TabIndex        =   3
      Tag             =   "103"
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lable 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "상  병"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   2
      Tag             =   "40304"
      Top             =   2520
      Width           =   540
   End
End
Attribute VB_Name = "frmBBS912"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TblColumn
    tcDELIVERY = 1
    TcBLOODNO
    TcABO
    tcCOMPONM
    tcVOLUMN
    tcTRANSDT
    tcTRANSVOL
    tcREACTDIV
    tcREACTRMK
End Enum
'-------------------------------------------------------
'프린트 출력시(처방별 수혈인지, 전체 수혈인지 구분한다.)
'True:전체수혈,False:처방별 수혈
'-------------------------------------------------------
Private blnPrint As Boolean


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim strChk As String
    
    If tblList.MaxRows = 0 Then Exit Sub
    If blnPrint = False Then
        strChk = MsgBox("처방별 수혈내역만을 출력하시겠습니까?", vbInformation + vbYesNo, "수혈내역출력")
        If strChk = vbYes Then
            Call Print_Go
        End If
    Else
        Call Print_Go
    End If
End Sub

Private Sub cmdQuery_Click()
    If txtPtid <> "" Then
        lvwOrder.ListItems.Clear
        tblList.MaxRows = 0
        Call Query
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim objstatic As New clsStatics
    Dim objPop    As New S2DLP.clsS2DLP
    Dim strTmp      As String
    
    Call objPop.ListPop(objstatic.Trans_PtRecordSet, Me.Top + cmdSearch.Top, _
                                                     Me.Left + cmdSearch.Left + cmdSearch.Width)
    
    strTmp = objPop.SelectedString
    If strTmp <> "" Then
        txtPtid = medGetP(objPop.SelectedString, 1, ";")
        lblPtNm.Caption = medGetP(objPop.SelectedString, 2, ";")
        Call Query_Pt(txtPtid)
    End If
    
    Set objstatic = Nothing
    Set objPop = Nothing
End Sub

Private Sub cmdTotal_Click()
'------------------------------
'환자에 대한 전체 수혈내역 조회
'------------------------------
    Dim objTrans As clsStatics
    Dim objPross As clsProgressBar
    Dim Rs       As DrRecordSet
    Dim BldNO    As String
    Dim ii       As Integer
    
    If lblPtNm.Caption = "" Then Exit Sub
    
    Set objTrans = New clsStatics
    
'    objTrans.setDbConn DBConn
    
    Set Rs = objTrans.BLood_TotalDeliveryRecord(txtPtid)
    
    If Not Rs.EOF Then
        Set objPross = New clsProgressBar
        Set objPross.StatusBar = StatusBar
        objPross.Max = Rs.RecordCount
        With tblList
            .MaxRows = Rs.RecordCount
            Do Until Rs.EOF
                ii = ii + 1
                .Row = ii
                .Col = TblColumn.tcDELIVERY: .Value = Format(Rs.Fields("deliverydt").Value & "", "####-##-##")
                BldNO = Rs.Fields("bldsrc").Value & "-" & Rs.Fields("bldyy").Value & "" & "-" & Format(Rs.Fields("bldno").Value & "", "000000")
                .Col = TblColumn.TcBLOODNO:  .Value = BldNO
                .Col = TblColumn.TcABO:      .Value = Rs.Fields("abo").Value & "" & Rs.Fields("rh").Value & ""
                .Col = TblColumn.tcCOMPONM:  .Value = Rs.Fields("componm").Value & ""
                .Col = TblColumn.tcVOLUMN:   .Value = Rs.Fields("volumn").Value & ""
                .Col = TblColumn.tcTRANSDT:  .Value = Format(Rs.Fields("transdt").Value & "", "####-##-##")
                .Col = TblColumn.tcTRANSVOL: .Value = Rs.Fields("transvol").Value & ""
                .Col = TblColumn.tcREACTDIV: .Value = IIf(Rs.Fields("reactdiv").Value = "1", "Y", "")
                .Col = TblColumn.tcREACTRMK: .Value = Rs.Fields("reactrmk").Value & ""
                objPross.Value = ii
                Rs.MoveNext
            Loop
        End With
        blnPrint = True
    End If
    lvwOrder.ListItems.Clear
    Call Query
    
    Set Rs = Nothing
    Set objTrans = Nothing
    Set objPross = Nothing
End Sub

Private Sub Form_Load()
    Call Clear
End Sub
Private Sub Clear()
    
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblWard.Caption = ""
    lblDept.Caption = ""
    lbldisCd.Caption = ""
    lblDisNm.Caption = ""
    lblInfection.Visible = False
    
    lblReaction.Visible = False
    lblABO.Caption = ""
    
    lvwOrder.ListItems.Clear
    tblList.MaxRows = 0
    
End Sub
Private Sub Query()
    Dim objstatic As clsStatics
    Dim objReason As clsQueryOrder
    Dim objGorddt As clsDictionary
    
    Dim Rs        As DrRecordSet
    Dim itmx      As ListItem
    Dim strTmp    As String
    
    Dim accdt     As String
    Dim accseq    As String
    Dim testnm    As String
    Dim unitqty   As String
    Dim reason    As String
    Dim reqdt     As String
    Dim orddt     As String
    Dim ordno     As String
    Dim ordseq    As String
    Dim ii        As Integer
    
    Set objstatic = New clsStatics
'    objstatic.setDbConn DBConn
    
    Set Rs = objstatic.Pt_TransBloodRecord(txtPtid)
    
    If Not Rs.EOF Then
        Set objReason = New clsQueryOrder
        Set objGorddt = New clsDictionary
        
        objGorddt.Clear
        objGorddt.FieldInialize "seq", "orddt"
        
        Do Until Rs.EOF
            ii = ii + 1
            accdt = Rs.Fields("accdt").Value
            accseq = Rs.Fields("accseq").Value
    
            strTmp = objstatic.Pt_OrderRecordset(accdt, accseq)
        
            If strTmp <> "" Then
                testnm = medGetP(strTmp, 1, COL_DIV)
                unitqty = medGetP(strTmp, 2, COL_DIV)
                reqdt = Format(medGetP(strTmp, 3, COL_DIV), "####-##-##")
                orddt = medGetP(strTmp, 4, COL_DIV)
                ordno = medGetP(strTmp, 5, COL_DIV)
                ordseq = medGetP(strTmp, 6, COL_DIV)
                reason = objReason.GetTransReason(txtPtid, orddt, ordno)
                Set itmx = lvwOrder.ListItems.Add(, , testnm)
                    itmx.SubItems(1) = medGetP(strTmp, 2, COL_DIV)
                    itmx.SubItems(2) = reason
                    itmx.SubItems(3) = reqdt
                    itmx.SubItems(4) = accdt & COL_DIV & accseq
                objGorddt.AddNew ii, orddt
            End If
            Rs.MoveNext
        Loop
        Call Detail_Ptinfo(objGorddt)
    End If
    Set Rs = Nothing
    Set objstatic = Nothing
End Sub
Private Sub Detail_Ptinfo(ByVal objDic As clsDictionary)
'------------------------------------------------
'혈액형,부작용,감염정보,상병코드,상병을 조회한다.
'------------------------------------------------
    Dim objABO       As New clsABO
    Dim objDisease   As New clsDisease
    Dim objinfection As New clsInfection
    Dim objReaction  As New clsReaction
    Dim Rs           As New DrRecordSet
    
    With objABO
        .ptid = txtPtid
        If .GetABO = True Then
            lblABO.Caption = .ABO & .Rh
        Else
            lblABO.Caption = ""
        End If
    End With
    
    With objinfection
        .ptid = txtPtid
        .GetInfection
        If .Infection = True Then
            lblInfection.Visible = True
        Else
            lblInfection.Visible = False
        End If
    End With
    
    With objReaction
        .ptid = txtPtid
        If .GetReaction = True Then
            lblReaction.Visible = .Reaction
        Else
            lblReaction.Visible = False
        End If
    End With
    
    If objDic.RecordCount > 0 Then
        objDic.MoveFirst
        With objDisease
            .ptid = txtPtid
            Do Until objDic.EOF
                .orddt = objDic.Fields("orddt")
                If .GetDisease = True Then
                    lbldisCd.Caption = .DiseaseCd
                    lblDisNm.Caption = .DiseaseNm
                    Exit Do
                Else
                    lbldisCd.Caption = "": lblDisNm.Caption = ""
                End If
                objDic.MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
    Set objABO = Nothing
    Set objDisease = Nothing
    Set objReaction = Nothing
    Set objinfection = Nothing
End Sub
Private Sub lvwOrder_Click()
    If lvwOrder.ListItems.Count = 0 Then Exit Sub
    Call Query_Delivery
    blnPrint = False
End Sub
Private Sub Query_Delivery()
    Dim objStat As New clsStatics
    Dim objBar  As clsProgressBar
    Dim Rs      As DrRecordSet
    Dim itmx    As ListItem
    Dim strTmp  As String
    Dim accdt   As String
    Dim accseq  As String
    Dim BldNO   As String
    Dim Compocd As String
    
    Dim ii      As Integer
    
    
'    objStat.setDbConn DBConn
    
    
    strTmp = lvwOrder.SelectedItem.SubItems(4)
    
    accdt = medGetP(strTmp, 1, COL_DIV)
    accseq = medGetP(strTmp, 2, COL_DIV)
    
    Set Rs = objStat.BLood_DeliveryRecord(accdt, accseq)
    If Not Rs.EOF Then
        Set objBar = New clsProgressBar
        Set objBar.StatusBar = StatusBar
        objBar.Max = Rs.RecordCount
        
        With tblList
            .MaxRows = Rs.RecordCount
            Do Until Rs.EOF
                ii = ii + 1
                .Row = ii
                .Col = TblColumn.tcDELIVERY: .Value = Format(Rs.Fields("deliverydt").Value, "####-##-##")
                BldNO = Rs.Fields("bldsrc").Value & "-" & Rs.Fields("bldyy").Value & "-" & Format(Rs.Fields("bldno").Value, "000000")
                .Col = TblColumn.TcBLOODNO:  .Value = BldNO
                .Col = TblColumn.TcABO:      .Value = Rs.Fields("abo").Value & Rs.Fields("rh").Value
                .Col = TblColumn.tcCOMPONM:  .Value = Rs.Fields("componm").Value
                .Col = TblColumn.tcVOLUMN:   .Value = Rs.Fields("volumn").Value
                .Col = TblColumn.tcTRANSDT:  .Value = Format(Rs.Fields("transdt").Value, "####-##-##") & " " & Format(Mid(Rs.Fields("transtm").Value, 1, 2), "00:00")
                .Col = TblColumn.tcTRANSVOL: .Value = Rs.Fields("transvol").Value & ""
                .Col = TblColumn.tcREACTDIV: .Value = IIf(Rs.Fields("reactdiv").Value = "1", "Y", "")
                .Col = TblColumn.tcREACTRMK: .Value = Rs.Fields("reactrmk").Value & ""
                objBar.Value = ii
                Rs.MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
    Set objBar = Nothing
    Set objStat = Nothing
End Sub


Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With tblList
        .Row = Row
        .Col = TblColumn.tcREACTRMK
        txtRmk = .Value
    End With
End Sub
Private Sub Query_Pt(ByVal ptid As String)
    Dim objstatic  As New clsStatics
    Dim strTmp     As String
    
'    objstatic.setDbConn DBConn
    strTmp = objstatic.Trans_PtNm(ptid)
    
    If strTmp <> "" Then
        lblPtNm.Caption = medGetP(strTmp, 1, COL_DIV)
        Call SexCheck(medGetP(strTmp, 2, COL_DIV))
        lblWard.Caption = medGetP(strTmp, 3, COL_DIV)
        lblDept.Caption = medGetP(strTmp, 4, COL_DIV)
    Else
        MsgBox "조건에 해당하는 환자가 없습니다.", vbInformation + vbOKOnly, "조회대상자 찾기"
        txtPtid = ""
    End If
    Set objstatic = Nothing
End Sub

Private Sub txtPtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtPtid = "" Then Exit Sub
        Call Query_Pt(txtPtid)
    End If
End Sub
Private Sub SexCheck(ByVal SSN As String)
    Dim strTmp As String
    Dim sex    As String
    Dim age    As String
    Dim lngsex As Long
    
    strTmp = Mid(SSN, 3, 6) & "-" & Mid(SSN, 9)
    
    If strTmp <> "" Then
        lngsex = Val(Mid(medGetP(strTmp, 2, "-"), 1, 1))
        If lngsex = 1 Or lngsex = 3 Then
            sex = "남"
        ElseIf lngsex = 2 Or lngsex = 4 Then
            sex = "여"
        Else
            sex = ""
        End If
    Else
        sex = ""
    End If
    
    If Len(SSN) = 15 Then
        age = medFindAge(Mid(SSN, 1, 8), "Y")
    End If
    lblSexAge.Caption = sex & "/" & age
    
End Sub

Private Sub Print_Go()
    Dim strTmp As String
    Dim intFNum As Integer
    Dim strRfile As String
    Dim strRptPath As String
    Dim ii As Integer
    Dim jj As Integer
    
    With tblList
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = TblColumn.tcDELIVERY To TblColumn.tcREACTDIV
                .Col = jj
                strTmp = strTmp & .Value & vbTab
            Next jj
            strTmp = strTmp & vbCr
        Next ii
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End With
    
    strRfile = BBSRPTPATH & "\CrystalReport.txt"
    strRptPath = BBSRPTPATH & "\frmBBS912.rpt"
    intFNum = FreeFile
    
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    With CReport
        .ParameterFields(0) = "ptid;" & txtPtid.Text & ";TRUE"
        .ParameterFields(1) = "ptnm;" & lblPtNm.Caption & ";TRUE"
        .ParameterFields(2) = "dept;" & lblDept.Caption & ";TRUE"
        .ParameterFields(3) = "ward;" & lblWard.Caption & ";TRUE"
        .ParameterFields(4) = "abo;" & lblABO.Caption & ";TRUE"
        .ParameterFields(5) = "sicknm;" & lblDisNm.Caption & ";TRUE"
        .ParameterFields(6) = "sexage;" & lblSexAge.Caption & ";TRUE"
        .ParameterFields(7) = "hostnm;" & HOSPITAL_NAME & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With
End Sub
