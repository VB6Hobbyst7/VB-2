VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmResultReview 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "결과 조회"
   ClientHeight    =   9075
   ClientLeft      =   5055
   ClientTop       =   330
   ClientWidth     =   10080
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResultReview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Tag             =   "ResultView2"
   Begin VB.Frame fraStatus 
      BackColor       =   &H00F7FFFF&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1320
      Left            =   2025
      TabIndex        =   16
      Top             =   3585
      Visible         =   0   'False
      Width           =   6585
      Begin MSComctlLib.ProgressBar barStatus 
         Height          =   195
         Left            =   1410
         TabIndex        =   17
         Top             =   390
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   390
         TabIndex        =   18
         Top             =   825
         Width           =   6045
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   495
         Picture         =   "frmResultReview.frx":038A
         Top             =   255
         Width           =   480
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   45
         Top             =   45
         Width           =   6465
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FCEFE9&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   10080
      TabIndex        =   2
      Top             =   0
      Width           =   10080
      Begin VB.CommandButton cmdExit 
         Caption         =   "종료 (&X)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8925
         TabIndex        =   30
         Tag             =   "128"
         Top             =   315
         Width           =   1110
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   975
         MaxLength       =   10
         TabIndex        =   0
         Top             =   90
         Width           =   1545
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   3855
         TabIndex        =   7
         Top             =   75
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         BackColor       =   16703181
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   300
         Left            =   6825
         TabIndex        =   12
         Top             =   405
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         BackColor       =   16703181
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   300
         Left            =   6825
         TabIndex        =   13
         Top             =   75
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         BackColor       =   16703181
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
      End
      Begin VB.Label lblBedDt 
         Height          =   240
         Left            =   885
         TabIndex        =   34
         Top             =   510
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblLocation1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "병    실 : "
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
         Left            =   5850
         TabIndex        =   15
         Tag             =   "102"
         Top             =   465
         Width           =   990
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "진 료 과 : "
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
         Left            =   5850
         TabIndex        =   14
         Tag             =   "40304"
         Top             =   150
         Width           =   990
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5415
         TabIndex        =   10
         Top             =   480
         Width           =   60
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   4845
         TabIndex        =   9
         Top             =   465
         Width           =   345
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4005
         TabIndex        =   8
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblPtId 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자 ID : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Tag             =   "105"
         Top             =   150
         Width           =   900
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성     명 : "
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
         Left            =   2790
         TabIndex        =   5
         Tag             =   "103"
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label lblSexAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성별/나이 : "
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
         Left            =   2790
         TabIndex        =   4
         Tag             =   "108"
         Top             =   465
         Width           =   1080
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FEDECD&
         Caption         =   "             /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3855
         TabIndex        =   11
         Top             =   405
         Width           =   1845
      End
   End
   Begin VB.PictureBox picOrder 
      Align           =   2  '아래 맞춤
      BackColor       =   &H00E0E0E0&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   10020
      TabIndex        =   3
      Top             =   780
      Width           =   10080
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00DBF2FD&
         Caption         =   "모든결과"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7170
         Style           =   1  '그래픽
         TabIndex        =   31
         Tag             =   "0"
         Top             =   30
         Width           =   1410
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00DBF2FD&
         Caption         =   "결과보고"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8595
         Style           =   1  '그래픽
         TabIndex        =   29
         Tag             =   "40102"
         Top             =   30
         Width           =   1410
      End
      Begin DRcontrol1.DrFrame fraRstText 
         Height          =   6615
         Left            =   0
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11668
         Title           =   "소견/감수성결과"
         TitlePos        =   1
         BackColor       =   14739427
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmdTxtClose 
            BackColor       =   &H00BECDC5&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   9345
            Style           =   1  '그래픽
            TabIndex        =   27
            Top             =   90
            Width           =   225
         End
         Begin RichTextLib.RichTextBox txtRstText 
            Height          =   6060
            Left            =   60
            TabIndex        =   33
            Top             =   435
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   10689
            _Version        =   393217
            BackColor       =   16054772
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmResultReview.frx":07CC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   7410
         Left            =   30
         TabIndex        =   1
         Top             =   705
         Width           =   9960
         _Version        =   196608
         _ExtentX        =   17568
         _ExtentY        =   13070
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowVert    =   0   'False
         MaxCols         =   12
         MaxRows         =   30
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   16252927
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmResultReview.frx":0871
         TextTip         =   4
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   315
         Left            =   6315
         TabIndex        =   19
         Top             =   405
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "검체"
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   75
         TabIndex        =   20
         Top             =   405
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "검  사  명"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   2865
         TabIndex        =   21
         Top             =   405
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "검사결과"
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   4230
         TabIndex        =   22
         Top             =   405
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "단위"
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   5370
         TabIndex        =   23
         Top             =   405
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "HL"
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   315
         Left            =   5805
         TabIndex        =   24
         Top             =   405
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "DP"
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   315
         Left            =   9255
         TabIndex        =   25
         Top             =   405
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "More"
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   315
         Left            =   7860
         TabIndex        =   28
         Top             =   405
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "보고일시"
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   7485
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   5025
         _Version        =   196608
         _ExtentX        =   8864
         _ExtentY        =   13203
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   3
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   11
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmResultReview.frx":1131
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         VisibleRows     =   22
         TextTip         =   4
      End
   End
End
Attribute VB_Name = "frmResultReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'% 폼단위 전역변수 선언

Private MyPatient As New clsLisPatient   '환자 클래스
Private MySql As New clsLISSqlStatement   'Sql문 클래스
Private ClearFg As Boolean
Private OrderFg As Boolean
Private ResultFg As Boolean
Private MsgFg As Boolean
Private OldRow As Long
Private OldBackColor As Long
Private TopLeftShow1 As Boolean
Private TopLeftShow2 As Boolean
Private aryMesg() As String

Public PtFg As Boolean
Public QueryFg As Boolean

Private StopFg As Boolean
Private m_BedinDt As String

Public Property Get BedinDt() As String
    BedinDt = m_BedinDt
End Property

Public Property Let BedinDt(ByVal vData As String)
    m_BedinDt = vData
End Property

Private Sub cmdAll_Click()

    Call StartQuery
    
    If cmdAll.Tag = "1" Then  '비정상결과만 보기
        cmdAll.Tag = "0"
        cmdAll.Caption = "모든결과"
    Else        '모두보기
        cmdAll.Tag = "1"
        cmdAll.Caption = "비정상결과"
    End If

End Sub

'%종료
Private Sub cmdExit_Click()
   Unload Me
   Set frmResultReview = Nothing
End Sub

'% 레포트 출력
Private Sub cmdReport_Click()

    With frmReport
        .ptid = txtPtId.Text
        .BedinDt = m_BedinDt
        Call .StartQuery
        .ZOrder 0
    End With

End Sub


'% 조회기간 입력 (To Date)
Private Sub StartQuery()
   
   '% 처방조회
   Dim i As Integer
   Dim ResultExist As Boolean
   
   On Error GoTo Err_Trap
   
'   cmdRefresh.Enabled = False
   
   Call TableClear
   
   'Status Bar Popup
   DoEvents
   lblStatus.Caption = lblPtNm.Caption & " 님의 검사 결과내역을 검색중입니다..."
   barStatus.Min = 0
   barStatus.Value = 0
   fraStatus.Visible = True
   fraStatus.ZOrder 0
   DoEvents
   
   With tblOrdSheet
      '.ReDraw = False
      .MaxRows = 0
         
      ResultExist = False
      ResultExist = ResultExist Or DisplayOrders("3")
      
      '.ReDraw = True
      .Col = 1: .Row = 1: .Action = ActionActiveCell
   End With
   fraStatus.Visible = False
'   cmdRefresh.Enabled = True
'   dtpFromDate.Enabled = True
'   dtpToDate.Enabled = True
   
   If Not ResultExist Then
      MsgBox "이 환자는 입력하신 기간동안에 보고된 결과가 없습니다."
'      dtpFromDate.SetFocus
      Exit Sub
   End If
   
   ClearFg = False
   ResultFg = False
   OrderFg = True
   tblOrdSheet.SetFocus
   Exit Sub
    
Err_Trap:
    Resume Next
End Sub

Private Function DisplayOrders(ByVal pTestDiv As String) As Boolean
   Dim i As Integer, j As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim strStartDt As String
   Dim tmpTestNm As String
   Dim tmpRs As New Recordset
   Dim NormalFg As Boolean
   Dim objSql As New clsLISSqlReview
   Dim strDPDiv As String
   
   'If StopFg Then Exit Function
   
   Me.Enabled = False
   
   Screen.MousePointer = vbArrowHourglass  '13
   
   SqlStmt = objSql.SqlQueryAllResults(txtPtId.Text, "orddt", m_BedinDt, _
                                      Format(Now, CS_DateDbFormat), pTestDiv)
   
    tmpRs.Open SqlStmt, DBConn
   barStatus.Max = tmpRs.RecordCount + 1
   
   DoEvents
   
   ReDim aryMesg(0)
   DisplayOrders = False
   
   With tblOrdSheet
   
      QueryFg = True
         
      .ReDraw = False
   
      While (Not tmpRs.EOF)
         If barStatus.Value < barStatus.Max Then barStatus.Value = barStatus.Value + 1
         DoEvents
      
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         
         NormalFg = True
         
         .Col = 1:  .ForeColor = DCM_LightBlue      '-- 검사명
                    tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestLongNm").Value), 1, 33)
                    If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                        Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                       .Value = tmpTestNm & " " & String(45 - Len(tmpTestNm), ".")
                    Else
                       .Value = "    " & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")         '-- 상세검사명
                    End If
         .Col = 2:  .ForeColor = DCM_Brown           '-- 결과명(코드일 경우..)
                    If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
                       .Value = "미확": .ForeColor = DCM_Gray:     .FontBold = False
                    Else
                       If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                          .TypeHAlign = TypeHAlignCenter
                          .Value = Trim("" & tmpRs.Fields("RstCd").Value)
                       Else
                          .CellType = CellTypeEdit
                          .TypeHAlign = TypeHAlignLeft
                          .Value = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
                       End If
                       If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                           .Value = "Growth"
                       ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                           .Value = Space(3)
                       End If
                    End If
         .Col = 3:  .Value = Trim("" & tmpRs.Fields("RstUnit").Value)         '-- 결과단위
         .Col = 4       '-- High / Low
                    .Value = ""
                    If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
                        If Trim("" & tmpRs.Fields("HLDiv").Value) = "H" Then .Value = "▲": .ForeColor = DCM_LightRed
                        If Trim("" & tmpRs.Fields("HLDiv").Value) = "L" Then .Value = "▼": .ForeColor = DCM_LightBlue
                        If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                        If Trim("" & tmpRs.Fields("HLDiv").Value) <> "" Then NormalFg = False
                    End If
         .Col = 5:
                    '## 1.1.44: 이상대(2005-05-23)
                    '   - Alpha결과 참고치를 "N"에서 "Abnormal"로 표시
                    strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
                    strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
                    .Value = strDPDiv: .ForeColor = vbRed                '빨간색
                    If Trim(.Value) <> "" Then NormalFg = False
         .Col = 6:  .Value = Trim("" & tmpRs.Fields("SpcNm").Value)    '검체명
         .Col = 7:  .Value = Trim("" & tmpRs.Fields("VfyDtTm").Value)         '-- 보고일시
         .Col = 8:
                    .Value = "": .ForeColor = vbBlue          '파랑색
                    If Trim("" & tmpRs.Fields("TxtFg").Value) > "0" Then .Value = "☞"
                    If Trim("" & tmpRs.Fields("TxtFg").Value) = "Y" Then .Value = "☞"
                    If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then .Value = "☞"
                    If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And Trim(tmpRs.Fields("DetailItem").Value) = "") Or _
                        Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                        If Trim("" & tmpRs.Fields("FootNoteFg").Value) = "1" Then .Value = "☞"
                        If Trim("" & tmpRs.Fields("RmkCd").Value) <> "" Then .Value = "☞"
                    End If
                    If Trim("" & tmpRs.Fields("DcFg").Value) = "1" Then .Value = .Value & "*"
         .Col = 9:  .Value = Trim("" & tmpRs.Fields("WorkArea").Value)         '-- workarea
         .Col = 10: .Value = Trim("" & tmpRs.Fields("AccDt").Value)         '-- accdt
         .Col = 11: .Value = Trim("" & tmpRs.Fields("AccSeq").Value)         '-- accseq
         .Col = 12: .Value = Trim("" & tmpRs.Fields("TestDiv").Value)         '-- testdiv
         
         If NormalFg And cmdAll.Tag = "1" Then .Action = ActionDeleteRow: .MaxRows = .MaxRows - 1
         
         ReDim Preserve aryMesg(UBound(aryMesg) + 1)
         aryMesg(UBound(aryMesg)) = Trim("" & tmpRs.Fields("Mesg").Value)          '-- 진료과Remark
         
         tmpRs.MoveNext
         
         DisplayOrders = True
      
      Wend
      
     .Row = -1: .Col = 2: .Col2 = 3
     .BlockMode = True
     .AllowCellOverflow = True
     .BlockMode = False
       
     .RowHeight(-1) = 11.5
     .ReDraw = True
'      tmpRs.RsClose
      
      barStatus.Value = barStatus.Max
      DoEvents
      fraStatus.Visible = False
      
      If .MaxRows < 29 Then .MaxRows = 29
      
      .ReDraw = True
   
   End With
   
NoData:
   'QueryFg = False
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   Set tmpRs = Nothing
   
End Function


Private Sub cmdTxtClose_Click()
    fraRstText.Visible = False
End Sub

Private Sub Form_Activate()
    'medMain.lblSubMenu.Caption = Me.Caption
    Me.Left = 4845
    Me.Top = 0

    MsgFg = False

End Sub

Private Sub Form_Terminate()
    StopFg = True
End Sub

'% 처방 선택(Click)하면 해당 결과 디스플레이...
Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
   
   Dim pWorkArea As String
   Dim pAccDt As String
   Dim pAccSeq As String
   Dim pTestDiv As String
   
   If Row = 0 Then Exit Sub
   If OldRow = Row Then Exit Sub
   
   With tblOrdSheet
      
      .Row = Row
      .Col = 1:  If .Value = "" Then Exit Sub
      
      .Col = 9: pWorkArea = .Value
      .Col = 10: pAccDt = .Value
      .Col = 11: pAccSeq = .Value
      .Col = 12: pTestDiv = .Value
      
      If pWorkArea = "" Or pAccDt = "" Or pAccSeq = "" Then
         MsgBox "접수번호가 없습니다. (전산실로 연락바람)", vbExclamation
         Exit Sub
      End If
      
      If OldRow > 0 Then
         .Row = OldRow
         .Col = -1
         .BackColor = OldBackColor
      End If
         
      .Row = Row
      .Col = -1
      OldRow = Row
      OldBackColor = .BackColor
      .BackColor = &HD9ECFF ' &HFCEFE9   ' &HF5FFF4       '연두색
      
      .Col = 8:
      If Not (.Value Like "☞*") Then Exit Sub
      
      Call DisplayResult(pWorkArea, pAccDt, pAccSeq, pTestDiv)
    
      tblResult.TopRow = 1
      ResultFg = True
      
   End With
End Sub

'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                          ByVal pTestDiv As String, Optional pQuery As Boolean = True)
   
   Dim i As Integer, j As Integer
   Dim MyResult As New clsCmtResultReview
   Dim SamTxtBuffer As String
   
   With MyResult
      
      Screen.MousePointer = vbArrowHourglass  '13
      
      Call .ResultMore(pWorkArea, pAccDt, pAccSeq, pTestDiv)
      
      If .ResultCnt = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
      
      
      '결과내역 Display
      tblResult.Row = 1
      tblResult.Row2 = tblResult.MaxRows
      tblResult.Col = 2
      tblResult.Col2 = tblResult.MaxCols
      tblResult.BlockMode = True
      tblResult.AllowCellOverflow = True
      
      '** 수정 특수검사 결과는 조회 안되는 문제
      '-- 원본 =======================================
'      tblResult.Clip = .ResultClipText & .SenClipText
      '===============================================
      
      '-- 수정
      tblResult.Clip = ""
      tblResult.Clip = .ResultClipText & .SenClipText ' & .SpeTextBuffer
      
      tblResult.BlockMode = False
      
      '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
      'If .SortFg Then
      If .SortFg Or .TestDiv = TST_MicTest Then
         tblResult.SortBy = SortByRow
         tblResult.SortKey(1) = 2  '항생제명
         tblResult.SortKeyOrder(1) = SortKeyOrderAscending
         tblResult.Col = -1
         tblResult.Row = .SortStartRow + .OffSet
         tblResult.Row2 = .SortEndRow + .OffSet
         tblResult.Action = ActionSort
         'tblResult.RowsFrozen = .SortStartRow - 1 + .OffSet
         '미생물 결과 : 균명컬럼 Align Left
         tblResult.Row = -1     '1: tblResult.Row2 = tblResult.MaxRows
         tblResult.Col = -1      '1: tblResult.COL2 = tblResult.MaxCols   '7
         tblResult.BlockMode = True
         tblResult.AllowCellOverflow = True
         tblResult.TypeHAlign = TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT   'TypeHAlignLeft
         tblResult.BlockMode = False
         tblResult.ColWidth(2) = 15
         tblResult.ColWidth(3) = 60
      Else
         '일반결과 : 결과컬럼 Align Center
         tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
         tblResult.Col = 3: tblResult.Col2 = 7
         tblResult.BlockMode = True
         tblResult.TypeHAlign = TypeHAlignCenter
         tblResult.BlockMode = False
         tblResult.ColWidth(2) = 11
         tblResult.ColWidth(3) = 9
      End If
      
      '검체리마크 & 풋노트 Display
      If .CommentFg Then
         SamTxtBuffer = .SamTextBuffer
      Else
         SamTxtBuffer = ""
      End If
      
   End With
   
   With tblResult
      .Col = 2: .Col2 = 3
      .Row = 1: .Row2 = .MaxRows
      .BlockMode = True
      
      '** 수정 특수검사 결과는 조회 안되는 문제
      '-- 원본 =======================================
'      txtRstText.Text = .Clip
      '===============================================
      
      '-- 수정 =======================================
      If pTestDiv = "1" Then
        txtRstText.TextRTF = MyResult.SpeTextBuffer
      Else
        txtRstText.Text = .Clip
      End If
      '===============================================
      
      .BlockMode = False
   End With
   
   txtRstText.Text = txtRstText.Text & vbCrLf & SamTxtBuffer
   Call HighlightText(txtRstText, "<< 검사 소견 >>", True)
   Call HighlightText(txtRstText, "<< Supplemental Report >>", False)
   Call HighlightText(txtRstText, "[ Susceptibility test ]", False)
   Call HighlightText(txtRstText, "Antibiotics", False, &HDF6A3E)
   Call HighlightText(txtRstText, "<< Remark >>", False)
   Call HighlightText(txtRstText, "<< Foot Note >>", False)
   
   fraRstText.Visible = True
   fraRstText.ZOrder 0
   Screen.MousePointer = vbDefault
   
End Sub


Private Sub Form_Load()
    
    Me.Left = 4845
    Me.Top = 0
    Me.Show
    
    OrderFg = False
    ResultFg = False
    ClearFg = True
    PtFg = False
    OldRow = 0
'    TopLeftShow = False
    
    'Set MyPatient.MyOraSE = OraSe
'    Set MyPatient.MyDb = DBConn
    

End Sub

'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
   If Not ClearFg Then
      Call ClearRtn
   End If
   StopFg = True
End Sub

'% 환자 ID
Private Sub txtPtId_GotFocus()
   With txtPtId
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

'% 환자정보 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
   End If
End Sub

Private Sub txtPtId_LostFocus()
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If MsgFg Then Exit Sub
      
    On Error GoTo Err_Trap
    
    If txtPtId.Text = "" Then
        txtPtId.SetFocus
        Exit Sub
    End If
    
      With MyPatient
         If .PtntQuery(txtPtId.Text, frmReport.BedinDt) Then
            lblPtNm.Caption = .PtNm
            lblSex.Caption = .SexNm
            lblAge.Caption = .Age
            lblAgeDiv.Caption = .AgeDiv
            lblDeptNm.Caption = .DeptNm
            'lblBedinDt.Caption = Format(.BedInDt, CS_DatelongFormat)
            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DatelongFormat)
            PtFg = True
            ClearFg = False
            Call cmdAll_Click
         Else
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            MsgFg = False
            Me.Enabled = True
            txtPtId.SetFocus
            PtFg = False
            Call txtPtId_GotFocus
            Exit Sub
         End If
      End With
      StopFg = False
'      If ActiveControl.Name <> cmdRefresh.Name Then
'         If dtpFromDate.Enabled Then dtpFromDate.SetFocus
'      End If
      Exit Sub
Err_Trap:
    Resume Next

End Sub

'% Clear 루틴
Private Sub ClearRtn()
   lblPtNm.Caption = ""
   lblSex.Caption = ""
   lblAge.Caption = ""
   lblAgeDiv.Caption = ""
   lblDeptNm.Caption = ""
   lblLocation.Caption = ""
   'lblBedinDt.Caption = ""
   'lblBedoutDt.Caption = ""
   Call TableClear
   ClearFg = True
   OrderFg = False
   MsgFg = False
   StopFg = False
   QueryFg = False
   OldRow = 0
End Sub


Private Sub TableClear()
   tblOrdSheet.MaxRows = 0
   tblOrdSheet.MaxRows = 100
   'tblResult.MaxRows = 0
   'tblResult.MaxRows = 100
   OldRow = 0
'   TopLeftShow = False
End Sub


Private Sub HighlightText(ByVal pTextBox As Object, ByVal pText As String, ByVal InitFg As Boolean, Optional COLOR As Long = &H80&)
   With pTextBox
      If InitFg Then
         .SelStart = 0
         .SelLength = Len(.Text)
         .SelColor = &H0&
         '.SelBold = False
      End If
      
      Dim Point2 As Long
      
      '-- 원본 ==================================================================
'      Point2 = .Find(pText, 0, , rtfWholeWord)
'      While (Point2 <> -1)
'         .SelStart = Point2
'         .SelLength = Len(pText)
'         .SelColor = COLOR         '&HFF8080       '&H8080FF           '&HDF6A3E
'         '.SelBold = True
'         Point2 = .Find(pText, Point2 + Len(pText), , rtfWholeWord)
'      Wend
'      .SelLength = 0
      '==========================================================================
      
      '-- 무한루프로 인한 변경 By MGCHOI ========================================
      Point2 = .Find(pText, 0, , rtfWholeWord)
      If Point2 <> -1 Then
         .SelStart = Point2
         .SelLength = Len(pText)
         .SelProtected = False
         .SelColor = COLOR         '&HFF8080       '&H8080FF           '&HDF6A3E
         '.SelBold = True
      End If
      .SelLength = 0
      '==========================================================================
   End With
End Sub

Public Sub Call_cmdAll_Click()
    Call cmdAll_Click
    cmdAll.SetFocus
End Sub

Public Sub Call_PtId_LostFocus()
    
    txtPtId.SetFocus
    SendKeys "{TAB}"
    'Call txtPtId_LostFocus
End Sub




