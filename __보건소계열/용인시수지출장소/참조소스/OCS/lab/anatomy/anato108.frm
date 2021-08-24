VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anato_Result_Print 
   BorderStyle     =   0  '없음
   Caption         =   "결과지출력"
   ClientHeight    =   8730
   ClientLeft      =   1605
   ClientTop       =   1800
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8730
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   10200
      TabIndex        =   26
      Top             =   3210
      Width           =   1695
      Begin Threed.SSOption optClass 
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   240
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "P"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optClass 
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optClass 
         Height          =   195
         Index           =   2
         Left            =   1140
         TabIndex        =   29
         Top             =   240
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "R"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frmR 
      Height          =   525
      Left            =   10200
      TabIndex        =   30
      Top             =   3690
      Width           =   1695
      Begin Threed.SSOption optClassR 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   180
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "RP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optClassR 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   180
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "RC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1755
      Left            =   10200
      TabIndex        =   6
      Top             =   840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   3096
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1350
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36312
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36312
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1140
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "보고일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   10200
      TabIndex        =   22
      Top             =   2550
      Width           =   1695
      Begin VB.TextBox txtClass 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         MaxLength       =   2
         TabIndex        =   23
         Top             =   180
         Width           =   350
      End
      Begin VB.TextBox txtDateYY 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   435
         MaxLength       =   4
         TabIndex        =   24
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txtSeqnum 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         MaxLength       =   5
         TabIndex        =   25
         Top             =   180
         Width           =   600
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1410
      Left            =   10200
      TabIndex        =   16
      Top             =   4290
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   2487
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optReport 
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   180
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "육안검사"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optReport 
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   420
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Preliminary"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optReport 
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   900
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "결과완료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optReport 
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   20
         Top             =   1140
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Additional"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optReport 
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   660
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "판독"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   915
      Left            =   10200
      TabIndex        =   12
      Top             =   5670
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1614
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optThree 
         Height          =   240
         Left            =   210
         TabIndex        =   13
         Top             =   630
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "3장출력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optTwo 
         Height          =   240
         Left            =   210
         TabIndex        =   14
         Top             =   390
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "2장출력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optOne 
         Height          =   240
         Left            =   210
         TabIndex        =   15
         Top             =   150
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "1장출력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin FPSpread.vaSpread ssResult 
      Height          =   7515
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   10095
      _Version        =   196608
      _ExtentX        =   17806
      _ExtentY        =   13256
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421376
      MaxCols         =   15
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "ANATO108.frx":0000
      VisibleCols     =   14
      VisibleRows     =   500
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   1740
      Left            =   10200
      ScaleHeight     =   1680
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   6600
      Width           =   1692
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   0
         TabIndex        =   3
         Top             =   555
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "출 력"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO108.frx":2653
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   0
         TabIndex        =   2
         Top             =   1110
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "종 료"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO108.frx":266F
      End
      Begin Threed.SSCommand cmdView 
         Height          =   555
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "조 회"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO108.frx":268B
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   732
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12012
      _Version        =   65536
      _ExtentX        =   21188
      _ExtentY        =   1291
      _StockProps     =   15
      Caption         =   "결   과   지   출   력"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   16.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
   End
End
Attribute VB_Name = "Anato_Result_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim GResultP            As String
    Dim LineCount           As Integer

    Dim TabCheck            As Integer


Private Sub Form_Activate()
    optOne.Value = True
    optReport(2).Value = True
    GResultP = "2"

End Sub

Private Sub Form_Load()
    dtFromJeobsu = Dual_Date_Get("yyyy-MM-dd")
    dtToJeobsu = Dual_Date_Get("yyyy-MM-dd")
    
    optClass(0).Value = True
    
End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub optClass_Click(Index As Integer, Value As Integer)
    
    Select Case Index
            Case 0
                   frmR.Visible = False
            Case 1
                   frmR.Visible = False
            Case 2
                   frmR.Visible = True
    End Select
'    optClass(0).Value = True

End Sub

Private Sub optReport_Click(Index As Integer, Value As Integer)
    Select Case Index
           Case 0
                GResultP = "0"      '육안검사
           Case 1
                GResultP = "1"      'preliminary
           Case 2
                GResultP = "2"      '결과완료
           Case 3
                GResultP = "3"      'Additional
           Case 4
                GResultP = "4"      '판독
    End Select
    
End Sub

Private Sub ssResult_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i                   As Integer
    
    If Row = 0 And Col = 1 Then
        ssResult.Col = 1
        ssResult.Row = 0
        If ssResult.Text = "A" Then
            ssResult.Col = 1
            ssResult.Row = 0
            ssResult.Text = "C"
            For i = 1 To ssResult.DataRowCnt
                ssResult.Row = i
                ssResult.Text = "0"
            Next i
        Else
            ssResult.Col = 1
            ssResult.Row = 0
            ssResult.Text = "A"
            For i = 1 To ssResult.DataRowCnt
                ssResult.Row = i
                ssResult.Text = "1"
            Next i
        End If
'    ElseIf Row > 0 And Col = 1 Then
'
    End If
    
End Sub



Private Sub cmdView_Click()
    
    Dim LsPtNo              As String * 8
    Dim LsStatus            As String * 1
    Dim LsCodeKy            As String
    Dim LsDrCode            As String * 6
    Dim LsDeptCode          As String * 4
    Dim LiReccnt            As Integer
    Dim i                   As Integer
    Dim LsRet
    
    Call SSInitialize(ssResult)
    gSFrDate = Format(dtFromJeobsu.Value, "YYYY-MM-DD")
    gSToDate = Format(dtToJeobsu.Value, "YYYY-MM-DD")
    
    If txtDateYY.Text <> "" And txtSeqNum.Text = "" Then
        MsgBox "접수번호입력 Error"
        txtClass.Text = ""
        txtDateYY.Text = ""
        txtSeqNum.Text = ""
        Exit Sub
    End If
    
    If txtDateYY.Text = "" And txtSeqNum.Text <> "" Then
        MsgBox "접수번호입력 Error"
        txtClass.Text = ""
        txtDateYY.Text = ""
        txtSeqNum.Text = ""
        Exit Sub
    End If
    
    Anato_Result_Print.MousePointer = vbHourglass
    
    If GResultP = "0" Then      '육안검사
        strSQL = ""
        strSQL = strSQL & " SELECT a.*, a.RowID, "
        strSQL = strSQL & "        TO_CHAR(a.Indate, 'YYYY-MM-DD') Indate, "
        strSQL = strSQL & "        TO_CHAR(a.JDate, 'YYYY-MM-DD') JDate1, "
        strSQL = strSQL & "        D.DRName  DRNAME2, c.Deptnamek "
        strSQL = strSQL & " FROM   TWANAT_DIAG     a, "
        strSQL = strSQL & "        TWBAS_DEPT      c, "
        strSQL = strSQL & "        TWBAS_DOCTOR    d  "
        If txtDateYY.Text = "" Then
'            strSQL = strSQL & " WHERE  a.JDate >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd') "
'            strSQL = strSQL & "   AND  a.JDate <= TO_DATE('" & gSToDate & "','yyyy-MM-dd') "
            strSQL = strSQL & " WHERE  a.DIAGDATE >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd') "
            strSQL = strSQL & "   AND  a.DIAGDATE <= TO_DATE('" & gSToDate & "','yyyy-MM-dd') "
            If optClass(0).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'P'"
            ElseIf optClass(1).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'C'"
            Else
                strSQL = strSQL & "   AND  a.Class = 'R'"
                If optClassR(0).Value = True Then
                    strSQL = strSQL & "   AND  a.ITEMGB = 'P'"
                Else
                    strSQL = strSQL & "   AND  a.ITEMGB = 'C'"
                End If
            End If
        Else
            strSQL = strSQL & " WHERE  a.Class = '" & txtClass.Text & "' "
            strSQL = strSQL & "   AND  a.DateYY = '" & txtDateYY.Text & "' "
            strSQL = strSQL & "   AND  a.Seqnum = '" & txtSeqNum.Text & "' "
        End If
        strSQL = strSQL & "   AND  a.GbGross = '1' "
        strSQL = strSQL & "   AND  a.GbResult = '0' "
        strSQL = strSQL & "   AND  a.DeptCode  = c.Deptcode(+) "
        strSQL = strSQL & "   AND  a.Drcode    = d.Drcode(+) "
        strSQL = strSQL & "   AND  a.DiagEye  is not null "
        strSQL = strSQL & " ORDER BY  a.CLASS, DATEYY, SEQNUM, JDATE1 "
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
            Anato_Result_Print.MousePointer = vbDefault
            Exit Sub
        End If
        
        i = 0
        Do Until rs.EOF
            ssResult.Row = i + 1
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            ssResult.Col = 2:  ssResult.Text = rs.Fields("CLASS").Value & "-" & _
                                               rs.Fields("DATEYY").Value & "-" & _
                                               rs.Fields("SEQNUM").Value & ""
'            ssResult.Col = 3:  ssResult.Text = rs.Fields("Chief").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("PTNO").Value & ""
            ssResult.Col = 5:  ssResult.Text = rs.Fields("SNAME").Value & ""
            ssResult.Col = 6:  ssResult.Text = Format(rs.Fields("OrderDt").Value & "", "yyyy-mm-dd")
            
            ssResult.Col = 7:  ssResult.Text = Format(rs.Fields("JDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 8:  ssResult.Text = Format(rs.Fields("DIAGDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 9:  ssResult.Text = rs.Fields("Deptnamek").Value & ""
            ssResult.Col = 10:  ssResult.Text = rs.Fields("Drname2").Value & ""
            ssResult.Col = 11: ssResult.Text = rs.Fields("ROOMCODE").Value & ""
'            ssResult.Col = 12: ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여") & "/" & rs.Fields("AGEYY").Value & ""
            Select Case rs.Fields("Sex").Value
                   Case "M"
                            ssResult.Col = 12: ssResult.Text = "남" & "/" & rs.Fields("AGEYY").Value & ""
                   Case "F"
                            ssResult.Col = 12: ssResult.Text = "여" & "/" & rs.Fields("AGEYY").Value & ""
                   Case Else
                            ssResult.Col = 12: ssResult.Text = "  " & "/" & rs.Fields("AGEYY").Value & ""
            End Select
            ssResult.Col = 13: ssResult.Text = rs.Fields("speSlide").Value & ""
'            ssResult.Col = 14: ssResult.Text = rs.Fields("Chief2").Value & ""
            ssResult.Col = 15: ssResult.Text = rs.Fields("ROWID").Value & ""
            
            rs.MoveNext: i = i + 1
        Loop
        AdoCloseSet rs
    ElseIf GResultP = "1" Then   ' Preliminary
        strSQL = ""
        strSQL = strSQL & " SELECT a.*, a.RowID, "
        strSQL = strSQL & "        TO_CHAR(a.Indate, 'YYYY-MM-DD') Indate, "
        strSQL = strSQL & "        TO_CHAR(a.JDate, 'YYYY-MM-DD') JDate1, "
'        strSQL = strSQL & "        B.DRNAME  DRNAME1 , "
        strSQL = strSQL & "        D.DRName  DRNAME2, c.Deptnamek "
        strSQL = strSQL & " FROM   TWANAT_DIAG     a, "
'        strSQL = strSQL & "        TWBAS_DOCTOR    b, "
        strSQL = strSQL & "        TWBAS_DEPT      c, "
        strSQL = strSQL & "        TWBAS_DOCTOR    d  "
        If txtDateYY.Text = "" Then
            strSQL = strSQL & " WHERE  a.DIAGDATE >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd') "
            strSQL = strSQL & "   AND  a.DIAGDATE <= TO_DATE('" & gSToDate & "','yyyy-MM-dd') "
            If optClass(0).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'P'"
            ElseIf optClass(1).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'C'"
            Else
                strSQL = strSQL & "   AND  a.Class = 'R'"
                If optClassR(0).Value = True Then
                    strSQL = strSQL & "   AND  a.ITEMGB = 'P'"
                Else
                    strSQL = strSQL & "   AND  a.ITEMGB = 'C'"
                End If
            End If
        Else
            strSQL = strSQL & " WHERE  a.Class = '" & txtClass.Text & "' "
            strSQL = strSQL & "   AND  a.DateYY = '" & txtDateYY.Text & "' "
            strSQL = strSQL & "   AND  a.Seqnum = '" & txtSeqNum.Text & "' "
        End If
        strSQL = strSQL & "   AND  a.GbGross = '1' "
        strSQL = strSQL & "   AND  a.GbResult  = '2' "
'        strSQL = strSQL & "   AND  A.Chief     = B.DRCODE(+) "
        strSQL = strSQL & "   AND  a.DeptCode  = c.Deptcode(+) "
        strSQL = strSQL & "   AND  a.Drcode    = d.Drcode(+) "
        strSQL = strSQL & "   AND  a.DiagPre   is not null "
        strSQL = strSQL & " ORDER BY JDATE1, a.CLASS, DATEYY, SEQNUM "
        
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
            Anato_Result_Print.MousePointer = vbDefault
            Exit Sub
        End If
        
        i = 0
        Do Until rs.EOF
            
'            If rs.Fields("DiagPre").Value Is Null Then
'                rs.MoveNext
'            End If
            ssResult.Row = i + 1
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            ssResult.Col = 2:  ssResult.Text = rs.Fields("CLASS").Value & "-" & _
                                               rs.Fields("DATEYY").Value & "-" & _
                                               rs.Fields("SEQNUM").Value & ""
'            ssResult.Col = 3:  ssResult.Text = rs.Fields("DRName1").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("PTNO").Value & ""
            ssResult.Col = 5:  ssResult.Text = rs.Fields("SNAME").Value & ""
            
            ssResult.Col = 6:  ssResult.Text = Format(rs.Fields("OrderDt").Value & "", "yyyy-mm-dd")
            
            ssResult.Col = 7:  ssResult.Text = Format(rs.Fields("JDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 8:  ssResult.Text = Format(rs.Fields("DIAGDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 9:  ssResult.Text = rs.Fields("Deptnamek").Value & ""
            ssResult.Col = 10:  ssResult.Text = rs.Fields("Drname2").Value & ""
            ssResult.Col = 11: ssResult.Text = rs.Fields("ROOMCODE").Value & ""
            
'            ssResult.Col = 12: ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여") & "/" & rs.Fields("AGEYY").Value & ""
            Select Case rs.Fields("Sex").Value
                   Case "M"
                            ssResult.Col = 12: ssResult.Text = "남" & "/" & rs.Fields("AGEYY").Value & ""
                   Case "F"
                            ssResult.Col = 12: ssResult.Text = "여" & "/" & rs.Fields("AGEYY").Value & ""
                   Case Else
                            ssResult.Col = 12: ssResult.Text = "  " & "/" & rs.Fields("AGEYY").Value & ""
            End Select
            
            ssResult.Col = 13: ssResult.Text = rs.Fields("speSlide").Value & ""
'            ssResult.Col = 14: ssResult.Text = rs.Fields("ROWID").Value & ""
            ssResult.Col = 15: ssResult.Text = rs.Fields("ROWID").Value & ""
            
            rs.MoveNext: i = i + 1
        Loop
        AdoCloseSet rs
    
    ElseIf GResultP = "2" Then          '결과완료
        strSQL = ""
        strSQL = strSQL & " SELECT a.*, a.RowID, "
        strSQL = strSQL & "        TO_CHAR(a.Indate, 'YYYY-MM-DD') Indate, "
        strSQL = strSQL & "        TO_CHAR(a.DiagDate, 'YYYY-MM-DD') DiagDate1, "
        strSQL = strSQL & "        b.DRName DRNAME1 , d.DRName DRNAME2, e.DRName DRNAME3, c.Deptnamek "
        strSQL = strSQL & " FROM   TWANAT_DIAG     a, "
        strSQL = strSQL & "        TWBAS_DOCTOR    b, "
        strSQL = strSQL & "        TWBAS_DEPT      c, "
        strSQL = strSQL & "        TWBAS_DOCTOR    d, "
        strSQL = strSQL & "        TWBAS_DOCTOR    e  "
        If txtDateYY.Text = "" Then
            strSQL = strSQL & " WHERE  a.DIAGDATE >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd') "
            strSQL = strSQL & "   AND  a.DIAGDATE <= TO_DATE('" & gSToDate & "','yyyy-MM-dd') "
            If optClass(0).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'P'"
            ElseIf optClass(1).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'C'"
            Else
                strSQL = strSQL & "   AND  a.Class = 'R'"
                If optClassR(0).Value = True Then
                    strSQL = strSQL & "   AND  a.ITEMGB = 'P'"
                Else
                    strSQL = strSQL & "   AND  a.ITEMGB = 'C'"
                End If
            End If
        Else
            strSQL = strSQL & " WHERE  a.Class = '" & txtClass.Text & "' "
            strSQL = strSQL & "   AND  a.DateYY = '" & txtDateYY.Text & "' "
            strSQL = strSQL & "   AND  a.Seqnum = '" & txtSeqNum.Text & "' "
        End If
        strSQL = strSQL & "   AND  a.GbResult >= '4' "
        strSQL = strSQL & "   AND  a.Chief     = b.DRCODE(+) "
        strSQL = strSQL & "   AND  a.Chief2    = e.DRCODE(+) "
        strSQL = strSQL & "   AND  a.Drcode    = d.Drcode(+) "
        strSQL = strSQL & "   AND  a.DeptCode  = c.Deptcode(+) "
        strSQL = strSQL & " ORDER BY DIAGDATE1, a.CLASS, DATEYY, SEQNUM "
        
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
            Anato_Result_Print.MousePointer = vbDefault
            Exit Sub
        End If
        
        i = 0
        Do Until rs.EOF
            
            ssResult.Row = i + 1
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            ssResult.Col = 2:  ssResult.Text = rs.Fields("CLASS").Value & "-" & _
                                               rs.Fields("DATEYY").Value & "-" & _
                                               rs.Fields("SEQNUM").Value & ""
            ssResult.Col = 3:  ssResult.Text = rs.Fields("DRName1").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("PTNO").Value & ""
            ssResult.Col = 5:  ssResult.Text = rs.Fields("SNAME").Value & ""
            ssResult.Col = 6:  ssResult.Text = Format(rs.Fields("OrderDt").Value & "", "yyyy-mm-dd")
            
            ssResult.Col = 7:  ssResult.Text = Format(rs.Fields("JDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 8:  ssResult.Text = Format(rs.Fields("DIAGDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 9:  ssResult.Text = rs.Fields("Deptnamek").Value & ""
            ssResult.Col = 10:  ssResult.Text = rs.Fields("Drname2").Value & ""
            ssResult.Col = 11: ssResult.Text = rs.Fields("ROOMCODE").Value & ""
'            ssResult.Col = 12: ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여") & "/" & rs.Fields("AGEYY").Value & ""
            
            Select Case rs.Fields("Sex").Value
                   Case "M"
                            ssResult.Col = 12: ssResult.Text = "남" & "/" & rs.Fields("AGEYY").Value & ""
                   Case "F"
                            ssResult.Col = 12: ssResult.Text = "여" & "/" & rs.Fields("AGEYY").Value & ""
                   Case Else
                            ssResult.Col = 12: ssResult.Text = "  " & "/" & rs.Fields("AGEYY").Value & ""
            End Select
            
            ssResult.Col = 13: ssResult.Text = rs.Fields("speSlide").Value & ""
            ssResult.Col = 14: ssResult.Text = rs.Fields("DRName3").Value & ""
            ssResult.Col = 15: ssResult.Text = rs.Fields("ROWID").Value & ""
            
            rs.MoveNext: i = i + 1
        Loop
        AdoCloseSet rs
    
    ElseIf GResultP = "3" Then          'Additional
        strSQL = ""
        strSQL = strSQL & " SELECT a.*, a.RowID, "
        strSQL = strSQL & "        TO_CHAR(a.Indate, 'YYYY-MM-DD') Indate, "
        strSQL = strSQL & "        TO_CHAR(a.DiagDate, 'YYYY-MM-DD') DiagDate1, "
        strSQL = strSQL & "        b.DRName DRNAME1, c.Deptnamek, d.Drname DRNAME2 "
        strSQL = strSQL & " FROM   TWANAT_DIAG     a, "
        strSQL = strSQL & "        TWBAS_DOCTOR    b, "
        strSQL = strSQL & "        TWBAS_DEPT      c, "
        strSQL = strSQL & "        TWBAS_DOCTOR    d  "
        If txtDateYY.Text = "" Then
            strSQL = strSQL & " WHERE  a.DIAGDATE >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd') "
            strSQL = strSQL & "   AND  a.DIAGDATE <= TO_DATE('" & gSToDate & "','yyyy-MM-dd') "
            If optClass(0).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'P'"
            ElseIf optClass(1).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'C'"
            Else
                strSQL = strSQL & "   AND  a.Class = 'R'"
                If optClassR(0).Value = True Then
                    strSQL = strSQL & "   AND  a.ITEMGB = 'P'"
                Else
                    strSQL = strSQL & "   AND  a.ITEMGB = 'C'"
                End If
            End If
        Else
            strSQL = strSQL & " WHERE  a.Class = '" & txtClass.Text & "' "
            strSQL = strSQL & "   AND  a.DateYY = '" & txtDateYY.Text & "' "
            strSQL = strSQL & "   AND  a.Seqnum = '" & txtSeqNum.Text & "' "
        End If
        strSQL = strSQL & "   AND  a.GbResult = '9' "
        strSQL = strSQL & "   AND  A.Chief     = B.DRCODE(+) "
        strSQL = strSQL & "   AND  a.Drcode    = d.Drcode(+) "
        strSQL = strSQL & "   AND  a.DeptCode  = c.Deptcode(+) "
        strSQL = strSQL & " ORDER BY DIAGDATE1, a.CLASS, DATEYY, SEQNUM "
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
            Anato_Result_Print.MousePointer = vbDefault
            Exit Sub
        End If
        
        i = 0
        Do Until rs.EOF
            ssResult.Row = i + 1
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            ssResult.Col = 2:  ssResult.Text = rs.Fields("CLASS").Value & "-" & _
                                               rs.Fields("DATEYY").Value & "-" & _
                                               rs.Fields("SEQNUM").Value & ""
            ssResult.Col = 3:  ssResult.Text = rs.Fields("DRName1").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("PTNO").Value & ""
            ssResult.Col = 5:  ssResult.Text = rs.Fields("SNAME").Value & ""
            
            ssResult.Col = 6:  ssResult.Text = Format(rs.Fields("OrderDt").Value & "", "yyyy-mm-dd")
            ssResult.Col = 7:  ssResult.Text = Format(rs.Fields("JDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 8:  ssResult.Text = Format(rs.Fields("DIAGDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 9:  ssResult.Text = rs.Fields("Deptnamek").Value & ""
            ssResult.Col = 10:  ssResult.Text = rs.Fields("Drname2").Value & ""
            ssResult.Col = 11: ssResult.Text = rs.Fields("ROOMCODE").Value & ""
'            ssResult.Col = 12: ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여") & "/" & rs.Fields("AGEYY").Value & ""
            Select Case rs.Fields("Sex").Value
                   Case "M"
                            ssResult.Col = 12: ssResult.Text = "남" & "/" & rs.Fields("AGEYY").Value & ""
                   Case "F"
                            ssResult.Col = 12: ssResult.Text = "여" & "/" & rs.Fields("AGEYY").Value & ""
                   Case Else
                            ssResult.Col = 12: ssResult.Text = "  " & "/" & rs.Fields("AGEYY").Value & ""
            End Select
            ssResult.Col = 13: ssResult.Text = rs.Fields("speSlide").Value & ""
'            ssResult.Col = 14: ssResult.Text = rs.Fields("ROWID").Value & ""
            ssResult.Col = 15: ssResult.Text = rs.Fields("ROWID").Value & ""
            
            
            rs.MoveNext: i = i + 1
        Loop
        AdoCloseSet rs
    
    ElseIf GResultP = "4" Then          '판독
        strSQL = ""
        strSQL = strSQL & " SELECT a.*, a.RowID, "
        strSQL = strSQL & "        TO_CHAR(a.Indate, 'YYYY-MM-DD') Indate, "
        strSQL = strSQL & "        TO_CHAR(a.DiagDate, 'YYYY-MM-DD') DiagDate1, "
        strSQL = strSQL & "        c.Deptnamek, D.DRNAME DRNAME2 "
        strSQL = strSQL & " FROM   TWANAT_DIAG     a, "
        strSQL = strSQL & "        TWBAS_DEPT      c, "
        strSQL = strSQL & "        TWBAS_DOCTOR    d  "
        If txtDateYY.Text = "" Then
'            strSQL = strSQL & " WHERE  a.GbResult = '3' "
            strSQL = strSQL & " WHERE  a.DIAGDATE >= TO_DATE('" & gSFrDate & "','yyyy-MM-dd') "
            strSQL = strSQL & "   AND  a.DIAGDATE <= TO_DATE('" & gSToDate & "','yyyy-MM-dd') "
            strSQL = strSQL & "   AND  a.GbResult = '3' "
            If optClass(0).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'P'"
            ElseIf optClass(1).Value = True Then
                strSQL = strSQL & "   AND  a.Class = 'C'"
            Else
                strSQL = strSQL & "   AND  a.Class = 'R'"
                If optClassR(0).Value = True Then
                    strSQL = strSQL & "   AND  a.ITEMGB = 'P'"
                Else
                    strSQL = strSQL & "   AND  a.ITEMGB = 'C'"
                End If
            End If
        Else
            strSQL = strSQL & " WHERE  a.Class = '" & txtClass.Text & "' "
            strSQL = strSQL & "   AND  a.DateYY = '" & txtDateYY.Text & "' "
            strSQL = strSQL & "   AND  a.Seqnum = '" & txtSeqNum.Text & "' "
        End If
        strSQL = strSQL & "   AND  a.DeptCode  = c.Deptcode(+) "
        strSQL = strSQL & "   AND  a.Drcode    = d.Drcode(+) "
        strSQL = strSQL & " ORDER BY DIAGDATE1, a.CLASS, DATEYY, SEQNUM "
        Result = AdoOpenSet(rs, strSQL)
        
        If Result = False Then
            Anato_Result_Print.MousePointer = vbDefault
            Exit Sub
        End If
        
        i = 0
        Do Until rs.EOF
            ssResult.Row = i + 1
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            If Val(rs.Fields("speSlide").Value & "") > "0" Then
                ssResult.ForeColor = RGB(0, 0, 255)
            Else
                ssResult.ForeColor = RGB(0, 0, 0)
            End If
            
            ssResult.Col = 2:  ssResult.Text = rs.Fields("CLASS").Value & "-" & _
                                               rs.Fields("DATEYY").Value & "-" & _
                                               rs.Fields("SEQNUM").Value & ""
'            ssResult.Col = 3:  ssResult.Text = rs.Fields("Name").Value & ""
            ssResult.Col = 4:  ssResult.Text = rs.Fields("PTNO").Value & ""
            ssResult.Col = 5:  ssResult.Text = rs.Fields("SNAME").Value & ""
            
            ssResult.Col = 6:  ssResult.Text = Format(rs.Fields("OrderDt").Value & "", "yyyy-mm-dd")
            ssResult.Col = 7:  ssResult.Text = Format(rs.Fields("JDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 8:  ssResult.Text = Format(rs.Fields("DIAGDATE").Value & "", "yyyy-mm-dd")
            ssResult.Col = 9:  ssResult.Text = rs.Fields("Deptnamek").Value & ""
            ssResult.Col = 10:  ssResult.Text = rs.Fields("Drname2").Value & ""
            ssResult.Col = 11: ssResult.Text = rs.Fields("ROOMCODE").Value & ""
'            ssResult.Col = 12: ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여") & "/" & rs.Fields("AGEYY").Value & ""
            Select Case rs.Fields("Sex").Value
                   Case "M"
                            ssResult.Col = 12: ssResult.Text = "남" & "/" & rs.Fields("AGEYY").Value & ""
                   Case "F"
                            ssResult.Col = 12: ssResult.Text = "여" & "/" & rs.Fields("AGEYY").Value & ""
                   Case Else
                            ssResult.Col = 12: ssResult.Text = "  " & "/" & rs.Fields("AGEYY").Value & ""
            End Select
            ssResult.Col = 13: ssResult.Text = rs.Fields("speSlide").Value & ""
'            ssResult.Col = 14: ssResult.Text = rs.Fields("ROWID").Value & ""
            ssResult.Col = 15: ssResult.Text = rs.Fields("ROWID").Value & ""
            
            rs.MoveNext: i = i + 1
        Loop
        AdoCloseSet rs
    
    End If
  
    txtClass.Text = ""
    txtDateYY.Text = ""
    txtSeqNum.Text = ""
  
    Anato_Result_Print.MousePointer = vbDefault
  
End Sub


Private Sub cmdPrint_Click()
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim Count               As Integer
    Dim LsJDate             As String * 10
    Dim LsDateYY            As String * 4
    Dim LsODate             As String * 10
    
    Dim LsClass             As String * 2
    Dim LsSeqNum            As String * 5
    Dim LsPtNo              As String * 8
    Dim LsEye               As String
    Dim LsDescr             As String
    Dim LsSname             As String * 10
    Dim LsSexAge            As String * 6
    Dim LsDiagdate          As String * 10
    Dim LsRoomCode          As String * 6
    Dim LsDpName            As String * 16
    Dim LsDrName            As String * 8
    Dim LsChiefName         As String * 10
    Dim LsChiefName2        As String * 10
    Dim LsAnatNo            As String * 13
    Dim LsLineString        As String
    Dim LsitemGb            As String
    
    Dim SpecialChar         As String
    Dim sSpecial(30)        As String
    Dim lsnumber
    
    Dim LsDescr_temp
    
    Dim SLIDENO
    
    Dim LsElectro
    Dim LsFlow
         
    LsJDate = ""
    LsClass = ""
    LsDateYY = ""
    LsSeqNum = ""
    LsPtNo = ""
    LsSname = ""
    LsSexAge = ""
    LsDiagdate = ""
    LsRoomCode = ""
    LsDpName = ""
    LsDrName = ""
    LsChiefName = ""     '
    LsChiefName2 = ""     '
    LsLineString = ""
    
    If ssResult.DataRowCnt = 0 Then Exit Sub
      
    Dim FCheck              As Boolean
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        If ssResult.Text = "1" Then
            FCheck = True
        End If
    Next i
    
    If FCheck = False Then
        MsgBox " 출력할 DATA를 선택하십시요. "
        Exit Sub
    End If
    
    If optOne.Value = True Then
        Count = 1
    ElseIf optTwo.Value = True Then
        Count = 2
    ElseIf optThree.Value = True Then
        Count = 3
    ElseIf optOne.Value = False And optTwo.Value = False And optThree.Value = False Then
        Count = 2
    End If
    
    Anato_Result_Print.MousePointer = vbHourglass
    
    For i = 1 To ssResult.DataRowCnt
        
        ssResult.Row = i
        ssResult.Col = 1 ' 1 '0   ' 1
        
        If ssResult.Text = "1" Then
          For k = 1 To Count
            
            Dim L   As Integer
            
            For L = 1 To 30
                sSpecial(L) = ""
            Next L
            
            LineCount = 0
        
            RSet lsnumber = i
            ssResult.Col = 7:        LsJDate = ssResult.Text
            ssResult.Col = 2:        LsAnatNo = ssResult.Text
                                     LsClass = MidH(ssResult.Text, 1, 2)
                                     LsDateYY = MidH(ssResult.Text, 4, 4)
                                     LsSeqNum = MidH(ssResult.Text, 9, 5)
            ssResult.Col = 4:        LsPtNo = ssResult.Text
            LsDescr = ""
            
            strSQL = ""
            strSQL = strSQL & "  SELECT  DiagEye, DiagPre, DESCR, DiagAdd,itemgb "    'DIAGNO
            strSQL = strSQL & "    FROM  TWANAT_DIAG"
            strSQL = strSQL & "   WHERE  JDATE  =  TO_DATE( '" & LsJDate & "','yyyy-MM-dd')"
            strSQL = strSQL & "     AND  CLASS  = '" & LsClass & "'"
            strSQL = strSQL & "     AND  DATEYY = '" & LsDateYY & "'"
            strSQL = strSQL & "     AND  SEQNUM = " & LsSeqNum
    
            Result = AdoOpenSet(rs, strSQL)
            
            If Result = False Then
                Anato_Result_Print.MousePointer = vbDefault
                Exit Sub
            End If
            
            LsitemGb = rs.Fields("itemgb").Value & ""
            
            '###################################
            
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            
            Select Case GResultP
                   Case 0, 1
                   Case Else
                        GoSub SUB_HEAD_PRINT
            End Select
            
            
            Select Case GResultP
                   Case "0"
                         LsDescr = rs.Fields("DiagEye").Value & ""
                   Case "1"
                         LsDescr = rs.Fields("DiagPre").Value & ""
                   Case "2"
                         LsEye = rs.Fields("DiagEye").Value & ""
                         LsDescr = rs.Fields("DESCR").Value & ""
                   Case "3"
                         LsDescr = rs.Fields("DiagAdd").Value & ""
                   Case "4"
                         LsEye = rs.Fields("DiagEye").Value & ""
                         LsDescr = rs.Fields("DESCR").Value & ""
            End Select
            
            AdoCloseSet rs
            
            ssResult.Col = 3:        LsChiefName = ssResult.Text    '판독자
            ssResult.Col = 5:        LsSname = ssResult.Text        '환자명
            ssResult.Col = 6:        LsODate = ssResult.Text        'Order Date

            ssResult.Col = 8:        LsDiagdate = ssResult.Text     '판독일
            ssResult.Col = 9:        LsDpName = ssResult.Text       '의뢰과
            ssResult.Col = 10:       LsDrName = ssResult.Text       '의뢰의사
            ssResult.Col = 11:       LsRoomCode = ssResult.Text     '병실
            ssResult.Col = 12:       LsSexAge = ssResult.Text       '성별/나이
            ssResult.Col = 14:       LsChiefName2 = Trim(ssResult.Text)   '판독자2
            
            Printer.FontName = "돋움체" '"바탕체"
            Printer.FontSize = 11
            
            Printer.FontBold = True
            
            Printer.FontItalic = False
            Printer.FontUnderline = False
            Printer.Print
            
'           Select Case GResultP
'                   Case "0"
'                         Printer.Print Tab(8); "육안검사";
'                   Case "1"
'                         Printer.Print Tab(8); "Preliminary";
'                   Case "2"
'                         Printer.Print Tab(8); "판독결과";
'                   Case "3"
'                         Printer.Print Tab(8); "Additional";
'            End Select
            
            
            Select Case GResultP
                   Case 0, 1
                   Case Else
                        
                        
                        Printer.FontName = "돋움체" '"바탕체"
                        Printer.FontSize = 18
                        Printer.Print Tab(7); "병리번호:  " & LsAnatNo
                        
                        Printer.FontName = "돋움체" '"바탕체"
                        Printer.FontSize = 11
                        
                        
                        Printer.Print
                        Printer.Print
                        
                        Printer.Print Tab(10); "의 뢰  의 사 : " & LsDrName;
                        Printer.Print Tab(39); "검사물의뢰일 : " & LsODate;
                        
                        Printer.Print Tab(10); "검사물접수일 : " & LsJDate;
                        Printer.Print Tab(39); "보   고   일 : " & LsDiagdate;
                        
                        Printer.Print
            
                        Printer.CurrentX = 8600:          Printer.CurrentY = 3000
                        
                        Printer.Print Tab(73); "등록번호  : " & LsPtNo
                        Printer.Print Tab(73); "환 자 명  : " & LsSname
                        Printer.Print Tab(73); "성별/나이 : " & LsSexAge
                        Printer.Print Tab(73); "의 뢰 과  : " & LsDpName
                        Printer.Print Tab(73); "병    실  : " & IIf(LsRoomCode = "      ", "외래", LsRoomCode)
            
            
            End Select
            
            '###################################
            Select Case GResultP
                   Case 0, 1
                   Case Else
                        GoSub SUB_LINE_PRINT1
            End Select
 
            '###################################
            Select Case GResultP
                   Case 0, 1
                   Case Else
                        GoSub SUB_LINE_PRINT2
            End Select
            
            '###################################
            
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '육안검사 출력
            Select Case GResultP
                   Case 0
                        LsDescr_temp = LsDescr
                        GoSub RESULT_PRINT
                        GoSub SUB_PRINT_PATIENT_RESULT_EYE
                   
                   Case 2
                        LsDescr_temp = LsDescr
                        GoSub SUB_PRINT_PATIENT_RESULT
             
                   Case 3
                        LsDescr_temp = LsDescr
                        GoSub RESULT_PRINT
                   
                   Case 4
                        LsDescr_temp = LsDescr
                        GoSub SUB_PRINT_PATIENT_RESULT
             
             End Select
            
            
            LsDescr = LsDescr_temp
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '결과출력
            Select Case GResultP
                   Case 2
                        GoSub RESULT_PRINT
                   Case 4
                        GoSub RESULT_PRINT
            End Select
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
'            Printer.FontName = "돋움체" '"바탕체"
'            Printer.FontSize = 11
'            Printer.FontBold = False
'            Printer.FontItalic = False
'            Printer.FontUnderline = False
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '특수검사 결과 출력
            
            Select Case GResultP
                   Case 2, 4
             
                            strSQL = ""
                            strSQL = strSQL & " SELECT * "
                            strSQL = strSQL & " FROM   TWANAT_DIAG "
                            strSQL = strSQL & " WHERE  Class  = '" & LsClass & "' "
                            strSQL = strSQL & "   AND  DateYY  = '" & LsDateYY & "' "
                            strSQL = strSQL & "   AND  SeqNum  = '" & LsSeqNum & "' "
                                                     
                            Result = AdoOpenSet(rs, strSQL)
                            If Result Then
                                Do Until rs.EOF
                                    For j = 1 To 30
                                        SpecialChar = "SPECIAL" & Format(j, "00")
                                        sSpecial(j) = Trim(rs.Fields(SpecialChar).Value & "")
                                    Next j
                                    
                                    SLIDENO = Trim(rs.Fields("slid").Value & "")
                                    
                                    LsElectro = Trim(rs.Fields("ELECTROSCOPE").Value & "")
                                    LsFlow = Trim(rs.Fields("FLOW").Value & "")
                                    
                                    rs.MoveNext
                                Loop
                            End If
                            
                            
                            '            If LsClass = "B " Or LsClass = "P" Then
                            '                Printer.Print Tab(7); "SLIDE : H&E   (" & SLIDENO & ")"
                            '                Printer.Print
                            '            Else
                            '                Printer.Print Tab(7); "SLIDE : SMEAR (" & SLIDENO & ")"
                            '                Printer.Print
                            '            End If
                            
                            Dim FlagSpecial1
                            Dim FlagSpecial2
                            Dim FlagSpecial3
                            Dim FlagSpecial4
                            Dim FlagSpecial5
                            
                            Dim Flag_Data1
                            Dim Flag_Data2
                            Dim Flag_Data3
                            Dim Flag_Data4
                            Dim Flag_Data5
                            
                            
                            FlagSpecial1 = ""
                            FlagSpecial2 = ""
                            FlagSpecial3 = ""
                            FlagSpecial4 = ""
                            FlagSpecial5 = ""
                            
                            Flag_Data1 = ""
                            Flag_Data2 = ""
                            Flag_Data3 = ""
                            Flag_Data4 = ""
                            Flag_Data5 = ""
                            
                            
                            
                            For j = 1 To 30
                                Select Case sSpecial(j)
                                        Case "853001" To "853999"
                                             FlagSpecial1 = 1
                                             Flag_Data1 = Flag_Data1 & Special_Load(sSpecial(j)) & ", "
                                        Case "857001" To "857999"
                                             FlagSpecial2 = 1
                                             Flag_Data2 = Flag_Data2 & Special_Load(sSpecial(j)) & ", "
                                        Case "854001" To "854999"
                                             FlagSpecial3 = 1
                                             Flag_Data3 = Flag_Data3 & Special_Load(sSpecial(j)) & ", "
                                        Case "856001" To "856999"
                                             FlagSpecial4 = 1
                                             Flag_Data4 = Flag_Data4 & Special_Load(sSpecial(j)) & ", "
                                        Case "855001"
                                             FlagSpecial5 = 1
                                             Flag_Data5 = "Y "
                                End Select
                            Next j
                            
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            If FlagSpecial1 = 1 Then
                                 If Len(Flag_Data1) <= 60 Then
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "SPECIAL STAIN : " & Mid(Flag_Data1, 1, Len(Flag_Data1) - 2)
                                     GoSub SUB_PRINT_NEWPAGE
                                 
                                 ElseIf Len(Flag_Data1) > 60 And Len(Flag_Data1) <= 120 Then
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "              : " & Mid(Flag_Data1, 61, Len(Flag_Data1) - 62)
                                     GoSub SUB_PRINT_NEWPAGE
                                 
                                 ElseIf Len(Flag_Data1) > 120 And Len(Flag_Data1) <= 180 Then
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "              : " & Mid(Flag_Data1, 61, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "              : " & Mid(Flag_Data1, 121, Len(Flag_Data1) - 122)
                                     GoSub SUB_PRINT_NEWPAGE
                                 
                                 ElseIf Len(Flag_Data1) > 181 Then
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "SPECIAL STAIN : " & Mid(Flag_Data1, 1, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "              : " & Mid(Flag_Data1, 61, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "              : " & Mid(Flag_Data1, 121, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "              : " & Mid(Flag_Data1, 181, Len(Flag_Data1) - 182)
                                     GoSub SUB_PRINT_NEWPAGE
                                 End If
                            End If
                            If FlagSpecial2 = 1 Then
                                If Len(Flag_Data2) <= 60 Then
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, Len(Flag_Data2) - 2)
                                     GoSub SUB_PRINT_NEWPAGE
                                
                                ElseIf Len(Flag_Data2) > 60 And Len(Flag_Data1) <= 120 Then
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, 60)
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "                          : " & Mid(Flag_Data2, 61, Len(Flag_Data2) - 62)
                                     GoSub SUB_PRINT_NEWPAGE
                                
                                ElseIf Len(Flag_Data2) > 120 And Len(Flag_Data1) <= 180 Then
                                     LineCount = LineCount + 3
                                     Printer.Print Tab(10); "IMMUNOHISTOCHEMICAL STAIN : " & Mid(Flag_Data2, 1, 60)
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "                          : " & Mid(Flag_Data2, 61, 60)
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "                          : " & Mid(Flag_Data2, 121, Len(Flag_Data2) - 122)
                                     GoSub SUB_PRINT_NEWPAGE
                                
                                ElseIf Len(Flag_Data2) > 180 Then
                                     LineCount = LineCount + 4
                                     Printer.Print Tab(10); "Immunohistochemical stain : " & Mid(Flag_Data2, 1, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "                          : " & Mid(Flag_Data2, 61, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "                          : " & Mid(Flag_Data2, 121, 60)
                                     GoSub SUB_PRINT_NEWPAGE
                                     
                                     LineCount = LineCount + 1
                                     Printer.Print Tab(10); "                          : " & Mid(Flag_Data2, 181, Len(Flag_Data2) - 182)
                                     GoSub SUB_PRINT_NEWPAGE
                                End If
                                 
                            End If
                            If FlagSpecial3 = 1 Then
                                 LineCount = LineCount + 1
                                 Printer.Print Tab(10); "IMMUNOFLUORESCENCE : " & Mid(Flag_Data3, 1, Len(Flag_Data3) - 2)
                                 GoSub SUB_PRINT_NEWPAGE
                            End If
                            If FlagSpecial4 = 1 Then
                                 LineCount = LineCount + 1
                                 Printer.Print Tab(10); "ENZYME HISTOCHEMISTRY : " & Mid(Flag_Data4, 1, Len(Flag_Data4) - 2)
                                 GoSub SUB_PRINT_NEWPAGE
                            End If
                            
                            'If FlagSpecial5 = 1 Then
                            '     LineCount = LineCount + 1
                            '     Printer.Print Tab(10); "ELECTRON MICROSCOPIC EXAM : " & Flag_Data5
                            '     GoSub SUB_PRINT_NEWPAGE
                            'End If
                            
                            If LsElectro = "Y" Then
                                 LineCount = LineCount + 1
                                 Printer.Print Tab(10); "ELECTRON MICROSCOPIC EXAM : Y "
                                 GoSub SUB_PRINT_NEWPAGE
                            End If
                            
                            If LsFlow = "Y" Then
                                 LineCount = LineCount + 1
                                 Printer.Print Tab(10); "FLOW CYTOMETRY : Y "
                                 GoSub SUB_PRINT_NEWPAGE
                            End If
                            
                            AdoCloseSet rs
             
             
             End Select
            
            
            Select Case GResultP
                   Case 0, 1
                   Case Else
            
                        Printer.FontName = "돋움체"
                        Printer.FontSize = 11
                        Printer.FontBold = True
                        Printer.FontItalic = False
                        Printer.FontUnderline = False
                        
'                        Printer.CurrentX = 700:          Printer.CurrentY = 14100
                        Printer.CurrentX = 700:          Printer.CurrentY = 15100
                        Printer.Print
                        If Trim(LsChiefName2) = "" Then
                            Printer.Print Tab(70); "판독의사 : " & LsChiefName
                        Else
                            Printer.Print Tab(70); "판독의사 : " & Trim(LsChiefName) & " / " & Trim(LsChiefName2)
                        End If
                        GoSub SUB_LINE_PRINT3
            
            End Select
            
            Printer.EndDoc
          Next k
          GoSub SUB_PRINT_COUNT_UPDATE
        End If
        
        
    Next i
    
    Anato_Result_Print.MousePointer = vbDefault
    
    Exit Sub
    
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

SUB_HEAD_PRINT:
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    
    Printer.FontName = "돋움체"
    Printer.FontSize = 28
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    ssResult.Col = 2:
    
    If Mid(ssResult.Text, 1, 1) = "C" Then
        Printer.Print Tab(12); " 세포병리 검사보고"     '5
    ElseIf Mid(ssResult.Text, 1, 1) = "P" Then
        Printer.Print Tab(12); " 조직병리 검사보고"     '5
    ElseIf Mid(ssResult.Text, 1, 1) = "R" Then
        If LsitemGb = "C" Then
            Printer.Print Tab(12); " 세포병리 검사보고"     '5
        Else
            Printer.Print Tab(12); " 조직병리 검사보고"     '5
        End If
    End If
    
    
    Printer.Print
    
    Return

SUB_LINE_PRINT1:
    'box
    
'    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "바탕체"
    Printer.FontSize = 20
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    Printer.DrawWidth = 3
    
'    Printer.Line (8100, 1900)-(11300, 1900)
'    Printer.Line (11300, 1900)-(11300, 3500)
'    Printer.Line (8100, 3500)-(11300, 3500)
'    Printer.Line (8100, 1900)-(8100, 3500)
    
'    Printer.Line (8100, 2500)-(11300, 2500)
'    Printer.Line (11300, 2500)-(11300, 4600)
'    Printer.Line (8100, 4600)-(11300, 4600)
'    Printer.Line (8100, 2500)-(8100, 4600)
    
    Printer.Line (8100, 2900)-(11300, 2900)
    Printer.Line (11300, 2900)-(11300, 4600)
    Printer.Line (8100, 4600)-(11300, 4600)
    Printer.Line (8100, 2900)-(8100, 4600)
    
    Return

SUB_LINE_PRINT2:
    Printer.FontName = "바탕체"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    Printer.DrawWidth = 4
    Printer.Line (900, 4700)-(11300, 4700)
    Return

SUB_LINE_PRINT3:
    Printer.FontName = "바탕체"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    Printer.DrawWidth = 4
'    Printer.Line (900, 14600)-(11300, 14600)
    Printer.Line (900, 15600)-(11300, 15600)
    
    Printer.FontName = "돋움체"
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
'    Printer.CurrentX = 700:          Printer.CurrentY = 14500
    Printer.CurrentX = 700:          Printer.CurrentY = 15500
    Printer.Print
    Printer.Print Tab(10); "건양대학교 병원 "; Tab(62); "조직병리과"
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    Return

RESULT_PRINT:
    Dim LsStr       As String
    Dim LsChr       As String
    Dim LsLineStr   As String
    Dim LiLen       As Integer
    Dim LiPos       As Integer
    Dim y           As Integer
    Dim LF
    Dim CR
    
    LF = Chr(10)
    CR = Chr(13)
    
    If Trim(LsDescr) = "" Then Return
    
    LiLen = LenB(LsDescr)
    LsStr = LsDescr
    LiPos = 0
    LsLineStr = ""
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
'    Printer.Print
    
    Do
        LiPos = LiPos + 1
        LsChr = Mid(LsStr, LiPos, 1)
        If LsChr <> CR And LsChr <> LF Then
            LsLineStr = LsLineStr & LsChr
        End If
        
        If LsChr = LF Then
            LineCount = LineCount + 1

            Printer.Print Tab(10); LsLineStr
            LsLineStr = ""
            Select Case LineCount
                    Case 45
                              GoSub SUB_LINE_PRINT3
                              Printer.NewPage
                    
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                    
                    Case 127
                              GoSub SUB_LINE_PRINT3
                              Printer.NewPage
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                    Case 209
                              GoSub SUB_LINE_PRINT3
                              Printer.NewPage
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                    Case 291
                              GoSub SUB_LINE_PRINT3
                              Printer.NewPage
            
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                    Case 373
                              GoSub SUB_LINE_PRINT3
                              Printer.NewPage
            
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
            
                    Case 455
                              GoSub SUB_LINE_PRINT3
                              Printer.NewPage
            
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
                                Printer.Print
            End Select
        
        ElseIf LenB(LsLineStr) > 180 Then      '74  150  ' max cols
            GoSub SUB_LINEFEED_PROC
        End If
    
    Loop Until (LiPos > LiLen)
    
    Printer.Print Tab(10); LsLineStr
    
    Return
    

SUB_LINEFEED_PROC:
    
    Dim LsTempChr       As String
    
    LsTempChr = RightB(LsLineStr, 1)
    
    If Trim(LsTempChr) = "" Then
        Printer.Print Tab(10); LsLineStr
    ElseIf LsTempChr >= "A" And LsTempChr <= "z" Then
        If Trim(MidH(LsStr, LiPos + 1, 1)) = "" Then
            Printer.Print Tab(10); LsLineStr
        Else
            Printer.Print Tab(10); LsLineStr & "-"
        End If
    Else
        Printer.Print Tab(10); LsLineStr
    End If
    
    LineCount = LineCount + 1
    
    Select Case LineCount
            Case 45
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            
            Case 110
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            Case 175
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            Case 235
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
    
    End Select
    
    LsLineStr = ""
    
    Return


SUB_PRINT_COUNT_UPDATE:
    
    Dim LiPrintCnt      As Integer
    
    ssResult.Row = i:    ssResult.Col = 13
    LiPrintCnt = Val(ssResult.Text) + 1
    If LiPrintCnt > 9 Then LiPrintCnt = 9
    
    ssResult.Row = i:    ssResult.Col = 15
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG"
    strSQL = strSQL & " SET    speSLIDE  =  '" & LiPrintCnt & "'"
    strSQL = strSQL & " WHERE  ROWID =  '" & ssResult.Text & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return


SUB_PRINT_PATIENT_RESULT:

    Printer.FontName = "바탕체"
    Printer.FontSize = 14
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    LsDescr = LsEye
    GoSub RESULT_PRINT
            
    Return


SUB_PRINT_PATIENT_RESULT_EYE:
    On Error Resume Next
    
    Dim wCnt
    
    strSQL = ""
    strSQL = strSQL & "  SELECT  A.CLASS, A.DATEYY, A.SEQNUM, A.DIAGNO, A.DIAGADD, TO_CHAR(A.DIAGDATE, 'YYYY-MM-DD') DIAGDATE, B.DRNAME "
    strSQL = strSQL & "    FROM  TWANAT_DIAG   A, "
    strSQL = strSQL & "          TWBAS_DOCTOR  B  "
    strSQL = strSQL & "   WHERE  A.CHIEF    = B.DRCODE(+) "
    strSQL = strSQL & "     AND  A.PTNO     = '" & LsPtNo & "' "
    strSQL = strSQL & "     AND  A.GBRESULT >= '4' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Return
       
        wCnt = 0
        Do Until rs.EOF
            
            wCnt = wCnt + 1
            Printer.FontName = "바탕체"
            Printer.FontSize = 11
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            
            LsDescr = rs.Fields("DIAGNO").Value & ""
            
            If LsDescr <> "" Then
                Printer.Print Tab(10); rs.Fields("CLASS").Value & "-" & _
                                       rs.Fields("DATEYY").Value & "-" & _
                                       rs.Fields("SEQNUM").Value & ""
                
                GoSub RESULT_PRINT
                
                Printer.Print Tab(80); rs.Fields("DIAGDATE").Value & ""
                Printer.Print Tab(80); rs.Fields("DRNAME").Value & ""
            End If
            
            rs.MoveNext
        Loop
    
    AdoCloseSet rs

    Return
 
SUB_PRINT_NEWPAGE:
    Select Case LineCount
            Case 45
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            
            Case 127
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            Case 209
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            Case 291
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            Case 373
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
            Case 455
                      GoSub SUB_LINE_PRINT3
                      Printer.NewPage
    
    End Select
    Return
    

End Sub


Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = Len(txtClass.Text)

End Sub

Private Sub TXTCLASS_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    
    
    If KeyAscii > 0 Then
        TabCheck = TabCheck + 1
    End If
    
    If TabCheck = 1 Then
        TabCheck = 0
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtDateYY_GotFocus()
    txtDateYY.SelStart = 0
    txtDateYY.SelLength = Len(txtDateYY.Text)

End Sub

Private Sub txtDATEYY_KeyPress(KeyAscii As Integer)
    If KeyAscii > 0 And KeyAscii <> 8 Then
        TabCheck = TabCheck + 1
    Else
        TabCheck = TabCheck - 1
    End If
    
    If TabCheck = 4 Then
        TabCheck = 0
        SendKeys "{tab}"
    End If
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    

    SendKeys "{tab}"

End Sub


Private Sub txtSeqnum_GotFocus()
    txtSeqNum.SelStart = 0
    txtSeqNum.SelLength = Len(txtSeqNum.Text)

End Sub

Private Sub txtSeqnum_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
'    SendKeys "{tab}"
    cmdView.SetFocus

End Sub
