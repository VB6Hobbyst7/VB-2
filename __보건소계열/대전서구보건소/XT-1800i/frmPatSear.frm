VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatSear 
   BorderStyle     =   1  '단일 고정
   Caption         =   "환자조회"
   ClientHeight    =   8895
   ClientLeft      =   7440
   ClientTop       =   2250
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9615
   StartUpPosition =   2  '화면 가운데
   Begin MSComCtl2.MonthView monvCal 
      Height          =   2370
      Left            =   3480
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60293121
      CurrentDate     =   36878
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "WorkList"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   27
      Top             =   8340
      Width           =   1395
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1305
      Left            =   4230
      TabIndex        =   25
      Top             =   1710
      Visible         =   0   'False
      Width           =   1935
      _Version        =   393216
      _ExtentX        =   3413
      _ExtentY        =   2302
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmPatSear.frx":0000
   End
   Begin Threed.SSPanel sspOrder 
      Height          =   4665
      Left            =   600
      TabIndex        =   9
      Top             =   3270
      Visible         =   0   'False
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   8229
      _StockProps     =   15
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.TextBox txtReqDate 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   270
         TabIndex        =   26
         Top             =   4110
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   1290
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   1290
         TabIndex        =   18
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   1290
         TabIndex        =   17
         Top             =   1080
         Width           =   1395
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "확인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   330
         TabIndex        =   15
         Top             =   3300
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1590
         TabIndex        =   14
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   1290
         TabIndex        =   13
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   1290
         TabIndex        =   12
         Top             =   1920
         Width           =   915
      End
      Begin VB.CheckBox chkAllOrder 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3540
         TabIndex        =   11
         Top             =   240
         Width           =   225
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   270
         TabIndex        =   10
         Top             =   2370
         Visible         =   0   'False
         Width           =   1965
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   4395
         Left            =   3000
         TabIndex        =   16
         Top             =   150
         Width           =   4725
         _Version        =   393216
         _ExtentX        =   8334
         _ExtentY        =   7752
         _StockProps     =   64
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   100
         ScrollBars      =   2
         SpreadDesigner  =   "frmPatSear.frx":0274
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "접수번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "환자번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "성별"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "나이"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1980
         Width           =   1005
      End
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   2610
      Left            =   930
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   5280
      _Version        =   393216
      _ExtentX        =   9313
      _ExtentY        =   4604
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   8
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSear.frx":1403
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7590
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   8340
      Width           =   1395
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   3645
      Left            =   4740
      TabIndex        =   6
      Top             =   2310
      Visible         =   0   'False
      Width           =   2745
      _Version        =   393216
      _ExtentX        =   4842
      _ExtentY        =   6429
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmPatSear.frx":262D
   End
   Begin VB.CheckBox chkAll 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   660
      TabIndex        =   4
      Top             =   1170
      Width           =   165
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order 전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1770
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8310
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdDown 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   900
      Picture         =   "frmPatSear.frx":28A1
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8340
      Width           =   705
   End
   Begin VB.CommandButton cmdUp 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      Picture         =   "frmPatSear.frx":29D3
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8340
      Width           =   705
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   1050
      Width           =   9405
      _Version        =   393216
      _ExtentX        =   16589
      _ExtentY        =   12303
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   15
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSear.frx":2B02
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1005
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   9405
      _Version        =   65536
      _ExtentX        =   16589
      _ExtentY        =   1773
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   41
         Top             =   570
         Width           =   255
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4620
         TabIndex        =   40
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox dtpSDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   35
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox dtpEDate 
         Height          =   315
         Left            =   3360
         TabIndex        =   34
         Top             =   540
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7170
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   330
         Width           =   1005
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   8220
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   330
         Width           =   1005
      End
      Begin VB.ComboBox cboGubun 
         Height          =   315
         ItemData        =   "frmPatSear.frx":3FDF
         Left            =   1500
         List            =   "frmPatSear.frx":3FE1
         TabIndex        =   31
         Top             =   150
         Width           =   1965
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   6060
         TabIndex        =   30
         Top             =   540
         Width           =   795
      End
      Begin VB.TextBox txtStart 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   4950
         TabIndex        =   29
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   39
         Top             =   615
         Width           =   120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "처방일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   38
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "검사종류"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   37
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5820
         TabIndex        =   36
         Top             =   630
         Width           =   120
      End
   End
End
Attribute VB_Name = "frmPatSear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim iIndex As Integer
'
'Public gGubun As String
'Public glRow As Long
'Public gOCnt As Integer
'Public gCount As String
'
'Private Sub cboGubun_Click()
'    gGubun = ""
'
'    IsolateCode cboGubun.Text
'    gGubun = Trim(gCode)
'End Sub
'
'Private Sub chkAll_Click()
'    Dim iRow As Integer
'
'    If chkAll.Value = 1 Then
'        For iRow = 1 To vasList.DataRowCnt
'            vasList.Row = iRow
'            vasList.Col = 1
'
'            vasList.Value = 1
'        Next iRow
'    ElseIf chkAll.Value = 0 Then
'        For iRow = 1 To vasList.DataRowCnt
'            vasList.Row = iRow
'            vasList.Col = 1
'
'            vasList.Value = 0
'        Next iRow
'    End If
'End Sub
'
'Private Sub chkAllOrder_Click()
'    If chkAllOrder.Value = 1 Then
'        vasOrder.Row = -1
'        vasOrder.Col = 1
'        vasOrder.Value = 1
'    Else
'        vasOrder.Row = -1
'        vasOrder.Col = 1
'        vasOrder.Value = 0
'    End If
'End Sub
'
'Private Sub cmdCalendar_Click(Index As Integer)
'    iIndex = Index
'    If Index = 0 Then
'        monvCal.Left = 1620
'        monvCal.Top = 870
'        monvCal.Visible = True
'
'        monvCal.Value = dtpSDate.Text
'    ElseIf Index = 1 Then
'        monvCal.Left = 3480
'        monvCal.Top = 870
'        monvCal.Visible = True
'
'        monvCal.Value = dtpEDate.Text
'    End If
'    'monvCal.Visible = True
'End Sub
'
'Private Sub cmdClose_Click()
''    txtDate.Text = ""
''    txtPID.Text = ""
''    txtName.Text = ""
''    txtSex.Text = ""
''    txtAge.Text = ""
''
''    ClearSpread vasOrder
''
'    sspOrder.Visible = False
'End Sub
'
'Private Sub cmdDown_Click()
'    Dim lRow As Long
'
'    lRow = vasList.ActiveRow
'
'    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow + 1
'    vasActiveCell vasList, lRow + 1, 2
'    vasList_Click 2, lRow + 1
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdOK_Click()
''Local에 환자에 대한 검사항목 저장하기
'    Dim sCnt As String
'    Dim iRow As Integer
'
'    Dim sExamCode As String
'    Dim sSubCode As String
'    Dim sExamName As String
'
'    Dim sEquipCode As String
'    Dim sAge As String
'    Dim i As Integer
'
'    sCnt = ""
'
'    '2005/10/18 이상은
'    '처방일자에서 검사일자로 변경함
'    'txtDate = Format(txtDate.Text, "yyyymmdd")
'    txtDate = Format(frmInterface.txtToday.Text, "yyyymmdd")
'
'    SQL = " Select count(*) From pat_res " & vbCrLf & _
'          " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(txtPID) & "' " & vbCrLf & _
'          " And sendflag = 'O' "
'    res = db_select_Var(gLocal, SQL, sCnt)
'
'    If sCnt = "" Then
'        sCnt = "0"
'    End If
'
'    If txtAge.Text = "" Then
'        txtAge.Text = "0"
'    Else
'        sAge = Trim(txtAge.Text)
'    End If
'
'    If sCnt > 0 Then
'            SQL = " Delete From pat_res " & vbCrLf & _
'                  " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
'                  " And equipno = '" & gEquip & "' " & vbCrLf & _
'                  " And barcode = '" & Trim(txtPID.Text) & "' " & vbCrLf & _
'                  " And sendflag = 'O' "
'            res = SendQuery(gLocal, SQL)
'
'            If res = -1 Then
'                SaveQuery SQL
'            End If
'    End If
'
'    For iRow = 1 To vasOrder.DataRowCnt
'        vasOrder.Row = iRow
'        vasOrder.Col = 1
'
'        If vasOrder.Value = 1 Then
'            sExamCode = Trim(GetText(vasOrder, iRow, 2))
'            sSubCode = Trim(GetText(vasOrder, iRow, 3))
'            sExamName = Trim(GetText(vasOrder, iRow, 4))
'
'            sEquipCode = GetEquip_ExamCode(sExamCode)
'
'            SQL = ""
'            SQL = " Insert Into pat_res(examdate, equipno, barcode, equipcode,  " & vbCrLf & _
'                  " examcode, subcode, examname, pid, pname, psex, page, recedate, resdate, sendflag)  " & vbCrLf & _
'                  " Values ( '" & Trim(txtDate) & "', '" & gEquip & "', '" & Trim(txtPID.Text) & "' , '" & Trim(sEquipCode) & "', " & vbCrLf & _
'                  " '" & sExamCode & "', '" & sSubCode & "', '" & sExamName & "', '" & Trim(txtPID.Text) & "', " & vbCrLf & _
'                  " '" & Trim(txtName.Text) & "', '" & Trim(txtSex.Text) & "', " & sAge & ", " & vbCrLf & _
'                  " '" & Trim(txtReqDate) & "', '', 'O') "
'            res = SendQuery(gLocal, SQL)
'
'            If res = -1 Then
'                SaveQuery SQL
'            End If
'        ElseIf vasOrder.Value = 0 Then
'            If sCnt = 0 Then
'
'            ElseIf sCnt > 0 Then
'                sExamCode = Trim(GetText(vasOrder, iRow, 2))
'
'                SQL = " Delete From pat_res " & vbCrLf & _
'                      " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
'                      " And equipno = '" & gEquip & "' " & vbCrLf & _
'                      " And barcode = '" & Trim(txtPID.Text) & "' " & vbCrLf & _
'                      " And examcode = '" & sExamCode & "' "
'                res = SendQuery(gLocal, SQL)
'
'                If res = -1 Then
'                    SaveQuery SQL
'                End If
'            End If
'        End If
'    Next iRow
'
'    sspOrder.Visible = False
'End Sub
'
'
'Private Sub cmdOrder_Click()
''Order 만들고 전송하기
'
'    Dim sRetOrder As String     'Order Text넣을 변수
'    Dim sOrder As String
'
'    Dim i As Integer
'    Dim jOrder As Integer
'    Dim kOrder As Integer
'    Dim OrderRow As Integer
'    Dim iRow As Integer
'    Dim iiRow As Integer
'    Dim jRow As Integer
'    Dim jjRow As Integer
'    Dim kRow As Integer
'
'    Dim llRow As Long
'
'    Dim sBarcode As String      '검체번호
'    Dim sPID As String
'    Dim sReceNo As String       '접수번호
'    Dim sRackNo As String
'    Dim sPosNo As String
'    Dim sORDT As String         '접수일자
'    Dim sExamCode As String     '검사코드
'    Dim sEquipCode As String    '장비코드
'    Dim sOrderCode As String
'
'    Dim sDate As String
'    Dim sHead As String
'    Dim sPatient As String
'    Dim sOCnt As String
'    Dim sMsgEnd As String
'
'    Dim s  As String
'    Dim j As Integer
'    Dim k As Integer
'
'    Dim sCnt As String
'
'    On Error GoTo errorchk
'
'    gPatCnt = 0
'    gOCnt = 1
'
'    jjRow = 1
'
'    '메인화면에 WorkList 디스플레이 하기
'    cmdWorkList_Click
'
'    'Order 만들기================================================
'    ClearSpread frmInterface.vasOrderBuf
'
'    sRetOrder = ""
'
'    sBarcode = ""
'    sReceNo = ""
'
'    llRow = 1
'
'    For iRow = 1 To vasList.DataRowCnt
'        vasList.Row = iRow
'        vasList.Col = 1
'
'        If vasList.Value = 1 Then
'            sRackNo = Trim(GetText(vasList, iRow, 2))
'            sPosNo = Trim(GetText(vasList, iRow, 3))
'
'            sPID = Trim(GetText(vasList, iRow, 5))
'
'            '====================================================
'            ClearSpread vasCode
'
'            '검사코드, 검사항목코드 가져오기
'            SQL = " Select examcode " & vbCrLf & _
'                  " From pat_res " & vbCrLf & _
'                  " Where examdate = '" & Format(frmInterface.txtToday.Text, "yyyymmdd") & "' " & vbCrLf & _
'                  " And equipno = '" & gEquip & "' " & vbCrLf & _
'                  " And barcode = '" & Trim(sPID) & "' " & vbCrLf & _
'                  " And examcode in (" & gAllExam & ") "
'
'            res = db_select_Vas(gLocal, SQL, vasCode)
'            '====================================================
'            'Order 생성
'            sOCnt = 1
'
'            sOrderCode = ""
'
'            For i = 1 To vasCode.DataRowCnt
'                sExamCode = Trim(GetText(vasCode, i, 1))
'
'                '검사코드로 장비코드 불러오기
'                sEquipCode = GetEquip_ExamCode(sExamCode)
'                SetText vasCode, sEquipCode, i, 3
'
'                If sEquipCode <> "" Then
'                    sOCnt = sOCnt + 1
'
'                    If sOrderCode = "" Then
'                        sOrderCode = "^^^1.000000+" & sEquipCode & "+1"
'                    Else
'                        sOrderCode = sOrderCode & "\" & sEquipCode & "+1"
'                    End If
'
'                End If
'            Next i
'
'            sRetOrder = "O|" & sOCnt & "|" & sPID & "^" & sRackNo & "^" & sPosNo & "||" & sOrderCode & "|R||||||N||||4||||||||||O||||||" & chrCR & chrETX
'            sOrder = sRetOrder
'
'            SetText frmInterface.vasOrderBuf, sOrder, llRow, 1
'            SetText frmInterface.vasOrderBuf, sBarcode, llRow, 2
'
'            llRow = llRow + 1
'            sOCnt = sOCnt + 1
'
'            SetText vasList, CStr(sOCnt - 1), iRow, 10
'        End If
'    Next iRow
'    '============================================================
'
'
'    'Order 전송하기==============================================
'    gPatCnt = 0
'    gOCnt = 1
'
'    '2004/03/06 이상은 - Order 전송시 Order 스프레드 Clear
'    ClearSpread frmInterface.vasOrder
'
'    glRow = 1
'
'    If glRow = 1 Then
'        gCurMsgCnt = 1
'        'Head
'        sDate = Format(GetDateFull, "yyyymmddhhmmss")
'
'        'gHeader = "H|\^&||||||||||P|1" & chrCR & chrETX
'        gHeader = "H|\^&||||||||||||" & sDate & chrCR & chrETX
'        sHead = chrSTX & CCur(gCurMsgCnt) & gHeader & CheckSum(CStr(gCurMsgCnt) & gHeader) & chrCR & chrLF
'        Save_Raw_Data "[O]" & sHead
'
'        gCurMsgCnt = gCurMsgCnt + 1
'        If gCurMsgCnt = 8 Then
'            gCurMsgCnt = 0
'        End If
'
'        SetText frmInterface.vasOrder, sHead, glRow, 1
'    End If
'    glRow = glRow + 1
'
'    For iiRow = 1 To vasList.DataRowCnt
'        'Patient
'        vasList.Row = iiRow
'        vasList.Col = 1
'
'        If vasList.Value = 1 Then
'
'            gPatCnt = gPatCnt + 1
'            gPatient = "P|" & gPatCnt & "||||||" & chrCR & chrETX
'            sPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
'            SaveQuery1 "[O]" & sPatient
'
'            gCurMsgCnt = gCurMsgCnt + 1
'            If gCurMsgCnt = 8 Then
'                gCurMsgCnt = 0
'            End If
'
'            SetText frmInterface.vasOrder, sPatient, glRow, 1
'            glRow = glRow + 1
'
'            'Order
'            s = Trim(GetText(vasList, iiRow, 12))
'
'            For j = 1 To frmInterface.vasOrderBuf.DataRowCnt
''                For k = 1 To s
'                    sRetOrder = Trim(GetText(frmInterface.vasOrderBuf, gOCnt, 1))
'
'                    sOrder = chrSTX & CCur(gCurMsgCnt) & sRetOrder & CheckSum(CStr(gCurMsgCnt) & sRetOrder) & chrCR & chrLF
'                    SetText frmInterface.vasOrder, sOrder, glRow, 1
'                    SaveQuery1 "[O]" & sOrder
'
'                    gCurMsgCnt = gCurMsgCnt + 1
'                    If gCurMsgCnt = 8 Then
'                        gCurMsgCnt = 0
'                    End If
'
'                    glRow = glRow + 1
'                    gOCnt = gOCnt + 1
'                    sOCnt = sOCnt + 1
''                Next k
'                Exit For
'            Next j
'
'            gCount = frmInterface.vasOrder.DataRowCnt
'
'            jjRow = jjRow + 1
'        End If
'    Next iiRow
'
'
'    'Terminator
'    If gCurMsgCnt = "" Then
'        Exit Sub
'    End If
'
'    gMsgEnd = "L|1" & chrCR & chrETX
'    sMsgEnd = Chr(2) & CCur(gCurMsgCnt) & gMsgEnd & CheckSum(CStr(gCurMsgCnt) & gMsgEnd) & chrCR & chrLF
'    Save_Raw_Data "[O]" & sMsgEnd
'
'    gCurMsgCnt = gCurMsgCnt + 1
'    If gCurMsgCnt = 8 Then
'        gCurMsgCnt = 1
'    End If
'
'    glRow = frmInterface.vasOrder.DataRowCnt + 1
'    SetText frmInterface.vasOrder, sMsgEnd, glRow, 1
'    SetText frmInterface.vasOrder, chrEOT, glRow + 1, 1
'    'SaveData "[TX]" & chrEOT
'
'    gOrderRow = 0
'
'    '장비에 ENQ를 던져서, ACK를 받으면 Order 전송함
'    gPreMsg = chrENQ
'
'    frmInterface.MSComm1.Output = chrENQ
'    Save_Raw_Data "[Tx]" & chrENQ
'    Me.MousePointer = 11
'
'    Exit Sub
'
'errorchk:
'    MsgBox "전송중 에러가 있습니다. 확인"
'    Me.MousePointer = 0
'End Sub
'
'Private Sub cmdPrint_Click()
'Dim iRow As Integer
'Dim j As Integer
'
'Dim sCurDate As String
'Dim sSerDate As String
'Dim sHead As String
'Dim sFoot As String
'
'    ClearSpread vasPrint
'
'    j = 1
'
'    For iRow = 1 To vasList.DataRowCnt
'        vasList.Row = iRow
'        vasList.Col = 1
'
'        If vasList.Value = 1 Then
'            SetText vasPrint, Trim(GetText(vasList, iRow, 2)), j, 1     '검체번호
'            SetText vasPrint, Trim(GetText(vasList, iRow, 3)), j, 2     '환자번호
'            SetText vasPrint, Trim(GetText(vasList, iRow, 4)), j, 3     '환자이름
'
'            SetText vasPrint, Trim(GetText(vasList, iRow, 5)), j, 4     '성별
'            SetText vasPrint, Trim(GetText(vasList, iRow, 6)), j, 5     '나이
'            'SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 6     '주민번호
'            SetText vasPrint, Trim(GetText(vasList, iRow, 8)), j, 7     '처방일자
'            SetText vasPrint, "", j, 8
'
'            j = j + 1
'        End If
'    Next iRow
'
'    If vasPrint.DataRowCnt < 1 Then
'        MsgBox "출력할 자료가 없습니다.", , "알 림"
'        Exit Sub
'    End If
'
'    sCurDate = GetDateFull
'
'    sSerDate = Trim(dtpSDate.Text) & " - " & Trim(dtpEDate.Text)
'
'    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
'    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
'    vasPrint.PrintJobName = "VITROS Eci WorkList 출력"
'
'    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 울산인산병원 진단검사의학과"
'
'    vasPrint.PrintHeader = sHead
'    vasPrint.PrintFooter = sFoot
'
'    vasPrint.PrintMarginTop = 680
'    vasPrint.PrintMarginBottom = 680
''현재 SS가 비대칭으로 출력함
''    vaslist.PrintMarginLeft = 720
'    vasPrint.PrintMarginLeft = 0
'    vasPrint.PrintMarginRight = 0
'
'    vasPrint.PrintColor = True
'    vasPrint.PrintGrid = True
'
''Set printing range
'    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
'
'    vasPrint.PrintShadows = True
'
'    vasPrint.Action = 13 'SS_ACTION_PRINT
'End Sub
'
'Private Sub cmdSearch_Click()
'    Dim sSch1, sSch2 As String
'    Dim iRow As Integer
'    Dim i As Integer
'
'    ClearSpread vasList
'
'    'vasList.MaxRows = 100
'
'    '체크, Rack, Pos, '', 환자번호, 환자이름, 성별, 나이, 주민번호, 접수일자
'    sSch1 = Format(dtpSDate.Text, "yyyymmdd")
'    sSch2 = Format(dtpEDate.Text, "yyyymmdd")
'
'    '접수번호, 환자번호, 환자이름, 성별, 나이, 주민번호
'    Select Case gGubun
'    Case 1      '검진
'        If cn_Server_Flag = True Then
'            DisConnect_Server
'        Else
'            Connect_Server
'        End If
'
'        SQL = " Select a.Per_Gum_Num, b.ChartNo, a.Per_Name, '', '', a.Per_Ssn, a.Per_Gumjin_Date " & CR & _
'              " From Gumjin_Interface a, TB_PERSON b " & CR & _
'              " Where a.Per_Gumjin_Date >= '" & Trim(sSch1) & "' " & CR & _
'              " And a.Per_Gumjin_Date <= '" & Trim(sSch2) & "' " '& CR & _
'
'        If txtStart <> "" And txtEnd <> "" Then
'            SQL = SQL & " And a.Per_Gum_Num >= '" & Trim(txtStart) & "' " & CR & _
'                        " And a.Per_Gum_Num <= '" & Trim(txtEnd) & "' "
'        End If
'
'        SQL = SQL & CR & _
'              " And a.EdpsCode IN (" & gAllExam & ") " & CR & _
'              " And a.Result = '' " & CR & _
'              " And a.Per_Gum_Num = b.Per_Gum_Num " & CR & _
'              " And a.Per_Ssn = b.Per_Ssn " & CR & _
'              " Group By a.Per_Gum_Num, b.ChartNo, a.Per_Name, a.Per_Ssn, a.Per_Gumjin_Date "
'
'        res = db_select_Vas(gServer, SQL, vasList, , 4)
'
'    Case 2      '진료
'        If cn_Server_Flag_1 = True Then
'            DisConnect_Server_1
'        Else
'            Connect_Server_1
'        End If
'
''        SQL = " Select a.WaitSeqNo, a.ChartNo, b.SujinName, '', '',  b.PassNo, a.EnterDate " & CR & _
''              " From WaitPrsnp a, PewPrsnp b, JUN370_RESULTTB d " & CR & _
''              " Where a.EnterDate >= '" & Trim(sSch1) & "' " & CR & _
''              " And a.EnterDate <= '" & Trim(sSch2) & "' " & CR & _
''              " And a.JunDal = '370' " & CR & _
''              " And a.Status = '1' " & CR & _
''              " And d.Map2SeqNo IN (" & gAllOcsExam & ") " & CR & _
''              " And d.Status = '0' " & CR & _
''              " And a.ChartNo = b.ChartNo " & CR & _
''              " And a.WaitSeqNo = d.WaitSeqNo " & CR & _
''              " Group By a.WaitSeqNo, a.ChartNo, b.SujinName, b.PassNo, a.EnterDate "
'
'        SQL = " Select a.WaitSeqNo, a.ChartNo, b.SujinName, '', '',  b.PassNo, a.EnterDate " & CR & _
'              " From WaitPrsnp a, PewPrsnp b, JUN370_RESULTTB d " & CR & _
'              " Where a.EnterDate >= '" & Trim(sSch1) & "' " & CR & _
'              " And a.EnterDate <= '" & Trim(sSch2) & "' " '& CR & _
'
'        If txtStart <> "" And txtEnd <> "" Then
'            SQL = SQL & " And a.WaitSeqNo >= '" & Trim(txtStart) & "' " & CR & _
'                        " And a.WaitSeqNo <= '" & Trim(txtEnd) & "' "
'        End If
'
'        SQL = SQL & CR & _
'              " And a.JunDal = '370' " & CR & _
'              " And a.SujinPart <> '62' " & CR & _
'              " And d.Map2SeqNo IN (" & gAllOcsExam & ") " & CR & _
'              " And a.ChartNo = b.ChartNo " & CR & _
'              " And a.WaitSeqNo = d.WaitSeqNo " & CR & _
'              " Group By a.WaitSeqNo, a.ChartNo, b.SujinName, b.PassNo, a.EnterDate "
'        res = db_select_Vas(gServer_1, SQL, vasList, , 4)
'    End Select
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    vasSort vasList, 4
'
'    For iRow = 1 To vasList.DataRowCnt
'        CalAgeSex Trim(GetText(vasList, iRow, 9)), frmInterface.txtToday
'
'        SetText vasList, gPatGen.Sex, iRow, 7
'        SetText vasList, gPatGen.Age, iRow, 8
'    Next iRow
'
'    vasList.MaxRows = vasList.DataRowCnt + 1
'    vasList.RowHeight(-1) = 13
'End Sub
'
'Private Sub cmdUp_Click()
'    Dim lRow As Long
'
'    lRow = vasList.ActiveRow
'
'    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow - 1
'    vasActiveCell vasList, lRow - 1, 2
'    vasList_Click 2, lRow - 1
'End Sub
'
'Private Sub cmdWorkList_Click()
'    Dim lRow As Long
'    Dim lCol As Long
'    Dim iRow As Integer
'    Dim lDestRow As Long
'
'    lDestRow = frmInterface.vasID.DataRowCnt + 1
'
'    If lDestRow > frmInterface.vasID.DataRowCnt Then
'        frmInterface.vasID.MaxRows = lDestRow
'    End If
'
'    For lRow = 1 To vasList.DataRowCnt
'        vasList.Row = lRow
'        vasList.Col = 1
'        If vasList.Value = 1 Then
'            If frmInterface.vasID.DataRowCnt >= lRow Then
'                For iRow = 1 To frmInterface.vasID.DataRowCnt
'                    If Trim(GetText(vasList, lRow, 4)) = Trim(GetText(frmInterface.vasID, iRow, 4)) Then
'
'                    End If
'                Next iRow
'            Else
'                gWorkFlag = gWorkFlag + 1
'
'                For lCol = 2 To 11
'                    If lCol = 2 Or lCol = 3 Then    'Rack, Pos
'                    ElseIf lCol = 4 Or lCol = 5 Or lCol = 6 Then    '접수번호, 환자번호,환자이름
'                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol
'                    ElseIf lCol = 7 Or lCol = 8 Then    '성별,나이
'                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 1
'                    ElseIf lCol = 9 Then                '주민번호
'                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 7
'                    ElseIf lCol = 10 Then               '접수일자
'                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 14
'                    End If
'                Next lCol
'
'                SetText frmInterface.vasID, Left(cboGubun.Text, 1), lDestRow, 15
'
'                'Local에 해당환자 정보 저장하기
'
'
'                lDestRow = lDestRow + 1
'
'                If lDestRow > frmInterface.vasID.DataRowCnt Then
'                    frmInterface.vasID.MaxRows = lDestRow
'                End If
'            End If
'        End If
'    Next lRow
'
'    chkAll.Value = 0
'    chkAll_Click
'
'    'Unload Me
'End Sub
'
'Private Sub Form_Activate()
'    vasActiveCell vasList, 1, 2
'End Sub
'
'Private Sub Form_Load()
'    dtpSDate.Text = Format(CDate(GetDateFull), "yyyy-mm-dd")
'    dtpEDate.Text = dtpSDate.Text
'
'    '검사종류
'    With cboGubun
'        .AddItem "1 검진"
'        .AddItem "2 진료"
'    End With
'
'    cboGubun.AddItem " ", 0
'    cboGubun.ListIndex = 1
'
'    ClearSpread vasList
'
'    chkAll.Value = 0
'End Sub
'
'Private Sub monvCal_DateClick(ByVal DateClicked As Date)
'    If iIndex = 0 Then
'        dtpSDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
'    Else
'        dtpEDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
'    End If
'    monvCal.Visible = False
'End Sub
'
'Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
'    If Row = 0 Then
'        Select Case Col
'        Case 4      '접수번호
'            vasSort vasList, 4, 6
'        Case 5      '챠트번호
'            vasSort vasList, 5, 6
'        Case 6      '성명
'            vasSort vasList, 6, 4
'        Case 10     '접수일자
'            vasSort vasList, 10, 4
'        End Select
'    End If
'
'    If Row < 0 Or Row > vasList.DataRowCnt Then
'        cmdUp.Enabled = False
'        cmdDown.Enabled = False
'    End If
'
'    If Row = 1 Then
'        cmdUp.Enabled = False
'        cmdDown.Enabled = True
'    ElseIf Row = vasList.DataRowCnt Then
'        cmdUp.Enabled = True
'        cmdDown.Enabled = False
'    Else
'        cmdUp.Enabled = True
'        cmdDown.Enabled = True
'    End If
'End Sub
'
'Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
''    Dim lRow, lCol As Long
''    Dim lDestRow As Long
''
''    lDestRow = Form_Main.vasExam.DataRowCnt + 1
''
''    lRow = vasList.ActiveRow
''
''    For lCol = 2 To 8
''        If lCol = 8 Then        '처방일자
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 8)), lDestRow, 12
''        ElseIf lCol = 2 Then    '검체번호
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 2)), lDestRow, 2
''        Else
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 3
''        End If
''    Next lCol
'
''    Unload Me
'
''===================================================================
''2004/08/03 이상은 - 환자 더블클릭시 상세 검사항목 디스플레이 되도록
'    Dim sGubun As String
'    Dim sCnt As String
'    Dim sExamCode As String
'    Dim sEquipCode As String
'    Dim iRow  As Integer
'
'    'Clear
'    txtDate = ""
'    txtReqDate = ""
'
'    txtNo = ""
'    txtPID = ""
'    txtName = ""
'    txtSex = ""
'    txtAge = ""
'
'    sGubun = ""
'    IsolateCode cboGubun.Text
'    sGubun = Trim(gCode)
'
'    '2006.01.31 이상은
'    'txtDate = GetText(vasList, Row, 10)
'    txtDate = Format(frmInterface.txtToday.Text, "yyyymmdd")
'
'    txtReqDate = Trim(GetText(vasList, Row, 10))    '처방일자
'
'
'    txtNo = Trim(GetText(vasList, Row, 4))          '접수번호
'    txtPID = Trim(GetText(vasList, Row, 5))         '등록번호
'    txtName = Trim(GetText(vasList, Row, 6))        '환자이름
'
'    txtSex = Trim(GetText(vasList, Row, 7))
'    txtAge = Trim(GetText(vasList, Row, 8))
'
'    chkAllOrder.Value = 0
'
'    ClearSpread vasOrder
'
'    '검사코드 가져오기
'    Select Case sGubun
'    Case 1
'        If cn_Server_Flag = True Then
'            DisConnect_Server
'        Else
'            Connect_Server
'        End If
'
'        SQL = " Select '', EdpsCode " & vbCrLf & _
'              " From Gumjin_Interface " & vbCrLf & _
'              " Where Per_Gumjin_Date = '" & Trim(txtReqDate) & "' " & CR & _
'              " And Per_Gum_Num = '" & txtNo & "' " & vbCrLf & _
'              " And EdpsCode in (" & gAllExam & ") "
'
'        res = db_select_Vas(gServer, SQL, vasOrder)
'
'    Case 2
'        If cn_Server_Flag_1 = True Then
'            DisConnect_Server_1
'        Else
'            Connect_Server_1
'        End If
'
'        SQL = " Select '', Map2SeqNo " & vbCrLf & _
'              " From JUN370_RESULTTB " & vbCrLf & _
'              " Where WaitSeqNo = '" & Trim(txtNo) & "' " & vbCrLf & _
'              " And Map2SeqNo in (" & gAllOcsExam & ") "
'
'        res = db_select_Vas(gServer_1, SQL, vasOrder)
'    End Select
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    vasOrder.MaxRows = vasOrder.DataRowCnt
'
'    sspOrder.Visible = True
'
'    For iRow = 1 To vasOrder.DataRowCnt
'        '장비구분, OCS코드, 검사명 디스플레이
'        Select Case sGubun
'        Case "1"
'            SQL = " Select subcode, ocscode, examname From equipexam " & CR & _
'                  " Where equipno = '" & gEquip & "' " & CR & _
'                  " And examcode = '" & Trim(GetText(vasOrder, iRow, 2)) & "' "
'        Case "2"
'            SQL = " Select examcode, ocscode, examname From equipexam " & CR & _
'                  " Where equipno = '" & gEquip & "' " & CR & _
'                  " And ocscode = '" & Trim(GetText(vasOrder, iRow, 2)) & "' "
'        End Select
'
'        res = db_select_Col(gLocal, SQL)
'
'        If gReadBuf(0) <> "" Then
'            vasOrder.SetText 3, iRow, Trim(gReadBuf(0))
'            vasOrder.SetText 4, iRow, Trim(gReadBuf(1))
'            vasOrder.SetText 5, iRow, Trim(gReadBuf(2))
'        End If
'
'        '로컬에 해당 검사항목 처리됐는지 확인함
'        SQL = " Select examcode " & CR & _
'              " From pat_res " & CR & _
'              " Where examdate = '" & txtDate & "' " & CR & _
'              " And equipno = '" & gEquip & "' " & CR & _
'              " And barcode = '" & Trim(txtNo.Text) & "' " & CR & _
'              " And examcode = '" & Trim(GetText(vasOrder, iRow, 2)) & "' " & CR & _
'              " And sendflag = 'O' "
'        res = db_select_Col(gLocal, SQL)
'
'        If res = 1 Then
'            vasOrder.Row = iRow
'            vasOrder.Col = 1
'            vasOrder.Value = 1
'        End If
'    Next iRow
'End Sub
'
'Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Integer
'    Dim iCol As Integer
'
'    Dim jRow As Integer
'    Dim iCnt As Integer
'    Dim sRack As String
'
'    iRow = vasList.ActiveRow
'    iCol = vasList.ActiveCol
'
'    If KeyCode = vbKeyReturn Then
'        If iCol = 2 Then    'Rack
'            If Trim(GetText(vasList, iRow, iCol)) <> "" Then
'                sRack = Trim(GetText(vasList, iRow, 2))
'
'                iCnt = -1
'                For jRow = 1 To vasList.DataRowCnt
'                    vasList.Row = jRow
'                    vasList.Col = 1
'
'                    If vasList.Value = 1 Then
'                        If iCnt = 9 Then
'                            sRack = CInt(sRack) + 1
'                            iCnt = -1
'                        End If
'
'                        iCnt = iCnt + 1
'                        SetText vasList, sRack, jRow, 2
'                        SetText vasList, CStr(iCnt), jRow, 3
'                    End If
'                Next jRow
'            End If
'        End If
'    End If
'End Sub
