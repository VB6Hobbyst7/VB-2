VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatSear 
   BorderStyle     =   1  '단일 고정
   Caption         =   " 환자조회"
   ClientHeight    =   8400
   ClientLeft      =   7440
   ClientTop       =   2250
   ClientWidth     =   9090
   Icon            =   "frmPatSear.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9090
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4800
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "오더 전송"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6210
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   7800
      Width           =   1335
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3435
      TabIndex        =   42
      Text            =   "1"
      Top             =   7995
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSComCtl2.MonthView monvCal 
      Height          =   2220
      Left            =   3930
      TabIndex        =   10
      Top             =   1650
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   21364737
      CurrentDate     =   36878
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   2610
      Left            =   930
      TabIndex        =   18
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
      SpreadDesigner  =   "frmPatSear.frx":08CA
   End
   Begin Threed.SSPanel sspOrder 
      Height          =   4005
      Left            =   1020
      TabIndex        =   19
      Top             =   2655
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7064
      _Version        =   131072
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
      Begin VB.TextBox txtNo 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   29
         Top             =   330
         Width           =   1395
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   28
         Top             =   750
         Width           =   1395
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   27
         Top             =   1170
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
         TabIndex        =   25
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
         Left            =   1620
         TabIndex        =   24
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   23
         Top             =   1590
         Width           =   915
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   22
         Top             =   2010
         Width           =   915
      End
      Begin VB.CheckBox chkAllOrder 
         Caption         =   "Check1"
         Height          =   345
         Left            =   3540
         TabIndex        =   21
         Top             =   390
         Width           =   225
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   20
         Top             =   2940
         Visible         =   0   'False
         Width           =   1965
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   3555
         Left            =   3000
         TabIndex        =   26
         Top             =   240
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   6271
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
         MaxCols         =   10
         MaxRows         =   100
         ScrollBars      =   2
         SpreadDesigner  =   "frmPatSear.frx":1AA1
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
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
         TabIndex        =   34
         Top             =   390
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
         TabIndex        =   33
         Top             =   810
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
         TabIndex        =   32
         Top             =   1230
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
         TabIndex        =   31
         Top             =   1650
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
         TabIndex        =   30
         Top             =   2070
         Width           =   1005
      End
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   3645
      Left            =   4740
      TabIndex        =   16
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
      SpreadDesigner  =   "frmPatSear.frx":2AFB
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
      Height          =   525
      Left            =   4635
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CheckBox chkAll 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   750
      TabIndex        =   9
      Top             =   1215
      Width           =   165
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7605
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   7800
      Width           =   1320
   End
   Begin VB.CommandButton cmdDown 
      Height          =   525
      Left            =   900
      Picture         =   "frmPatSear.frx":2D1C
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   7830
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdUp 
      Height          =   525
      Left            =   150
      Picture         =   "frmPatSear.frx":2E4E
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   7830
      Visible         =   0   'False
      Width           =   705
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   750
      Left            =   180
      TabIndex        =   5
      Top             =   30
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   1323
      _Version        =   131072
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
      Begin VB.ComboBox cboGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   40
         Top             =   990
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.OptionButton optState 
         BackColor       =   &H00E0E0E0&
         Caption         =   "모두"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3390
         TabIndex        =   38
         Top             =   690
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.OptionButton optState 
         BackColor       =   &H00E0E0E0&
         Caption         =   "결과"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2430
         TabIndex        =   37
         Top             =   690
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.OptionButton optState 
         BackColor       =   &H00E0E0E0&
         Caption         =   "접수"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1470
         TabIndex        =   36
         Top             =   690
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   945
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
         Left            =   4680
         TabIndex        =   12
         Top             =   240
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
         Index           =   0
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox dtpEDate 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3420
         TabIndex        =   13
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox dtpSDate 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   210
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7050
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         TabIndex        =   39
         Top             =   1050
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "진행상태"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   35
         Top             =   690
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방일자"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   8
         Top             =   285
         Width           =   780
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
         Left            =   3150
         TabIndex        =   7
         Top             =   285
         Width           =   120
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6945
      Left            =   150
      TabIndex        =   4
      Top             =   780
      Width           =   8775
      _Version        =   393216
      _ExtentX        =   15478
      _ExtentY        =   12250
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
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
      MaxCols         =   12
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15461355
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSear.frx":2F7D
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "※ Start No"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1725
      TabIndex        =   43
      Top             =   8025
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "결과완료 : 빨간색, 미완료 : 검정색"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   180
      TabIndex        =   41
      Top             =   7440
      Visible         =   0   'False
      Width           =   3675
   End
End
Attribute VB_Name = "frmPatSear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iIndex As Integer

Public glRow As Long
Public gOCnt As Integer
Public gCount As String

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

Private Sub chkAllOrder_Click()
    If chkAllOrder.Value = 1 Then
        vasOrder.Row = -1
        vasOrder.Col = 1
        vasOrder.Value = 1
    Else
        vasOrder.Row = -1
        vasOrder.Col = 1
        vasOrder.Value = 0
    End If
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    iIndex = Index
    If Index = 0 Then
        monvCal.Left = 1680
        monvCal.Top = 570
        monvCal.Visible = True
        
        monvCal.Value = dtpSDate.Text
    ElseIf Index = 1 Then
        monvCal.Left = 3600
        monvCal.Top = 570
        monvCal.Visible = True
        
        monvCal.Value = dtpEDate.Text
    End If
    'monvCal.Visible = True
End Sub

Private Sub cmdClose_Click()
    txtDate.Text = ""
    txtPID.Text = ""
    txtName.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    
    ClearSpread vasOrder
    
    sspOrder.Visible = False
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow + 1
    vasActiveCell vasList, lRow + 1, 2
    vasList_Click 2, lRow + 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'Local에 환자에 대한 검사항목 저장하기
Dim sCnt As String
Dim iRow As Integer
Dim sExamCode As String
Dim sEquipCode As String
Dim sAge As String
Dim i As Integer

    sCnt = ""
    
    SQL = " Select count(*) From pat_res " & vbCrLf & _
          " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(txtNo) & "' " & vbCrLf & _
          " And sendflag = 'O' "
    res = db_select_Var(gLocal, SQL, sCnt)
    
    If sCnt = "" Then
        sCnt = "0"
    End If
    
    If txtAge.Text = "" Then
        txtAge.Text = "0"
    Else
        sAge = Trim(txtAge.Text)
    End If
    
    If sCnt > 0 Then
            SQL = " Delete From pat_res " & vbCrLf & _
                  " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
                  " And equipno = '" & gEquip & "' " & vbCrLf & _
                  " And barcode = '" & Trim(txtNo.Text) & "' " & vbCrLf & _
                  " And sendflag = 'O' "
            res = SendQuery(gLocal, SQL)
            
            If res = -1 Then
                SaveQuery SQL
            End If
    End If
    
    For iRow = 1 To vasOrder.DataRowCnt
        vasOrder.Row = iRow
        vasOrder.Col = 1
        
        If vasOrder.Value = 1 Then
            sExamCode = Trim(GetText(vasOrder, iRow, 2))
            sEquipCode = GetEquip_ExamCode(sExamCode)

            SQL = " Insert Into pat_res(examdate, equipno, barcode, equipcode,  " & vbCrLf & _
                  " examcode, pid, pname, psex, page, resdate, sendflag)  " & vbCrLf & _
                  " Values ( '" & Trim(txtDate) & "', '" & gEquip & "', '" & Trim(txtNo.Text) & "' , '" & Trim(sEquipCode) & "', " & vbCrLf & _
                  " '" & sExamCode & "', '" & Trim(txtPID.Text) & "', " & vbCrLf & _
                  " '" & Trim(txtName.Text) & "', '" & Trim(txtSex.Text) & "', " & sAge & ", " & vbCrLf & _
                  " '" & Trim(GetDateFull) & "', 'O') "
            res = SendQuery(gLocal, SQL)
            
            If res = -1 Then
                SaveQuery SQL
            End If
        ElseIf vasOrder.Value = 0 Then
            If sCnt = 0 Then
            
            ElseIf sCnt > 0 Then
                sExamCode = Trim(GetText(vasOrder, iRow, 2))
                
                SQL = " Delete From pat_res " & vbCrLf & _
                      " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(txtNo.Text) & "' " & vbCrLf & _
                      " And examcode = '" & sExamCode & "' "
                res = SendQuery(gLocal, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                End If
            End If
        End If
    Next iRow
    
    sspOrder.Visible = False
End Sub

Private Sub CmdOrder_Click()
'Order 만들고 전송하기

    Dim sRetOrder As String     'Order Text넣을 변수
    Dim sOrder As String
    
    Dim i As Integer
    Dim jOrder As Integer
    Dim kOrder As Integer
    Dim OrderRow As Integer
    Dim iRow As Integer
    Dim iiRow As Integer
    Dim jRow As Integer
    Dim jjRow As Integer
    Dim kRow As Integer
    
    Dim llRow As Long
    
    Dim sGubun As String        '검사종류
    Dim sBarCode As String      '검체번호
    Dim sPID As String
    Dim sReceNo As String       '접수번호
    Dim sRackNo As String
    Dim sPosNo As String
    Dim sORDT As String         '접수일자
    Dim sExamCode As String     '검사코드
    Dim sEquipCode As String    '장비코드
    Dim sOrderCode As String
    
    Dim sDate As String
    Dim sHead As String
    Dim sPatient As String
    Dim sOCnt As String
    Dim sMsgEnd As String
    
    Dim s  As String
    Dim j As Integer
    Dim k As Integer
    
    Dim sCnt As String
    
    On Error GoTo errorchk
    
    IsolateCode cboGubun.Text
    sGubun = Trim(gCode)
    
    gPatCnt = 0
    gOCnt = 1
    
    jjRow = 1
    
    cmdWorkList_Click
    
    'Order 만들기================================================
    ClearSpread frmInterface.vasOrderBuf
    
    sRetOrder = ""
    
    sBarCode = ""
    sReceNo = ""
    
    llRow = 1
    
    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.Col = 1
        
        If vasList.Value = 1 Then
            sBarCode = Trim(GetText(vasList, iRow, 4))
            sRackNo = Trim(GetText(vasList, iRow, 2))
            sPosNo = Trim(GetText(vasList, iRow, 3))
            sReceNo = Trim(GetText(vasList, iRow, 11))
            sPID = Trim(GetText(vasList, iRow, 5))
            
            '====================================================
            '검사코드, 검사항목코드 가져오기
            ClearSpread vasCode
            
            Select Case sGubun
            Case "1"
                SQL = " Select ExamCode From ExamRes" & vbCrLf & _
                      " Where HID = '116' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And SpecimenID = '" & Trim(sBarCode) & "'" & vbCrLf & _
                      " And ExamCode in (" & gAllExam & ") "
            Case "2"
                SQL = " Select ExamCode From ExamRes" & vbCrLf & _
                      " Where HID = '116' " & vbCrLf & _
                      " And PID = '" & sPID & "' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And ExamCode in (" & gAllExam & ") "
            End Select
            
            res = db_select_Vas(gServer, SQL, vasCode)
            '====================================================
            'Order 생성
            sOCnt = 1
            
            sOrderCode = ""
        
            For i = 1 To vasCode.DataRowCnt
                sExamCode = Trim(GetText(vasCode, i, 1))
        
                '검사코드로 장비코드 불러오기
                sEquipCode = GetEquip_ExamCode(sExamCode)
                SetText vasCode, sEquipCode, i, 3
                
                If sEquipCode <> "" Then
'                    If i = 1 Then
'                        'sRetOrder = "O|" & sOCnt & "|" & argNo & "||" & "^^^1.000000+" & sEquipCode & "+1" & "|R||||||N||||4||||||||||O||||||" & chrCR & chrETX
'                        sRetOrder = "O|" & sOCnt & "|" & argNo & "||" & "^^^1.000000+" & sEquipCode & "+1" & "|R||||||N||||4||||||||||O||||||" & chrCR & chrETX
'                    Else
'                        'sRetOrder = "O|" & sOCnt & "|" & argNo & "^06" & "^0" & "||" & "^^^1.000000+" & sEquipCode & "+1" & "|R||||||A||||4||||||||||F||||||" & chrCR & chrETX
'                        sRetOrder = "O|" & sOCnt & "|" & argNo & "||" & "^^^1.000000+" & sEquipCode & "+1" & "|R||||||A||||4||||||||||O||||||" & chrCR & chrETX
'                    End If
'
'                    sOrder = sRetOrder
'
'                    SetText frminterface.vasOrderBuf, sOrder, llRow, 1
'                    SetText frminterface.vasOrderBuf, argNo, llRow, 2
'
'                    llRow = llRow + 1
                    sOCnt = sOCnt + 1
                    
                    If sOrderCode = "" Then
                        sOrderCode = "^^^1.000000+" & sEquipCode & "+1"
                    Else
                        sOrderCode = sOrderCode & "\" & sEquipCode & "+1"
                    End If
                    
                End If
            Next i
        
            sRetOrder = "O|" & sOCnt & "|" & sBarCode & "^" & sRackNo & "^" & sPosNo & "||" & sOrderCode & "|R||||||N||||4||||||||||O||||||" & chrCR & chrETX
            sOrder = sRetOrder
        
            SetText frmInterface.vasOrderBuf, sOrder, llRow, 1
            SetText frmInterface.vasOrderBuf, sBarCode, llRow, 2
        
            llRow = llRow + 1
            sOCnt = sOCnt + 1
        
            SetText vasList, CStr(sOCnt - 1), iRow, 10
        End If
    Next iRow
    '============================================================
    
    
    'Order 전송하기==============================================
    gPatCnt = 0
    gOCnt = 1

    '2004/03/06 이상은 - Order 전송시 Order 스프레드 Clear
    ClearSpread frmInterface.vasOrder

    glRow = 1

    If glRow = 1 Then
        gCurMsgCnt = 1
        'Head
        sDate = Format(GetDateFull, "yyyymmddhhmmss")
        
        'gHeader = "H|\^&||||||||||P|1" & chrCR & chrETX
        gHeader = "H|\^&||||||||||||" & sDate & chrCR & chrETX
        sHead = chrSTX & CCur(gCurMsgCnt) & gHeader & CheckSum(CStr(gCurMsgCnt) & gHeader) & chrCR & chrLF
        SaveQuery1 "[O]" & sHead
        
        gCurMsgCnt = gCurMsgCnt + 1
        If gCurMsgCnt = 8 Then
            gCurMsgCnt = 0
        End If

        SetText frmInterface.vasOrder, sHead, glRow, 1
    End If
    glRow = glRow + 1
        
    For iiRow = 1 To vasList.DataRowCnt
        'Patient
        vasList.Row = iiRow
        vasList.Col = 1
        
        If vasList.Value = 1 Then
        
            gPatCnt = gPatCnt + 1
            gPatient = "P|" & gPatCnt & "||||||" & chrCR & chrETX
            sPatient = chrSTX & CCur(gCurMsgCnt) & gPatient & CheckSum(CStr(gCurMsgCnt) & gPatient) & chrCR & chrLF
            SaveQuery1 "[O]" & sPatient
            
            gCurMsgCnt = gCurMsgCnt + 1
            If gCurMsgCnt = 8 Then
                gCurMsgCnt = 0
            End If
        
            SetText frmInterface.vasOrder, sPatient, glRow, 1
            glRow = glRow + 1
            
            'Order
            s = Trim(GetText(vasList, iiRow, 12))
            
            For j = 1 To frmInterface.vasOrderBuf.DataRowCnt
'                For k = 1 To s
                    sRetOrder = Trim(GetText(frmInterface.vasOrderBuf, gOCnt, 1))
                
                    sOrder = chrSTX & CCur(gCurMsgCnt) & sRetOrder & CheckSum(CStr(gCurMsgCnt) & sRetOrder) & chrCR & chrLF
                    SetText frmInterface.vasOrder, sOrder, glRow, 1
                    SaveQuery1 "[O]" & sOrder
                    
                    gCurMsgCnt = gCurMsgCnt + 1
                    If gCurMsgCnt = 8 Then
                        gCurMsgCnt = 0
                    End If
        
                    glRow = glRow + 1
                    gOCnt = gOCnt + 1
                    sOCnt = sOCnt + 1
'                Next k
                Exit For
            Next j
            
            gCount = frmInterface.vasOrder.DataRowCnt
            
            jjRow = jjRow + 1
        End If
    Next iiRow
    
    
    'Terminator
    If gCurMsgCnt = "" Then
        Exit Sub
    End If
    
    gMsgEnd = "L|1" & chrCR & chrETX
    sMsgEnd = Chr(2) & CCur(gCurMsgCnt) & gMsgEnd & CheckSum(CStr(gCurMsgCnt) & gMsgEnd) & chrCR & chrLF
    SaveQuery1 "[O]" & sMsgEnd
    
    gCurMsgCnt = gCurMsgCnt + 1
    If gCurMsgCnt = 8 Then
        gCurMsgCnt = 1
    End If
    
    glRow = frmInterface.vasOrder.DataRowCnt + 1
    SetText frmInterface.vasOrder, sMsgEnd, glRow, 1
    SetText frmInterface.vasOrder, chrEOT, glRow + 1, 1
    'SaveData "[TX]" & chrEOT
    
    gOrderRow = 0
    
    gPreMsg = chrENQ
    
    frmInterface.MSComm1.Output = chrENQ
    SaveQuery1 "[Tx]" & chrENQ
    Me.MousePointer = 11
    
    Exit Sub
    
errorchk:
    MsgBox "전송중 에러가 있습니다. 확인"
    Me.MousePointer = 0
    
End Sub

Private Sub cmdPrint_Click()
Dim iRow As Integer
Dim j As Integer

Dim sCurDate As String
Dim sSerDate As String
Dim sHead As String
Dim sFoot As String
    
    ClearSpread vasPrint

    j = 1

    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.Col = 1

        If vasList.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasList, iRow, 4)), j, 1     '검체번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 5)), j, 2     '환자번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 6)), j, 3     '환자이름

            SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 4     '성별
            SetText vasPrint, Trim(GetText(vasList, iRow, 8)), j, 5     '나이
            'SetText vasPrint, Trim(GetText(vasList, iRow, 9)), j, 6     '주민번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 10)), j, 7     '처방일자
            SetText vasPrint, "", j, 8
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    sCurDate = GetDateFull
    
    sSerDate = Trim(dtpSDate.Text) & " - " & Trim(dtpEDate.Text)
    
    '2004/08/11 이상은 - 세로방향에서 가로방향으로 수정
    vasPrint.PrintOrientation = 2   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "Hitachi7080 WorkList 출력"
    

'    sHead = "/fn""궁서체"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "▣ AXSYM WorkList ▣" & "/n/n " & _
'                "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "처방일자 : " & sSerDate & "/n/n" & _

    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 영천마야병원 진단검사의학과"
    
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
    
'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
End Sub

Private Sub cmdSearch_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim i, X As Long
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim FilNum
    Dim TxtString As String
    Dim TxtRece As String
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    Dim PExamname As String
    Dim PEquipCode As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim TxtPat As String
    Dim TestNum, IOGubun As String
    Dim FindFile As String
    Dim StartDate As String
    Dim EndDate As String
    Dim strSrcfile  As String
    Dim strDestFile As String

    Screen.MousePointer = 11
    
    sSch1 = Format(dtpSDate.Text, "yyyymmdd") '& "0000"
    sSch2 = Format(dtpEDate.Text, "yyyymmdd") '& "2400"
    
    ClearSpread vasList
    XmlTxt = ""
    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
    
    If FindFile <> "" Then
        FilNum = FreeFile
        Open "C:\UBCare\SINAI\IF\ExamIF_In.xml" For Input As FilNum
'        Open "\\192.168.0.47\C\UBCare\SINAI\IF\ExamIF_In.xml" For Input As FilNum
        Do While Not EOF(FilNum)
            'Input #FilNum, TxtString
            Line Input #FilNum, TxtString ' 변수로 데이터 행을 읽어들입니다.
            
            XmlTxt = XmlTxt & TxtString
        Loop
        Close #FilNum
        
        XmlTxtHead = ""
        XmlTxtTail = ""
        TxtPat = TxtString
        i = InStr(1, TxtPat, "<검사>")
        X = InStr(1, XmlTxt, "<UBCare검사정보>")
        XmlTxtHead = Mid(XmlTxt, 1, X + 11)
        X = InStr(1, XmlTxt, "</UBCare검사정보>")
        XmlTxtTail = Mid(XmlTxt, X, 13)
        While i > 0
            If InStr(1, TxtPat, "<검사>") Then
            '환자별로 text 를 구별
                i = InStr(1, TxtPat, "<검사>")
                
                TxtPat = Mid(TxtPat, i)
                i = InStr(1, TxtPat, "</검사>")
                TxtRece = Mid(TxtPat, 1, i + 4)
                TxtPat = Mid(TxtPat, i + 5)
                
            '차트번호, 환자이름, 주민번호, 내원번호, 검사코드 구분,의뢰일
            '차트번호
                i = InStr(1, TxtRece, "<차트번호>")
                STxt = i + 6
                i = InStr(1, TxtRece, "</차트번호>")
                NumTxt = i - STxt
                PChartNum = Mid(TxtRece, STxt, NumTxt)
            '환자이름
                i = InStr(1, TxtRece, "<수진자명>")
                STxt = i + 6
                i = InStr(1, TxtRece, "</수진자명>")
                NumTxt = i - STxt
                PName = Mid(TxtRece, STxt, NumTxt)
            '주민번호
                i = InStr(1, TxtRece, "<주민등록번호>")
                STxt = i + 8
                i = InStr(1, TxtRece, "</주민등록번호>")
                NumTxt = i - STxt
                PJumin = Mid(TxtRece, STxt, NumTxt)
                PJumin = Left(PJumin, 6) & Right(PJumin, 7)
                CalAgeSex PJumin, Format(Date, "yyyy/mm/dd")
                pSex = gPatGen.Sex
                PAge = gPatGen.Age
                
            '내원번호
                i = InStr(1, TxtRece, "<내원번호>")
                STxt = i + 6
                i = InStr(1, TxtRece, "</내원번호>")
                NumTxt = i - STxt
                PID = Mid(TxtRece, STxt, NumTxt)
            '검사코드
                i = InStr(1, TxtRece, "<검사ID>")
                STxt = i + 6
                i = InStr(1, TxtRece, "</검사ID>")
                NumTxt = i - STxt
                PExamCode = Mid(TxtRece, STxt, NumTxt)
            '접수날짜
                i = InStr(1, TxtRece, "<의뢰일>")
                STxt = i + 5
                i = InStr(1, TxtRece, "</의뢰일>")
                NumTxt = i - STxt
                PReceDate = Mid(TxtRece, STxt, NumTxt)
            '<검사번호><입원외래구분>TestNum, IOGubun
                i = InStr(1, TxtRece, "<검사번호>")
                STxt = i + 6
                i = InStr(1, TxtRece, "</검사번호>")
                NumTxt = i - STxt
                TestNum = Mid(TxtRece, STxt, NumTxt)
                
                i = InStr(1, TxtRece, "<입원외래구분>")
                STxt = i + 8
                i = InStr(1, TxtRece, "</입원외래구분>")
                NumTxt = i - STxt
                IOGubun = Mid(TxtRece, STxt, NumTxt)
                
                    SQL = "select equipno, equipcode, examname from equipexam where examcode = '" & PExamCode & "' "
                    res = db_select_Col(gLocal, SQL)
                    
                    If res > 0 Then
                    
                        PEquipno = gReadBuf(0)
                        PEquipCode = gReadBuf(1)
                        PExamname = gReadBuf(2)

                        SQL = "select barcode from pat_res where barcode = '" & PChartNum & "' and examcode = '" & PExamCode & "' and recedate = '" & PReceDate & "'"
                        res = db_select_Col(gLocal, SQL)
                        If res = 0 Then
                                  SQL = "insert into pat_res(examdate, equipno, barcode, receno, pid,"
                            SQL = SQL & "pname, pjumin, page, psex, recedate, "
                            SQL = SQL & "equipcode, examcode, examname, seqno,result)  "
                            SQL = SQL & " values ("
                            SQL = SQL & "'" & Format(Date, "yyyymmdd") & "',"
                            SQL = SQL & "'" & PEquipno & "',"
                            SQL = SQL & "'" & PChartNum & "',"
                            SQL = SQL & "'" & PChartNum & "',"
                            SQL = SQL & "'" & PID & "',"
                            SQL = SQL & "'" & PName & "',"
                            SQL = SQL & "'" & PJumin & "',"
                            SQL = SQL & "'" & PAge & "',"
                            SQL = SQL & "'" & pSex & "',"
                            SQL = SQL & "'" & PReceDate & "',"
                            SQL = SQL & "'" & PEquipCode & "',"
                            SQL = SQL & "'" & PExamCode & "',"
                            SQL = SQL & "'" & PExamname & "',"
                            SQL = SQL & "'" & TestNum & "',"
                            SQL = SQL & "'')"
                        
                        'SQL = "insert into pat_res(equipno, barcode,equipcode, " & vbCrLf & _
                              "examcode, recedate, pid,pname, psex, page, pjumin, examname,examdate, gubun, subcode, result) " & vbCrLf & _
                              "values('" & PEquipno & "','" & PChartNum & "'," & vbCrLf & _
                              "'" & PEquipCode & "','" & PExamCode & "'," & vbCrLf & _
                              "'" & PReceDate & "','" & PID & "'," & vbCrLf & _
                              "'" & PName & "','" & pSex & "'," & vbCrLf & _
                              "'" & PAge & "','" & PJumin & "'," & vbCrLf & _
                              "'" & PExamname & "','" & Format(Date, "yyyymmdd") & "','" & IOGubun & "','" & TestNum & "','')"
                              res = SendQuery(gLocal, SQL)
                              If res = -1 Then
                                SaveQuery SQL
                              End If

                        Else
                            SQL = " Update pat_res Set " & CR & _
                                  " barcode = '" & PChartNum & "', " & CR & _
                                  " pname = '" & PName & "', " & CR & _
                                  " psex = '" & pSex & "', " & CR & _
                                  " page = '" & PAge & "', " & CR & _
                                  " examdate =  '" & Format(frmInterface.txtToday.Text, "yyyymmdd") & "', " & CR & _
                                  " pjumin =  '" & PJumin & "', " & CR & _
                                  " pid =  '" & PID & "', " & CR & _
                                  " seqno =  '" & TestNum & "' " & CR & _
                                  " Where recedate = '" & PReceDate & "' " & CR & _
                                  " and examcode = '" & PExamCode & "'" & CR & _
                                  " and barcode = '" & PChartNum & "' "
                            
'                                  " gubun =  '" & IOGubun & "', " & CR & _
                                  " subcode =  '" & TestNum & "', " & CR & _

                            res = SendQuery(gLocal, SQL)
                        End If
                        
                        
                        BarFlag = 0
                        For j = 1 To vasList.DataRowCnt
                            If GetText(vasList, j, 11) = PChartNum Then
                                BarFlag = 1
                            End If
                        Next
    '                    If BarFlag = 0 Then
    '                        SQL = " Select pid,  pname,  psex, page,pjumin,recedate, barcode " & CR & _
    '                              " From pat_res " & CR & _
    '                              " Where barcode = '" & PChartNum & "' and pid = '" & PID & "' and recedate = '" & PReceDate & "' " & CR & _
    '                              " Group By barcode, pname, psex, page, pjumin, recedate, pid "
    '                        res = db_select_Vas(gLocal, SQL, vasList, vasList.DataRowCnt + 1, 5)
    '                    End If
                    
                    End If
                
                i = InStr(1, TxtPat, "<검사>")
            Else
            End If
        
        Wend
        XmlTxtTail = TxtPat
        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"
        
        'strSrcfile = "C:\UBCare\SINAI\IF\ExamIF_In.xml"
        'strDestFile = App.Path & "\Log\XML_" & Format(Now, "yyyymmddhhmm") & ".xml"
        
        '원본을 대상에 복사
        'FileCopy strSrcfile, strDestFile
            
        ClearSpread vasList
        SQL = " Select barcode, pid, pname,  psex, page,pjumin,recedate, barcode, '', '' as gubun, count(examcode) as cnt " & CR & _
              " From pat_res " & CR & _
              " Where  recedate >= '" & sSch1 & "' and recedate <= '" & sSch2 & "' and result = '' " & CR & _
              "   And sendflag is null " & _
              " Group By barcode, pname, psex, page, pjumin, recedate, pid " & CR & _
              " Order By pname, barcode "
        res = db_select_Vas_Work(gLocal, SQL, vasList, 1, 4)
        'SaveQuery SQL
        vasList.SetText 12, iRow, "A" & Format(Trim(GetText(vasList, iRow, 9)), "yyyymmdd") & "-" & Trim(GetText(vasList, iRow, 10))
        
    Else
        ClearSpread vasList
        SQL = " Select barcode, pid,pname,  psex, page,pjumin,recedate, barcode, '', '' as gubun, count(examcode) as cnt " & CR & _
              " From pat_res " & CR & _
              " Where  recedate >= '" & sSch1 & "' and recedate <= '" & sSch2 & "' and result = '' " & CR & _
              "   And sendflag is null " & _
              " Group By barcode, pname, psex, page, pjumin, recedate, pid " & CR & _
              " Order By pname, barcode "
        res = db_select_Vas_Work(gLocal, SQL, vasList, 1, 4)
        'SaveQuery SQL
        vasList.SetText 12, iRow, "O" & Format(Trim(GetText(vasList, iRow, 9)), "yyyymmdd") & "-" & Trim(GetText(vasList, iRow, 10))
     End If
     
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 13.3
    
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow - 1
    vasActiveCell vasList, lRow - 1, 2
    vasList_Click 2, lRow - 1
End Sub

Private Sub cmdWorkList_Click()
    Dim lRow As Long
    Dim lCol As Long
    Dim lDestRow As Long
    
'    frmInterface.vasID.MaxRows = vasList.DataRowCnt
    frmInterface.vasID.MaxRows = 0
    lDestRow = frmInterface.vasID.DataRowCnt + 1
    
     frmInterface.vasID.MaxRows = lDestRow
     
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        If vasList.Value = 1 Then
            'SetText frmInterface.vasID, Trim(dtpDate), lDestRow, 2
            For lCol = 2 To 12
                
                If lCol = 2 Then
                    SetText frmInterface.vasID, Trim(txtSNo.Text), lDestRow, 2
                    
                    If IsNumeric(txtSNo) Then
                        txtSNo = CInt(txtSNo) + 1
                        gStartNo = txtSNo
                    End If
                ElseIf lCol = 4 Then        '검체번호
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 5
                'ElseIf lCol = 5 Or lCol = 6 Then    '환자번호,환자이름
                ElseIf lCol = 6 Then     '환자번호,환자이름
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol
                ElseIf lCol = 7 Or lCol = 8 Then   '성별,나이,접수번호
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 1
                ElseIf lCol = 11 Then   '성별,나이,접수번호
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 2
                ElseIf lCol = 9 Then    '주민번호
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 8
                ElseIf lCol = 10 Then    '처방일자
                    'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 15
                ElseIf lCol = 12 Then    '검사순서
                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 4
                End If
            Next lCol
'            ElseIf lCol = 7 Or lCol = 8 Or lCol = 11 Then   '성별,나이,접수번호
'                    SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 2
            lDestRow = lDestRow + 1
            frmInterface.vasID.MaxRows = lDestRow
        End If
    Next lRow
    
'    chkAll.Value = 0
'
    Unload Me
End Sub

Private Sub Form_Activate()
    'dtpSDate.SetFocus
    vasActiveCell vasList, 1, 2
End Sub

Private Sub Form_Load()
    dtpSDate.Text = Format(CDate(GetDateFull) - 1, "yyyy/mm/dd")
    dtpEDate.Text = Format(CDate(GetDateFull), "yyyy/mm/dd")
    
    '검사종류
'    With cboGubun
'        .AddItem "1 항목검사", 0
'        .AddItem "2 종합검진", 1
'        .AddItem "3 피보험자검진", 2
'        .AddItem "4 신체검사", 3
'    End With
    
    With cboGubun
        .AddItem "1 항목검사", 0
        .AddItem "2 검진검사", 1
    End With
    cboGubun.ListIndex = 0
    
    ClearSpread vasList
    'vasList.MaxRows = 1
    
    chkAll.Value = 0
    
    txtSNo.Text = gStartNo
End Sub

Private Sub monvCal_DateClick(ByVal DateClicked As Date)
    If iIndex = 0 Then
        dtpSDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
    Else
        dtpEDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
    End If
    monvCal.Visible = False
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 0 Or Row > vasList.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasList.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sCnt As String
    Dim sExamCode As String
    Dim sEquipCode As String
    
    Dim iRow As Integer
    Dim jRow As Integer

    txtDate = GetText(vasList, Row, 10)
    
    txtNo = Trim(GetText(vasList, Row, 4))
    txtPID = Trim(GetText(vasList, Row, 5))
    txtName = Trim(GetText(vasList, Row, 6))
    
    txtSex = Trim(GetText(vasList, Row, 8))
    txtAge = Trim(GetText(vasList, Row, 7))
    
    ClearSpread vasOrder
    
    '검사코드 가져오기
    SQL = " Select '', ExamCode, '' " & vbCrLf & _
          " From ExamRes" & vbCrLf & _
          " Where ReceNo = '" & Trim(GetText(vasList, Row, 11)) & "' " & vbCrLf & _
          " And PID = '" & txtPID & "' " & vbCrLf & _
          " And ExamCode in (" & gAllExam & ") "
    res = db_select_Vas(gServer, SQL, vasOrder)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasOrder.MaxRows = vasOrder.DataRowCnt
    
    For jRow = 1 To vasOrder.DataRowCnt
        SQL = " select ExamName from EquipExam " & vbCrLf & _
              " where Equipno = '" & gEquip & "' and ExamCode = '" & Trim(GetText(vasOrder, jRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        
        If res = 1 Then
            SetText vasOrder, Trim(gReadBuf(0)), jRow, 3
        End If
    Next jRow
    
    sspOrder.Visible = True
    
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Integer
    Dim iCol As Integer
        
    Dim jRow As Integer
    Dim iCnt As Integer
    Dim sRack As String

    iRow = vasList.ActiveRow
    iCol = vasList.ActiveCol
    
    If KeyCode = vbKeyReturn Then
        If iCol = 2 Then    'Rack
            If Trim(GetText(vasList, iRow, iCol)) <> "" Then
                sRack = Trim(GetText(vasList, iRow, 2))
                
                For jRow = 1 To vasList.DataRowCnt
                    vasList.Row = jRow
                    vasList.Col = 1
                    
                    If vasList.Value = 1 Then

                    End If
                Next jRow
            End If
        End If
    End If
End Sub

