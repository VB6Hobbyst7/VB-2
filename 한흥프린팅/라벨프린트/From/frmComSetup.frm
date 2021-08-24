VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmComSetup 
   Caption         =   "바코드 설정"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   18735
   WindowState     =   2  '최대화
   Begin VB.Frame Frame3 
      Caption         =   " 바코드 통신설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5115
      Left            =   12210
      TabIndex        =   12
      Top             =   90
      Width           =   4575
      Begin VB.ComboBox cboBarType 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":0000
         Left            =   1740
         List            =   "frmComSetup.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   40
         Top             =   4140
         Width           =   1740
      End
      Begin VB.ComboBox cboPort 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1740
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Top             =   330
         Width           =   1740
      End
      Begin VB.ComboBox cboSpeed 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":0004
         Left            =   1740
         List            =   "frmComSetup.frx":0006
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   795
         Width           =   1740
      End
      Begin VB.ComboBox cboParity 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":0008
         Left            =   4980
         List            =   "frmComSetup.frx":000A
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   1710
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.ComboBox cboDataBits 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":000C
         Left            =   4980
         List            =   "frmComSetup.frx":000E
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   1260
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.ComboBox cboStopBits 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":0010
         Left            =   4980
         List            =   "frmComSetup.frx":0012
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   2175
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.ComboBox cboDPI 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":0014
         Left            =   1740
         List            =   "frmComSetup.frx":0016
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   1530
         Width           =   1260
      End
      Begin VB.ComboBox cboPrtSpeed 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":0018
         Left            =   1740
         List            =   "frmComSetup.frx":001A
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   1980
         Width           =   1740
      End
      Begin VB.ComboBox cboThermo 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmComSetup.frx":001C
         Left            =   1740
         List            =   "frmComSetup.frx":001E
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   2430
         Width           =   1740
      End
      Begin VB.TextBox txtXPos 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2880
         Width           =   945
      End
      Begin VB.TextBox txtYPos 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   13
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "바코드타입 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   41
         Top             =   4230
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "통신포트 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   35
         Top             =   420
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "전송속도 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   34
         Top             =   885
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "중단 비트 :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   33
         Top             =   2265
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "패리티 :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4140
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "데이터 비트 :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3630
         TabIndex        =   31
         Top             =   1350
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "해상도 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   900
         TabIndex        =   30
         Top             =   1590
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "출력속도 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   690
         TabIndex        =   29
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "헤드 온도 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   28
         Top             =   2490
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "원점 X :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   870
         TabIndex        =   27
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "원점 Y :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   870
         TabIndex        =   26
         Top             =   3480
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DPI"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3060
         TabIndex        =   25
         Top             =   1590
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2910
         TabIndex        =   24
         Top             =   3030
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2910
         TabIndex        =   23
         Top             =   3480
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 바코드 용지설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7815
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   12105
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6270
         TabIndex        =   39
         Top             =   7080
         Width           =   1245
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9090
         TabIndex        =   38
         Top             =   7080
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7680
         TabIndex        =   37
         Top             =   7080
         Width           =   1245
      End
      Begin VB.TextBox txtComName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1830
         TabIndex        =   9
         Top             =   450
         Width           =   1815
      End
      Begin VB.ListBox lstComName 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5910
         Left            =   180
         TabIndex        =   8
         Top             =   960
         Width           =   2685
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "엑셀만들기"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10410
         TabIndex        =   7
         Top             =   420
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "항목추가"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7320
         TabIndex        =   6
         Top             =   420
         Width           =   1455
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "항목제거"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   5
         Top             =   420
         Width           =   1455
      End
      Begin VB.CommandButton cmdComAdd 
         Caption         =   "상품추가"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4050
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtComCode 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         TabIndex        =   2
         Top             =   450
         Width           =   795
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   7080
         Visible         =   0   'False
         Width           =   1245
      End
      Begin FPSpread.vaSpread tblExcel 
         Height          =   525
         Left            =   180
         TabIndex        =   4
         Top             =   7200
         Visible         =   0   'False
         Width           =   2625
         _Version        =   393216
         _ExtentX        =   4630
         _ExtentY        =   926
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
         SpreadDesigner  =   "frmComSetup.frx":0020
      End
      Begin FPSpread.vaSpread spdComBar 
         Height          =   5985
         Left            =   3000
         TabIndex        =   10
         Top             =   960
         Width           =   8955
         _Version        =   393216
         _ExtentX        =   15796
         _ExtentY        =   10557
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         MaxRows         =   13
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   3
         ShadowColor     =   16761024
         SpreadDesigner  =   "frmComSetup.frx":5DDC
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   6720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "고객사"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   11
         Top             =   540
         Width           =   630
      End
   End
   Begin VB.Label lblBar 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   12210
      TabIndex        =   36
      Top             =   5280
      Width           =   4605
   End
End
Attribute VB_Name = "frmComSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim strMaxNum As String
    
    With spdComBar
        .Row = .MaxRows
        .Col = 1
        strMaxNum = .Text
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .Text = strMaxNum + 1
    End With
    
End Sub

Private Sub cmdClear_Click()
    Call SetPort
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDel_Click()
    Dim i As Integer

    With spdComBar
        .Row = .ActiveRow
        .Action = ActionDeleteRow
        .MaxRows = .MaxRows - 1
        
        If .MaxRows = 0 Then
            cmdAdd.Enabled = False
            cmdDel.Enabled = False
            cmdExcel.Enabled = False
        Else
            cmdAdd.Enabled = True
            cmdDel.Enabled = True
            cmdExcel.Enabled = True
            
            For i = 1 To .MaxRows
                .SetText 1, i, i
            Next
        End If
    
    End With


End Sub

Private Sub cmdDelete_Click()
    Dim sqlDoc  As String
    Dim sqlRet  As Long
    
    If MsgBox(txtComName & "를 삭제하시겠습니까?", vbYesNo + vbDefaultButton2 + vbCritical, "기초코드 삭제") = vbYes Then
             
                 sqlDoc = " Delete From INTERFACE001 "
        sqlDoc = sqlDoc & "  Where COMCODE = '" & Trim(txtComCode.Text) & "'"
    
        AdoCn_Jet.Execute sqlDoc, sqlRet
                 
                 sqlDoc = " Delete From INTERFACE002 "
        sqlDoc = sqlDoc & "  Where COMCODE = '" & Trim(txtComCode.Text) & "'"
    
        AdoCn_Jet.Execute sqlDoc, sqlRet
     
        txtComCode.Text = ""
        txtComName.Text = ""
        
        Call SetPort
        
        Call LoadBarList
    
    End If
    
    
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    Dim i As Integer
        
    If spdComBar.DataRowCnt = 0 And spdComBar.DataRowCnt = 0 Then Exit Sub
    
    With spdComBar
        .Row = 1: .Row2 = .MaxRows
        .Col = 4: .Col2 = 4 '.MaxCols
        .BlockMode = True
        strTmp = .Clip
        strTmp = Replace(strTmp, vbNewLine, Chr(9))
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblExcel
        .MaxRows = 1 'spdComBar.MaxRows + 1
        .MaxCols = spdComBar.MaxRows
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = spdComBar.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    With tblExcel
        .MaxRows = 1 'spdComBar.MaxRows + 1
        .MaxCols = spdComBar.MaxRows
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = spdComBar.MaxCols
        .BlockMode = True
        '.CellType = CellTypeEdit
        .Clip = strTmp
        .BlockMode = False
    End With
    
    CommonDialog1.InitDir = App.Path & "\excel"
    CommonDialog1.Filter = "ExCelFile(*.XLS)|*.XLS"
    CommonDialog1.FileName = txtComName.Text & "_" & Format(Now, "yyyymmdd")
    CommonDialog1.ShowSave

'    Call tblExcel.SaveTabFile(CommonDialog1.FileName)
    Call tblExcel.SaveTabFileU(CommonDialog1.FileName)
End Sub

Private Sub cmdSave_Click()
    Dim sqlDoc  As String
    Dim sqlRet  As Long
    Dim intCnt As Integer
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim strValue(8) As String
    
    intCnt = 0
    Erase strValue
    
    intCnt = chkInsUpData(Trim(txtComCode.Text))
        
    '-- insert
    If intCnt = 0 Then
        '-- 장비통신
                 sqlDoc = " Insert into INTERFACE001 "
        sqlDoc = sqlDoc & "        (COMCODE,COMNAME,COM_PORT,COM_SPEED,COM_DATABIT,COM_PARITYBIT,"
        sqlDoc = sqlDoc & "         COM_STOPBIT,COM_HANDSHAK,COM_INPUTMOD,COM_DTR,COM_EOF,COM_NULDIS,COM_RTS)"
        sqlDoc = sqlDoc & "  Values ("
        sqlDoc = sqlDoc & "        '" & Trim(txtComCode.Text) & "',"
        sqlDoc = sqlDoc & "        '" & Trim(txtComName.Text) & "',"
        sqlDoc = sqlDoc & "        '" & cboPort.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboSpeed.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboDataBits.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboParity.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboStopBits.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboDPI.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboPrtSpeed.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboThermo.Text & "',"
        sqlDoc = sqlDoc & "        '" & txtXPos.Text & "',"
        sqlDoc = sqlDoc & "        '" & txtYPos.Text & "',"
        sqlDoc = sqlDoc & "        '" & cboBarType.Text & "')"
        
        AdoCn_Jet.Execute sqlDoc, sqlRet
        
        With spdComBar
            For intRow = 1 To .MaxRows
                .GetText 1, intRow, varTmp: strValue(0) = varTmp
                .GetText 2, intRow, varTmp: strValue(1) = varTmp
                .GetText 3, intRow, varTmp: strValue(2) = varTmp
                .GetText 4, intRow, varTmp: strValue(3) = varTmp
                .GetText 5, intRow, varTmp: strValue(4) = varTmp
                .GetText 6, intRow, varTmp: strValue(5) = varTmp
                .GetText 7, intRow, varTmp: strValue(6) = varTmp
                .GetText 8, intRow, varTmp: strValue(7) = varTmp
                .GetText 9, intRow, varTmp: strValue(8) = varTmp
                
                         sqlDoc = " Insert into INTERFACE002 "
                sqlDoc = sqlDoc & "        (COMCODE,COMNAME,SEQ,TITLEPRT,CONTENTPRT,NAME,POS1,POS2,POS3,POS4,REMARK)"
                sqlDoc = sqlDoc & "  Values ("
                sqlDoc = sqlDoc & "        '" & Trim(txtComCode.Text) & "',"
                sqlDoc = sqlDoc & "        '" & Trim(txtComName.Text) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(0) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(1) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(2) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(3) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(4) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(5) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(6) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(7) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(8) & "')"
                
                AdoCn_Jet.Execute sqlDoc, sqlRet
                Erase strValue
            Next
        End With
    
    Else
        '-- update
        '-- 장비통신
                 sqlDoc = " Update INTERFACE001 Set " & vbNewLine
        sqlDoc = sqlDoc & "        COM_PORT      = '" & cboPort.Text & "'," & vbNewLine
        sqlDoc = sqlDoc & "        COM_SPEED     = '" & cboSpeed.Text & "'," & vbNewLine
        sqlDoc = sqlDoc & "        COM_DATABIT   = '" & cboDataBits.Text & "'," & vbNewLine
        sqlDoc = sqlDoc & "        COM_PARITYBIT = '" & cboParity.Text & "'," & vbNewLine
        sqlDoc = sqlDoc & "        COM_STOPBIT   = '" & cboStopBits.Text & "'," & vbNewLine
        sqlDoc = sqlDoc & "        COM_HANDSHAK  = '" & cboDPI.Text & "'," & vbNewLine  'dpi
        sqlDoc = sqlDoc & "        COM_INPUTMOD  = '" & cboPrtSpeed.Text & "'," & vbNewLine  'speed
        sqlDoc = sqlDoc & "        COM_DTR       = '" & cboThermo.Text & "'," & vbNewLine  'thermo
        sqlDoc = sqlDoc & "        COM_EOF       = '" & txtXPos.Text & "'," & vbNewLine  'x pos
        sqlDoc = sqlDoc & "        COM_NULDIS    = '" & txtYPos.Text & "'," & vbNewLine  'y pos
        sqlDoc = sqlDoc & "        COM_RTS       = '" & cboBarType.Text & "'" & vbNewLine  'y pos
        sqlDoc = sqlDoc & "  Where COMCODE = '" & Trim(txtComCode.Text) & "'"
        
        AdoCn_Jet.Execute sqlDoc, sqlRet
        
        
                 sqlDoc = " Delete From INTERFACE002 "
        sqlDoc = sqlDoc & "  Where COMCODE = '" & Trim(txtComCode.Text) & "'"
        
        AdoCn_Jet.Execute sqlDoc, sqlRet
                 
        With spdComBar
            For intRow = 1 To .MaxRows
                .GetText 1, intRow, varTmp: strValue(0) = varTmp
                .GetText 2, intRow, varTmp: strValue(1) = varTmp
                .GetText 3, intRow, varTmp: strValue(2) = varTmp
                .GetText 4, intRow, varTmp: strValue(3) = varTmp
                .GetText 5, intRow, varTmp: strValue(4) = varTmp
                .GetText 6, intRow, varTmp: strValue(5) = varTmp
                .GetText 7, intRow, varTmp: strValue(6) = varTmp
                .GetText 8, intRow, varTmp: strValue(7) = varTmp
                .GetText 9, intRow, varTmp: strValue(8) = varTmp
                
                         sqlDoc = " Insert into INTERFACE002 "
                sqlDoc = sqlDoc & "        (COMCODE,COMNAME,SEQ,TITLEPRT,CONTENTPRT,NAME,POS1,POS2,POS3,POS4,REMARK)"
                sqlDoc = sqlDoc & "  Values ("
                sqlDoc = sqlDoc & "        '" & Trim(txtComCode.Text) & "',"
                sqlDoc = sqlDoc & "        '" & Trim(txtComName.Text) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(0) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(1) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(2) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(3) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(4) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(5) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(6) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(7) & "',"
                sqlDoc = sqlDoc & "        '" & strValue(8) & "')"
                
                AdoCn_Jet.Execute sqlDoc, sqlRet
                Erase strValue
            Next
        End With
    
    End If

    Call SetPort
    
    Call LoadBarList

End Sub

Private Sub Form_Load()
'    Dim objLogo     As New clsLogo
    
'    With objLogo
'        .DrawingObject = picLogo
'        .Caption = "검사장비 통신설정"
'        .Draw
'    End With

    Call DbConnect_Jet

    Call SetPort
    
    Call LoadBarList
    
End Sub

Private Sub SetPort()
    Dim i As Integer
    
    cboPort.Clear
    cboSpeed.Clear
    cboDataBits.Clear
    cboParity.Clear
    cboStopBits.Clear
    
    For i = 1 To 20
        cboPort.AddItem i
    Next
    cboPort.ListIndex = 0
        
    cboSpeed.AddItem "128000"
    cboSpeed.AddItem "115200"
    cboSpeed.AddItem "57600"
    cboSpeed.AddItem "38400"
    cboSpeed.AddItem "19200"
    cboSpeed.AddItem "14400"
    cboSpeed.AddItem "9600"
    cboSpeed.AddItem "4800"
    cboSpeed.AddItem "2400"
    cboSpeed.AddItem "1200"
    cboSpeed.ListIndex = 0
    
    cboDataBits.AddItem "7"
    cboDataBits.AddItem "8"
    cboDataBits.ListIndex = 1
    
    cboParity.AddItem "Even"
    cboParity.AddItem "Odd"
    cboParity.AddItem "None"
    cboParity.ListIndex = 2
    
    cboStopBits.AddItem "1"
    cboStopBits.AddItem "2"
    cboStopBits.ListIndex = 0

    cboDPI.AddItem "200"
    cboDPI.AddItem "300"
    cboDPI.ListIndex = 0

    For i = 2 To 8
        cboPrtSpeed.AddItem i
    Next
    cboPrtSpeed.ListIndex = 0
    
    For i = 1 To 30
        cboThermo.AddItem i
    Next
    cboThermo.ListIndex = 0
    
    txtXPos.Text = "0"
    txtYPos.Text = "0"
    
    
    cboBarType.AddItem "CODE128"
    cboBarType.AddItem "CODE11"
    cboBarType.AddItem "CODE39"
    cboBarType.AddItem "CODE93"
    cboBarType.AddItem "CODEBAR"
    cboBarType.AddItem "UPC-A"
    cboBarType.AddItem "UPC-E"
    cboBarType.AddItem "EAN-8"
    cboBarType.AddItem "EAN-13"
    
    cboBarType.ListIndex = 0

    

End Sub

Private Sub LoadComInfo(ByVal strComCode As String)
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
             sqlDoc = " Select * From INTERFACE001 "
    sqlDoc = sqlDoc & "  Where COMCODE = '" & strComCode & "' "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    Do While Not adoRS.EOF
        cboPort.Text = Trim$(adoRS("COM_PORT") & "")
        cboSpeed.Text = Trim$(adoRS("COM_SPEED") & "")
        cboDataBits.Text = Trim$(adoRS("COM_DATABIT") & "")
        cboParity.Text = Trim$(adoRS("COM_PARITYBIT") & "")
        cboStopBits.Text = Trim$(adoRS("COM_STOPBIT") & "")
        cboDPI.Text = Trim$(adoRS("COM_HANDSHAK") & "")
        cboPrtSpeed.Text = Trim$(adoRS("COM_INPUTMOD") & "")
        cboThermo.Text = Trim$(adoRS("COM_DTR") & "")
        txtXPos.Text = Trim$(adoRS("COM_EOF") & "")
        txtYPos.Text = Trim$(adoRS("COM_NULDIS") & "")
        cboBarType.Text = Trim$(adoRS("COM_RTS") & "")
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
        
End Sub

Private Sub LoadBarList()
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    lstComName.Clear
    spdComBar.MaxRows = 0
    spdComBar.RowHeight(-1) = 15
    
             sqlDoc = " Select distinct COMCODE, COMNAME From INTERFACE002 "
    sqlDoc = sqlDoc & "  Order By COMCODE,COMNAME "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    
    Do While Not adoRS.EOF
        lstComName.AddItem Trim$(adoRS("COMCODE") & "") & "|" & Trim$(adoRS("COMNAME") & "")
        adoRS.MoveNext
    
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub LoadComBarSet(ByVal strComCode As String)
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim i, j As Integer
    Dim intPosVal(3) As Integer
    Dim intPos(3) As Variant
    Dim ctrl
    
    With spdComBar
        .MaxRows = 0
        i = 1
        
                 sqlDoc = " Select * From INTERFACE002 "
        sqlDoc = sqlDoc & "  Where COMCODE = '" & strComCode & "' "
        sqlDoc = sqlDoc & "  Order By SEQ * 10 "
        
        adoRS.CursorLocation = adUseClient
        adoRS.Open sqlDoc, AdoCn_Jet
        
        If adoRS.RecordCount > 0 Then adoRS.MoveFirst
        .MaxRows = adoRS.RecordCount
        .RowHeight(-1) = 15
        
        '**동적으로 생성된 라벨 삭제
'        For j = 0 To 30
'            For Each ctrl In Me.Controls
'                If ctrl.Name = "txtPrt" & i Then
'                   Controls.Remove ctrl
'                   Set ctrl = Nothing
'                End If
'            Next
'        Next
        
        Do While Not adoRS.EOF
            .Row = i
            .SetText 1, i, Trim$(adoRS("SEQ") & "")
            .SetText 2, i, Trim$(adoRS("TITLEPRT") & "")
            .SetText 3, i, Trim$(adoRS("CONTENTPRT") & "")
            .SetText 4, i, Trim$(adoRS("NAME") & "")
            .SetText 5, i, Trim$(adoRS("POS1") & "")
            .SetText 6, i, Trim$(adoRS("POS2") & "")
            .SetText 7, i, Trim$(adoRS("POS3") & "")
            .SetText 8, i, Trim$(adoRS("POS4") & "")
            .SetText 9, i, Trim$(adoRS("REMARK") & "")
            
            '**동적으로 라벨 생성
            intPosVal(0) = IIf(Trim$(adoRS("POS1") & "") = "", 0, Trim$(adoRS("POS1") & ""))
            intPosVal(1) = IIf(Trim$(adoRS("POS2") & "") = "", 0, Trim$(adoRS("POS2") & ""))
            intPosVal(2) = IIf(Trim$(adoRS("POS3") & "") = "", 0, Trim$(adoRS("POS3") & ""))
            intPosVal(3) = IIf(Trim$(adoRS("POS4") & "") = "", 0, Trim$(adoRS("POS4") & ""))
        
            intPos(0) = (intPosVal(0) * 12) + lblBar.Left
            intPos(1) = (intPosVal(1) * 10) + lblBar.Top
            
            If InStr(Trim$(adoRS("NAME") & ""), "바코드") > 0 Then
                intPos(2) = (intPosVal(2) * 200) * 6
                intPos(3) = intPosVal(3) * 10
            Else
                intPos(2) = intPosVal(2) * 60
                intPos(3) = intPosVal(3) * 5
            End If
        
            '**동적으로 생성된 라벨 삭제
            For Each ctrl In Me.Controls
                If ctrl.Name = "txtPrt" & i Then
                   Controls.Remove ctrl
                   Set ctrl = Nothing
                End If
            Next

    
            If Trim$(adoRS("CONTENTPRT") & "") = "1" Then
                Dim MyTxt As Object
                Set MyTxt = Controls.Add("VB.TextBox", "txtPrt" & i)
                MyTxt.Text = Trim$(adoRS("NAME") & "")
                
                MyTxt.BorderStyle = 1
                MyTxt.Appearance = 0
                MyTxt.BackColor = &HC0FFFF
                
                MyTxt.Move intPos(0), intPos(1), intPos(2), intPos(3)
'                MyTxt.Move 10960, 6730, 3000, 300
                            
'                Debug.Print "Left=" & intPos(0) & "  Top=" & intPos(1) & "  Width=" & intPos(2) & "  Height=" & intPos(3)
                MyTxt.Visible = True
                
                Erase intPos
            End If
            
            i = i + 1
            adoRS.MoveNext
        
        Loop
        adoRS.Close:    Set adoRS = Nothing
    
    End With
    
End Sub


'Private Function PutLYYNConfig(ByVal strConfigNm As String, ByVal strIpValue As String) As String
'
'Dim strFileName As String
'Dim strReturnedString As String
'
'    strFileName = App.Path & "\LYYN.ini"
'
'    strReturnedString = String(1024, " ")
''    WritePrivateProfileString "LYYN", strConfigNm, strIpValue, strFileName
'    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
'    PutLYYNConfig = strReturnedString
'
'End Function
'
'
'
'Private Function GetLYYNConfig(ByVal strConfigNm As String) As String
'
'Dim strFileName As String
'Dim strReturnedString As String
'
'    strFileName = App.Path & "\LYYN.ini"
'
'    strReturnedString = String(1024, " ")
'    GetPrivateProfileString "LYYN", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
'    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
'    GetLYYNConfig = strReturnedString
'
'End Function


Private Sub lstComName_Click()
    Dim strComInfo As Variant
    Dim strComCode As String
    Dim strComName As String
    
    strComInfo = Split(lstComName.Text, "|")
    strComCode = strComInfo(0)
    strComName = strComInfo(1)
    
    Call LoadComBarSet(strComCode)
    
    Call LoadComInfo(strComCode)

    txtComCode.Text = strComCode
    txtComName.Text = strComName
    
    If Trim(txtComName.Text) <> "" Then
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExcel.Enabled = True
    Else
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        cmdExcel.Enabled = False
    End If
    
End Sub

Private Sub txtComName_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strComCode As String
    
    strComCode = ""
    
    If KeyCode = vbKeyReturn Then
        strComCode = chkSameData(Trim(txtComName.Text))
        If strComCode = "" Then
            cmdComAdd.Enabled = True
            cmdAdd.Enabled = True
            cmdDel.Enabled = True
            cmdExcel.Enabled = True
            
            spdComBar.MaxRows = 0
            spdComBar.MaxRows = 1
                        
            txtComCode.Text = getMaxComcode
            
            spdComBar.SetText 1, 1, 1

            
        Else
            cmdComAdd.Enabled = False
            cmdAdd.Enabled = False
            cmdDel.Enabled = False
            cmdExcel.Enabled = False
            
            txtComCode.Text = strComCode
            lstComName.Text = Trim(txtComCode.Text) & "|" & Trim(txtComName.Text)
            Call lstComName_Click
        
        End If
    End If

End Sub

Private Function getMaxComcode() As String
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    getMaxComcode = ""
    
    sqlDoc = " Select max(COMCODE) as COMCODE From INTERFACE001 "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        getMaxComcode = Trim$(adoRS("COMCODE") & "") + 1
        getMaxComcode = Format(getMaxComcode, "00000")
    End If
'    Do While Not adoRS.EOF
        'lstComName.AddItem Trim$(adoRS("COMNAME") & "")
'        chkSameData = Trim$(adoRS("COMCODE") & "")
'        adoRS.MoveNext
        
'    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Function


Private Function chkInsUpData(ByVal strComCode As String) As Integer
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    chkInsUpData = 0
    
             sqlDoc = "Select count(*) as CNT From INTERFACE001 "
    sqlDoc = sqlDoc & " Where COMCODE = '" & strComCode & "'"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        chkInsUpData = Trim$(adoRS("CNT") & "")
    End If
    
    adoRS.Close:    Set adoRS = Nothing
    
End Function


Private Function chkSameData(ByVal strComName As String) As String
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    chkSameData = ""
    
             sqlDoc = " Select COMCODE From INTERFACE002 "
    sqlDoc = sqlDoc & "  Where COMNAME = '" & strComName & "' "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        chkSameData = Trim$(adoRS("COMCODE") & "")
    End If
'    Do While Not adoRS.EOF
        'lstComName.AddItem Trim$(adoRS("COMNAME") & "")
'        chkSameData = Trim$(adoRS("COMCODE") & "")
'        adoRS.MoveNext
        
'    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Function
