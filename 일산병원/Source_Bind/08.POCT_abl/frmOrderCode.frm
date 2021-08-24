VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOrderCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   7320
   ClientLeft      =   2670
   ClientTop       =   1290
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   8190
   Begin VB.Frame Frame3 
      Caption         =   "Hide"
      Height          =   2715
      Left            =   9225
      TabIndex        =   44
      Top             =   2610
      Width           =   4875
      Begin VB.TextBox txtDec 
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
         Left            =   1095
         TabIndex        =   47
         Top             =   405
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox txtRefLow 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1095
         TabIndex        =   46
         Top             =   1245
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtRefHigh 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         TabIndex        =   45
         Top             =   1245
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "소 수 점"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 고 치"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   255
         TabIndex        =   49
         Top             =   1335
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1845
         TabIndex        =   48
         Top             =   1245
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame fracalculation 
      Height          =   1485
      Left            =   9675
      TabIndex        =   19
      Top             =   765
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtIFCC1 
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
         Left            =   510
         TabIndex        =   26
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox txtIFCC2 
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
         Left            =   2400
         TabIndex        =   25
         Top             =   180
         Width           =   585
      End
      Begin VB.CheckBox chkAdd_IFCC 
         Caption         =   "-"
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
         Left            =   1950
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox txteAg1 
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
         Left            =   510
         TabIndex        =   23
         Top             =   660
         Width           =   585
      End
      Begin VB.TextBox txteAg2 
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
         Left            =   2400
         TabIndex        =   22
         Top             =   660
         Width           =   585
      End
      Begin VB.CheckBox chkAdd_eAg 
         Caption         =   "-"
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
         Left            =   1950
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton cmdAddSave 
         Caption         =   "저 장"
         Height          =   345
         Left            =   1950
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "IFCC"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   60
         TabIndex        =   30
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "* A1c"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   29
         Top             =   210
         Width           =   675
      End
      Begin VB.Label eAg 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "eAg"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   28
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "* A1c"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   27
         Top             =   690
         Width           =   675
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6435
      Left            =   120
      TabIndex        =   18
      Top             =   780
      Width           =   4410
      _Version        =   393216
      _ExtentX        =   7779
      _ExtentY        =   11351
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmOrderCode.frx":0000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      ScaleHeight     =   525
      ScaleWidth      =   7935
      TabIndex        =   16
      Top             =   120
      Width           =   7965
      Begin Threed.SSPanel SSPanel1 
         Height          =   585
         Left            =   -60
         TabIndex        =   17
         Top             =   0
         Width           =   7755
         _Version        =   65536
         _ExtentX        =   13679
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "   장비 코드 설정"
         ForeColor       =   8388608
         BackColor       =   16056319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         FloodColor      =   12582912
         Alignment       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   4575
      TabIndex        =   9
      Top             =   720
      Width           =   3525
      Begin VB.Frame Frame2 
         Caption         =   "Comment"
         Height          =   2325
         Left            =   120
         TabIndex        =   31
         Top             =   2490
         Width           =   3285
         Begin VB.ComboBox cboCmtDest 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmOrderCode.frx":0C7F
            Left            =   1050
            List            =   "frmOrderCode.frx":0C81
            TabIndex        =   42
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox cboCmtFlag 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmOrderCode.frx":0C83
            Left            =   1050
            List            =   "frmOrderCode.frx":0C85
            TabIndex        =   41
            Top             =   630
            Width           =   1575
         End
         Begin VB.TextBox txtFHigh 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1050
            TabIndex        =   39
            Top             =   1890
            Width           =   585
         End
         Begin VB.TextBox txtFLow 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2070
            TabIndex        =   38
            Top             =   1890
            Width           =   555
         End
         Begin VB.TextBox txtCmtCode 
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
            Left            =   1050
            TabIndex        =   35
            Top             =   1050
            Width           =   585
         End
         Begin VB.TextBox txtMHigh 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1050
            TabIndex        =   34
            Top             =   1470
            Width           =   585
         End
         Begin VB.TextBox txtMLow 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2070
            TabIndex        =   33
            Top             =   1470
            Width           =   555
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "적용구분"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   43
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lblF 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "여    자"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   40
            Top             =   1980
            Width           =   720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "코    드"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   37
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lblM 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "남    자"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   36
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "성별구분"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   32
            Top             =   720
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         Height          =   495
         Left            =   2550
         TabIndex        =   8
         Top             =   4950
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   1770
         TabIndex        =   7
         Top             =   4950
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   495
         Left            =   990
         TabIndex        =   6
         Top             =   4950
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         Height          =   495
         Left            =   210
         TabIndex        =   5
         Top             =   4950
         Width           =   795
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         Picture         =   "frmOrderCode.frx":0C87
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   15
         Top             =   1140
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtSeq 
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
         Left            =   1110
         TabIndex        =   4
         Top             =   2025
         Width           =   585
      End
      Begin VB.TextBox txtMuch 
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
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   300
         Width           =   2115
      End
      Begin VB.TextBox txtName 
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
         Left            =   1110
         TabIndex        =   3
         Top             =   1590
         Width           =   2115
      End
      Begin VB.TextBox txtCode 
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
         Left            =   1110
         TabIndex        =   2
         Top             =   1170
         Width           =   2115
      End
      Begin VB.TextBox txtEquipCode 
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
         Left            =   1110
         TabIndex        =   1
         Top             =   735
         Width           =   2115
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    서"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   2115
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비구분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   810
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOrderCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ClearText()
'화면초기화
    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtDec = "1"
    txtSeq = ""
    txtRefLow = ""
    txtRefHigh = ""
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
'검사항목 조회
    ClearSpread vasList

'    SQL = "SELECT equipcode, examcode, examname, resprec, seqno, reflow, refhigh " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " group by examcode, equipcode, examname, resprec, seqno, reflow, refhigh "
          
          
          
    SQL = "SELECT equipcode, examcode, examname, resprec, seqno, reflow, refhigh, " & CR & _
          "       cmtdest,cmtflag , cmtcode, cmhigh, cmlow, cfhigh, cflow " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " group by examcode, equipcode, examname, resprec, seqno, reflow, refhigh, cmtdest, cmtflag , cmtcode, cmhigh, cmlow, cfhigh, cflow "
          
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer
'장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT equipcode, examcode, examname, resprec, seqno, reflow, refhigh " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          "   AND equipcode = '" & asEquipCode & "' "
    If Trim(asSuga) <> "" Then
        SQL = SQL & CR & _
          "   AND examcode = '" & asSuga & "' "
    End If
    res = db_select_Col(gLocal, SQL)
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode = 1
End Function


Private Sub cboCmtFlag_Change()
    Call SetCmtView
End Sub

Private Sub cboCmtFlag_Click()
    Call SetCmtView
End Sub

Private Sub SetCmtView()
    Select Case Mid(Trim(cboCmtFlag.Text), 1, 1)
    Case "C"
        lblM.Caption = "공   통"
        lblM.Visible = True
        lblF.Visible = False
        txtMHigh.Visible = True
        txtMLow.Visible = True
        txtFHigh.Visible = False
        txtFLow.Visible = False
    Case "M"
        lblM.Caption = "남   자"
        lblM.Visible = True
        lblF.Visible = False
        txtMHigh.Visible = True
        txtMLow.Visible = True
        txtFHigh.Visible = False
        txtFLow.Visible = False
    Case "F"
        lblM.Caption = "남   자"
        lblM.Visible = False
        lblF.Visible = True
        txtMHigh.Visible = False
        txtMLow.Visible = False
        txtFHigh.Visible = True
        txtFLow.Visible = True
    End Select

End Sub
Private Sub chkAdd_eAg_Click()
    If chkAdd_eAg.Value = 1 Then
        chkAdd_eAg.Caption = "+"
    Else
        chkAdd_eAg.Caption = "-"
    End If
End Sub

Private Sub chkAdd_IFCC_Click()
    If chkAdd_IFCC.Value = 1 Then
        chkAdd_IFCC.Caption = "+"
    Else
        chkAdd_IFCC.Caption = "-"
    End If
End Sub

Private Sub cmdAddSave_Click()
    SQL = "UPDATE calculation "
    SQL = SQL & " SET IFCC1 = '" & txtIFCC1 & "', "
    SQL = SQL & "     IFCC2 = '" & txtIFCC2 & "', "
    SQL = SQL & "     EAG1 = '" & txteAg1 & "', "
    SQL = SQL & "     EAG2 = '" & txteAg2 & "', "
    SQL = SQL & "     ADDIFCC = '" & chkAdd_IFCC.Caption & "', "
    SQL = SQL & "     ADDEAG = '" & chkAdd_eAg.Caption & "' "
    SendQuery gLocal, SQL
    
    fracalculation.Visible = False
End Sub

Private Sub cmdCancel_Click()
    ClearText
    txtEquipCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtCode) = "" Then
'        txtCode.SetFocus
'        Exit Sub
'    End If
        
'    db_BeginTran gLocal
    
     SQL = "Delete From equipexam " & CR & _
          "Where equipno = '" & gEquip & "' " & CR & _
          "  and equipcode = '" & Trim(txtEquipCode) & "' " & CR & _
          "  and examcode = '" & txtCode & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        Exit Sub
    End If
    
'    db_Commit gLocal

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
    Dim lsFlag As String
    Dim lsResFlag As String
    Dim liSeqNo As Integer

    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        MsgBox "장비코드를 입력하세요", vbInformation
        Exit Sub
    End If
    
    If Trim(txtDec) = "" Then
        txtDec.Text = 1

    End If
    
    If IsNumeric(txtSeq) Then
        liSeqNo = CInt(txtSeq)
    Else
        liSeqNo = 0
    End If
    
'    db_BeginTran gLocal
    'equipno, equipcode, examcode, examname, resprec, seqno, reflow, refhigh
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        SQL = "Update equipexam " & CR & _
              "Set resprec = " & Trim(txtDec) & ", " & vbCrLf & _
              "    examname = '" & Trim(txtName) & "', " & vbCrLf & _
              "    reflow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    refhigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    seqno = " & liSeqNo & ", " & vbCrLf & _
              "    cmtdest = '" & Mid(Trim(cboCmtDest.Text), 1, 1) & "', " & vbCrLf & _
              "    cmtflag = '" & Mid(Trim(cboCmtFlag.Text), 1, 1) & "', " & vbCrLf & _
              "    cmtcode = '" & Trim(txtCmtCode) & "', " & vbCrLf & _
              "    cmhigh = '" & Trim(txtMHigh) & "', " & vbCrLf & _
              "    cmlow = '" & Trim(txtMLow) & "', " & vbCrLf & _
              "    cfhigh = '" & Trim(txtFHigh) & "', " & vbCrLf & _
              "    cflow = '" & Trim(txtFLow) & "' " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into equipexam " & CR & _
              "       (equipno,equipcode, examcode, examname, resprec, " & CR & _
              "        seqno, reflow, refhigh," & CR & _
              "        cmtdest, cmtflag, cmtcode," & CR & _
              "        cmhigh, cmlow, cfhigh, cflow )" & CR & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & CR & _
              "         " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "', " & CR & _
              "        '" & Mid(Trim(cboCmtDest.Text), 1, 1) & "', '" & Mid(Trim(cboCmtFlag.Text), 1, 1) & "', '" & Trim(txtCmtCode) & "', " & CR & _
              "        '" & Trim(txtMHigh) & "', '" & Trim(txtMLow) & "', '" & Trim(txtFHigh) & "', '" & Trim(txtFLow) & "') "
    End If

    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        SaveQuery SQL
        Exit Sub
    End If
    
'    db_Commit gLocal
    
    'gEquip = txtMuch
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And fracalculation.Visible = True Then
        fracalculation.Visible = False
    End If
End Sub

Private Sub Form_Load()

            
    ClearText
    DisplayList

    txtMuch = gEquip
    
    cboCmtDest.AddItem "0 적용 안함"
    cboCmtDest.AddItem "1 검사전체적용"
    cboCmtDest.AddItem "2 해당검사적용"
    
    cboCmtDest.ListIndex = 0
    
    cboCmtFlag.AddItem "C 공통"
    cboCmtFlag.AddItem "M 남자"
    cboCmtFlag.AddItem "F 여자"
    
    cboCmtFlag.ListIndex = 0
    
End Sub

Private Sub txtEquipCode_GotFocus()
    SelectFocus txtEquipCode
End Sub

Private Sub txtEquipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtEquipCode = "" Then
            txtEquipCode.SetFocus
            Exit Sub
        End If
        txtCode.SetFocus
    End If
End Sub

Private Sub txtDec_GotFocus()
    SelectFocus txtDec
End Sub

Private Sub txtDec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDec = "" Then
            txtDec.SetFocus
'            Exit Sub
        End If
        
        txtRefLow.SetFocus
    End If
End Sub

Private Sub txtcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        txtCode = UCase(txtCode)
        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf res = 0 Then
            cmdSave.Caption = "저장"
            
        ElseIf res = 1 Then
            cmdSave.Caption = "수정"
            txtName = Trim(gReadBuf(2))
            txtDec = Trim(gReadBuf(3))
            txtSeq = Trim(gReadBuf(4))
            txtRefLow = Trim(gReadBuf(5))
            txtRefHigh = Trim(gReadBuf(6))
        End If
        
        txtName.SetFocus
    End If
End Sub

'Private Sub txtRefhigh_GotFocus()
'    SelectFocus txtRefHigh
'End Sub
'
'Private Sub txtRefhigh_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        'txtPLow.SetFocus
'        cmdSave.SetFocus
'    End If
'End Sub
'
'Private Sub txtRefLow_GotFocus()
'    SelectFocus txtRefLow
'End Sub
'
'Private Sub txtRefLow_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        txtRefHigh.SetFocus
'    End If
'End Sub

Private Sub txtMuch_GotFocus()
    SelectFocus txtMuch
End Sub

Private Sub txtMuch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtMuch.Text) = "" Then
            txtMuch.SetFocus
            Exit Sub
        End If
        txtEquipCode.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtName.Text) = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        txtDec.SetFocus
        
    End If
End Sub

Private Sub txtSeq_GotFocus()
    SelectFocus txtSeq
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSeq.Text) = "" Then
            txtSeq.SetFocus
            Exit Sub
        End If

        cmdSave.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Select Case Col
        Case 1
            vasSort vasList, 1, 2
        Case 2
            vasSort vasList, 2, 1
        End Select
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtCode = Trim(GetText(vasList, Row, 2))
    txtName = Trim(GetText(vasList, Row, 3))
    txtDec = Trim(GetText(vasList, Row, 4))
    txtSeq = Trim(GetText(vasList, Row, 5))
    txtRefLow = Trim(GetText(vasList, Row, 6))
    txtRefHigh = Trim(GetText(vasList, Row, 7))

    cboCmtDest.Text = Trim(GetText(vasList, Row, 8))
    cboCmtFlag.Text = Trim(GetText(vasList, Row, 9))
    txtCmtCode = Trim(GetText(vasList, Row, 10))
    txtMHigh = Trim(GetText(vasList, Row, 11))
    txtMLow = Trim(GetText(vasList, Row, 12))
    txtFHigh = Trim(GetText(vasList, Row, 13))
    txtFLow = Trim(GetText(vasList, Row, 14))

    
    
    cmdSave.Caption = "수정"
End Sub
