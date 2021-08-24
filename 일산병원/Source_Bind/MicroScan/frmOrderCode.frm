VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOrderCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   10275
   ClientLeft      =   2670
   ClientTop       =   1290
   ClientWidth     =   16980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   16980
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   12180
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   41
      Top             =   7140
      Width           =   4665
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "다시 불러오기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1020
         TabIndex        =   44
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "저 장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1020
         TabIndex        =   43
         Top             =   930
         Width           =   2655
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "닫 기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1020
         TabIndex        =   42
         Top             =   1620
         Width           =   2655
      End
   End
   Begin FPSpread.vaSpread vasKit 
      Height          =   6435
      Left            =   14250
      TabIndex        =   40
      Top             =   690
      Width           =   2565
      _Version        =   393216
      _ExtentX        =   4524
      _ExtentY        =   11351
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":0000
   End
   Begin FPSpread.vaSpread vasSlip 
      Height          =   6435
      Left            =   12180
      TabIndex        =   39
      Top             =   690
      Width           =   2055
      _Version        =   393216
      _ExtentX        =   3625
      _ExtentY        =   11351
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":08AD
   End
   Begin FPSpread.vaSpread vasAnti 
      Height          =   8955
      Left            =   6150
      TabIndex        =   38
      Top             =   690
      Width           =   6015
      _Version        =   393216
      _ExtentX        =   10610
      _ExtentY        =   15796
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":0F2F
   End
   Begin FPSpread.vaSpread vasOrg 
      Height          =   8955
      Left            =   120
      TabIndex        =   37
      Top             =   690
      Width           =   6015
      _Version        =   393216
      _ExtentX        =   10610
      _ExtentY        =   15796
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":17C8
   End
   Begin VB.Frame Frame1 
      Height          =   4605
      Left            =   6090
      TabIndex        =   10
      Top             =   7770
      Visible         =   0   'False
      Width           =   3525
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
         Left            =   2130
         TabIndex        =   34
         Top             =   2850
         Width           =   555
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
         Left            =   1110
         TabIndex        =   33
         Top             =   2850
         Width           =   585
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         Height          =   495
         Left            =   2550
         TabIndex        =   9
         Top             =   3420
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   1770
         TabIndex        =   8
         Top             =   3420
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   495
         Left            =   990
         TabIndex        =   7
         Top             =   3420
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         Height          =   495
         Left            =   210
         TabIndex        =   6
         Top             =   3420
         Width           =   795
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         Picture         =   "frmOrderCode.frx":205F
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   17
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
         TabIndex        =   5
         Top             =   2430
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
         Left            =   1110
         TabIndex        =   4
         Top             =   2010
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
         Left            =   1860
         TabIndex        =   36
         Top             =   2850
         Width           =   135
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
         Left            =   270
         TabIndex        =   35
         Top             =   2940
         Width           =   720
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
         TabIndex        =   16
         Top             =   2520
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   1665
         Width           =   720
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
         Left            =   240
         TabIndex        =   13
         Top             =   2085
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   810
         Width           =   720
      End
   End
   Begin VB.Frame fracalculation 
      Height          =   1485
      Left            =   9750
      TabIndex        =   21
      Top             =   7740
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton cmdAddSave 
         Caption         =   "저 장"
         Height          =   345
         Left            =   1950
         TabIndex        =   22
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   690
         Width           =   675
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6435
      Left            =   120
      TabIndex        =   20
      Top             =   7740
      Visible         =   0   'False
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
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
      MaxCols         =   7
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmOrderCode.frx":468D
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   16665
      TabIndex        =   18
      Top             =   60
      Width           =   16695
      Begin Threed.SSPanel SSPanel1 
         Height          =   585
         Left            =   -120
         TabIndex        =   19
         Top             =   0
         Width           =   9825
         _Version        =   65536
         _ExtentX        =   17330
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "   MicroScan 코드설정"
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
End
Attribute VB_Name = "frmOrderCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OrgSort_Flag    As Integer
Dim AntiSort_Flag   As Integer
Dim SlipSort_Flag   As Integer
Dim KitSort_Flag    As Integer
'Dim intSaveIdx      As Integer


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

    SQL = "SELECT equipcode, examcode, examname, resprec, seqno, reflow, refhigh " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " group by examcode, equipcode, examname, resprec, seqno, reflow, refhigh "
          
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt


'균코드 조회
    ClearSpread vasOrg

    SQL = "SELECT seq, orgnm, morgcd, horgcd " & CR & _
          "  From orgtable " & CR & _
          " order by seq "
          
    res = db_select_Vas(gLocal, SQL, vasOrg)
    
    vasOrg.MaxRows = vasOrg.DataRowCnt
    vasOrg.RowHeight(-1) = 12

'항생제코드 조회
    ClearSpread vasAnti

    SQL = "SELECT seq, antinm, manticd, hanticd " & CR & _
          "  From antitable " & CR & _
          " order by seq "
          
    res = db_select_Vas(gLocal, SQL, vasAnti)
    
    vasAnti.MaxRows = vasAnti.DataRowCnt
    vasAnti.RowHeight(-1) = 12

'SLIP코드 조회
    ClearSpread vasSlip

    SQL = "SELECT seq, slipcd " & CR & _
          "  From sliptable " & CR & _
          " order by seq "
          
    res = db_select_Vas(gLocal, SQL, vasSlip)
    
    vasSlip.MaxRows = vasSlip.DataRowCnt
    vasSlip.RowHeight(-1) = 12

'KIT(Other)코드 조회
    ClearSpread vasKit

    SQL = "SELECT seq, kit, other, state, kitseq " & CR & _
          "  From othertable " & CR & _
          " order by seq "
          
    res = db_select_Vas(gLocal, SQL, vasKit)
    
    vasKit.MaxRows = vasKit.DataRowCnt
    vasKit.RowHeight(-1) = 12

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
          "  and examcode = '" & Trim(txtCode) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        Exit Sub
    End If
    
'    db_Commit gLocal

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
Dim i As Integer
Dim varTmp

    With vasOrg
        SQL = "Delete from orgtable "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        For i = 1 To .MaxRows
            SQL = "Insert into orgtable (seq, orgnm, morgcd, horgcd) values ("
            SQL = SQL & "'" & Trim(GetText(vasOrg, i, 1)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasOrg, i, 2)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasOrg, i, 3)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasOrg, i, 4)) & "')"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
            End If
        Next
    End With
    
    With vasAnti
        SQL = "Delete from antitable "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        For i = 1 To .MaxRows
            SQL = "Insert into antitable (seq, antinm, manticd, hanticd) values ("
            SQL = SQL & "'" & Trim(GetText(vasAnti, i, 1)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasAnti, i, 2)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasAnti, i, 3)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasAnti, i, 4)) & "')"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
            End If
        Next
    End With
    
    With vasSlip
        SQL = "Delete from sliptable "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        For i = 1 To .MaxRows
            SQL = "Insert into sliptable (seq, slipcd) values ("
            SQL = SQL & "'" & Trim(GetText(vasSlip, i, 1)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasSlip, i, 2)) & "')"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
            End If
        Next
    End With
    
    With vasKit
        SQL = "Delete from othertable "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        For i = 1 To .MaxRows
            SQL = "Insert into othertable (seq, kit, other, state, kitseq) values ("
            SQL = SQL & "'" & Trim(GetText(vasKit, i, 1)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasKit, i, 2)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasKit, i, 3)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasKit, i, 4)) & "',"
            SQL = SQL & "'" & Trim(GetText(vasKit, i, 5)) & "')"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
            End If
        Next
    End With
    
End Sub

Private Sub cmdRefresh_Click()

    ClearText
    DisplayList

    txtMuch = gEquip

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
              "Set resprec = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    examname = '" & Trim(txtName) & "', " & vbCrLf & _
              "    reflow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    refhigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    seqno = " & liSeqNo & " " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into equipexam (equipno,equipcode, examcode, examname, resprec, seqno , reflow, refhigh) " & CR & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "') "
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
'    Me.Height = 7725
'    Me.Width = 9945
            
    ClearText
    
    DisplayList

    txtMuch = gEquip

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


Private Sub vasAnti_Click(ByVal Col As Long, ByVal Row As Long)
            
    If Row = 0 Then
        If AntiSort_Flag = 1 Then
            Call SpreadSheetSort(vasAnti, Col, 2)
            AntiSort_Flag = 2
        Else
            Call SpreadSheetSort(vasAnti, Col, 1)
            AntiSort_Flag = 1
        End If
    End If


End Sub

Private Sub vasAnti_KeyPress(KeyAscii As Integer)
    Dim varTmp
    Dim strSeq As String
    
    With vasAnti
        If KeyAscii = 13 Then
            If .ActiveRow = .MaxRows Then
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .GetText 2, .MaxRows, varTmp
                If varTmp <> "" Then
                    .MaxRows = .MaxRows + 1
                    .SetText 1, .MaxRows, Format(strSeq + 1, "000")
                End If
            End If
        End If
    End With

End Sub

Private Sub vasAnti_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim varTmp
Dim strSeq  As String

    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 코드 추가", "-", " ▒ 코드 삭제")

    With vasAnti
        Select Case lMenuChosen
            Case 1
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .MaxRows = .MaxRows + 1
                .SetText 1, .MaxRows, Format(strSeq + 1, "000")
            Case 3
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
        End Select
    End With

End Sub

Private Sub vasKit_Click(ByVal Col As Long, ByVal Row As Long)
            
    If Row = 0 Then
        If KitSort_Flag = 1 Then
            Call SpreadSheetSort(vasKit, Col, 2)
            KitSort_Flag = 2
        Else
            Call SpreadSheetSort(vasKit, Col, 1)
            KitSort_Flag = 1
        End If
    End If


End Sub

Private Sub vasKit_KeyPress(KeyAscii As Integer)
    Dim varTmp
    Dim strSeq As String
    
    With vasKit
        If KeyAscii = 13 Then
            If .ActiveRow = .MaxRows Then
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .GetText 2, .MaxRows, varTmp
                If varTmp <> "" Then
                    .MaxRows = .MaxRows + 1
                    .SetText 1, .MaxRows, Format(strSeq + 1, "000")
                End If
            End If
        End If
    End With

End Sub

Private Sub vasKit_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim varTmp
Dim strSeq  As String

    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 코드 추가", "-", " ▒ 코드 삭제")

    With vasKit
        Select Case lMenuChosen
            Case 1
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .MaxRows = .MaxRows + 1
                .SetText 1, .MaxRows, Format(strSeq + 1, "000")
            Case 3
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
        End Select
    End With

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

    
    
    cmdSave.Caption = "수정"
End Sub

Private Sub vasOrg_Click(ByVal Col As Long, ByVal Row As Long)
            
    If Row = 0 Then
        If OrgSort_Flag = 1 Then
            Call SpreadSheetSort(vasOrg, Col, 2)
            OrgSort_Flag = 2
        Else
            Call SpreadSheetSort(vasOrg, Col, 1)
            OrgSort_Flag = 1
        End If
    End If

End Sub

Private Sub vasOrg_KeyPress(KeyAscii As Integer)
    Dim varTmp
    Dim strSeq As String
    
    With vasOrg
        If KeyAscii = 13 Then
            If .ActiveRow = .MaxRows Then
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .GetText 2, .MaxRows, varTmp
                If varTmp <> "" Then
                    .MaxRows = .MaxRows + 1
                    .SetText 1, .MaxRows, Format(strSeq + 1, "000")
                End If
            End If
        End If
    End With
    
End Sub

Private Sub vasOrg_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim varTmp
Dim strSeq  As String

    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 코드 추가", "-", " ▒ 코드 삭제")

    With vasOrg
        Select Case lMenuChosen
            Case 1
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .MaxRows = .MaxRows + 1
                .SetText 1, .MaxRows, Format(strSeq + 1, "000")
            Case 3
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
        End Select
    End With

End Sub

Private Sub vasSlip_Click(ByVal Col As Long, ByVal Row As Long)
            
    If Row = 0 Then
        If SlipSort_Flag = 1 Then
            Call SpreadSheetSort(vasSlip, Col, 2)
            SlipSort_Flag = 2
        Else
            Call SpreadSheetSort(vasSlip, Col, 1)
            SlipSort_Flag = 1
        End If
    End If


End Sub

Private Sub vasSlip_KeyPress(KeyAscii As Integer)
    Dim varTmp
    Dim strSeq As String
    
    With vasSlip
        If KeyAscii = 13 Then
            If .ActiveRow = .MaxRows Then
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .GetText 2, .MaxRows, varTmp
                If varTmp <> "" Then
                    .MaxRows = .MaxRows + 1
                    .SetText 1, .MaxRows, Format(strSeq + 1, "000")
                End If
            End If
        End If
    End With

End Sub

Private Sub vasSlip_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim varTmp
Dim strSeq  As String

    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 코드 추가", "-", " ▒ 코드 삭제")

    With vasSlip
        Select Case lMenuChosen
            Case 1
                .GetText 1, .MaxRows, varTmp: strSeq = varTmp
                .MaxRows = .MaxRows + 1
                .SetText 1, .MaxRows, Format(strSeq + 1, "000")
            Case 3
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
        End Select
    End With

End Sub
