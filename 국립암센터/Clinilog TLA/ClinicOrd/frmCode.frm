VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11850
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5640
      TabIndex        =   40
      Top             =   6720
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '평면
      Caption         =   "순서"
      Height          =   645
      Left            =   2790
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   2940
      TabIndex        =   36
      Top             =   7170
      Visible         =   0   'False
      Width           =   1965
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   5055
      Left            =   960
      TabIndex        =   35
      Top             =   1290
      Visible         =   0   'False
      Width           =   5445
      _Version        =   393216
      _ExtentX        =   9604
      _ExtentY        =   8916
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
      MaxRows         =   1000
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmCode.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   7305
      Left            =   7500
      TabIndex        =   1
      Top             =   750
      Width           =   4275
      Begin VB.CheckBox Check1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "인덱스플래그"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   4560
         Width           =   2145
      End
      Begin VB.TextBox txtOrdCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   38
         Top             =   3990
         Width           =   2655
      End
      Begin VB.TextBox txtName1 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1410
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   3300
         Width           =   2655
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "검사"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2010
         TabIndex        =   32
         Top             =   3330
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optRes 
         Caption         =   "계산"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3090
         TabIndex        =   31
         Top             =   3330
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   29
         Top             =   1515
         Width           =   2655
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   22
         Top             =   210
         Width           =   2655
      End
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   18
         Top             =   2850
         Width           =   2655
      End
      Begin VB.TextBox txtRefLow 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   16
         Top             =   2400
         Width           =   945
      End
      Begin VB.TextBox txtRefHigh 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2820
         TabIndex        =   15
         Top             =   2400
         Width           =   945
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3150
         TabIndex        =   12
         Top             =   6540
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2160
         TabIndex        =   11
         Top             =   6540
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1170
         TabIndex        =   10
         Top             =   6540
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         TabIndex        =   9
         Top             =   6540
         Width           =   975
      End
      Begin VB.TextBox txtRang 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   8
         Top             =   1965
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   5
         Top             =   660
         Width           =   2655
      End
      Begin VB.TextBox txtEquipCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   3
         Top             =   225
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtSeqNo 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   24
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkUse 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사    용"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   30
         Top             =   3660
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오더코드"
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
         Left            =   240
         TabIndex        =   39
         Top             =   4050
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
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
         Left            =   240
         TabIndex        =   34
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
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
         Left            =   240
         TabIndex        =   23
         Top             =   285
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "단    위"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2925
         Width           =   1050
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
         Left            =   2520
         TabIndex        =   17
         Top             =   2475
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 조 치"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2490
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비소수"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과처리"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1590
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사약칭"
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
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
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
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    서"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1140
         Width           =   1050
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7215
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   7425
      _Version        =   393216
      _ExtentX        =   13097
      _ExtentY        =   12726
      _StockProps     =   64
      ColHeaderDisplay=   1
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
      MaxCols         =   22
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmCode.frx":029A
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "  장비 코드 설정"
      ForeColor       =   8388608
      BackColor       =   16774393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.26
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.ComboBox cboOrdGubun 
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
         ItemData        =   "frmCode.frx":2093
         Left            =   9060
         List            =   "frmCode.frx":20B2
         TabIndex        =   28
         Top             =   390
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cboGubun 
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
         ItemData        =   "frmCode.frx":214A
         Left            =   8610
         List            =   "frmCode.frx":2157
         TabIndex        =   26
         Top             =   150
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtEquip 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5310
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오더구분"
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
         Left            =   7530
         TabIndex        =   27
         Top             =   180
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비구분"
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
         Left            =   4140
         TabIndex        =   21
         Top             =   195
         Visible         =   0   'False
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsGubun As String
Dim lsCode As String
Dim lsName As String
Dim lsSeqNo As String
Dim lsRang As String
Dim lsRefLow As String
Dim lsRefHigh As String
Dim lsUnit As String

Sub ClearText()
    lsGubun = ""
    lsName = ""
    lsRang = ""

    
    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtOrdCode = ""
    txtSeqNo = ""
    
    cboGubun.ListIndex = -1
    txtRang = ""
    txtRefLow = ""
    txtRefHigh = ""
    txtUnit = ""
    
    txtName1 = ""
    
    Check1.Value = 0
    
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
    SQL = "SELECT EquipCode,ExamCode, ExamName, ExamName1,  OrdCode, Seqno,OrdGubun, RSGubun,PointSize,RefLow + ' - ' + RefHigh, IndexFlag, UnitCode " & CR & _
          "  From EquipExam " & CR & _
          " WHERE Equip = '" & gEquip & "' " & CR & _
          " Order by seqno "
          
    db_select_Vas gLocal, SQL, vasList
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asExamCode As String) As Integer
'장비코드와 검사명에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EquipCode, ExamCode, ExamName, SeqNo, RSGubun, PointSize, RefLow, RefHigh, UnitCode " & CR & _
          "  From EquipExam " & CR & _
          " WHERE Equip = '" & gEquip & "' " & CR & _
          "   AND EquipCode = '" & asEquipCode & "' "
    If Trim(asExamCode) <> "" Then
        SQL = SQL & "   AND ExamCode = '" & asExamCode & "' "
    End If
    res = db_select_Col(gLocal, SQL)
    
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Then
        Exit Function
    End If
    
    lsCode = Trim(gReadBuf(1))
    lsName = Trim(gReadBuf(2))
    lsSeqNo = Trim(gReadBuf(3))
    lsGubun = Trim(gReadBuf(4))
    lsRang = Trim(gReadBuf(5))
    lsRefLow = Trim(gReadBuf(6))
    lsRefHigh = Trim(gReadBuf(7))
    lsUnit = Trim(gReadBuf(8))
    
    ExistOfEquipCode = 1
End Function

Function Select_Suga_Info(asSuga As String) As Integer
    Select_Suga_Info = -1
    
    If Trim(asSuga) = "" Then
        Exit Function
    End If
    
    SQL = "Select coifcode, coifrclf, coifrfit, coifrfpr, coifleng, coifcona " & CR & _
          "From ABCCOIFM" & CR & _
          "Where coifcode = '" & asSuga & "' "
    res = db_select_Col(gServer, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    ElseIf res = 0 Then
        Select_Suga_Info = 0
        Exit Function
    End If
    If Trim(gReadBuf(0)) <> asSuga Then
        Select_Suga_Info = 0
        Exit Function
    End If
    
    txtRang = ""
    Select Case Left(Trim(gReadBuf(1)), 1)
    Case "N"
        If IsNumeric(Trim(gReadBuf(3))) = True Then
            If CInt(gReadBuf(3)) > 0 Then
                cboGubun.ListIndex = 1
                txtRang = Trim(gReadBuf(2)) & "." & Trim(gReadBuf(3))
            Else
                cboGubun.ListIndex = 0
                txtRang = Trim(gReadBuf(2))
            End If
        Else
            cboGubun.ListIndex = -1
        End If
    Case "T"
        cboGubun.ListIndex = 2
        txtRang = Trim(gReadBuf(4))
    Case Else
        cboGubun.ListIndex = -1
    End Select
    txtName = Trim(gReadBuf(5))
    
    Select_Suga_Info = 1
End Function

Private Sub cboGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboGubun.ListIndex < 0 Then
            cboGubun.SetFocus
            Exit Sub
        End If
        
        txtRang.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearText
    txtCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
         
'    db_BeginTran gLocal
    
    SQL = "Delete from EquipExam " & CR & _
          "Where Equip = '" & gEquip & "' " & CR & _
          "  and EquipCode = '" & Trim(txtEquipCode) & "' " & CR & _
          "  and ExamCode = '" & Trim(txtCode) & "' "
          
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
    Dim lsGubun As String
    Dim lsCutOffFlag As String
    Dim lsNegFlag As String
    Dim lsPosFlag As String
    Dim lsTiterFlag As String
    
    If Trim(txtCode) = "" Then
        txtCode.SetFocus
        Exit Sub
    End If
    txtEquipCode = txtCode
        
    If Trim(txtRang) = "" Then
        txtRang.Text = "0"
    End If
    
    If optOrder.Value = True Then
        lsGubun = "O"
    End If
    If optRes.Value = True Then
        lsGubun = "R"
    End If
        
    If IsNumeric(txtSeqNo) = False Then
        If IsNumeric(txtEquipCode) = True Then
            txtSeqNo = txtEquipCode
        Else
            txtSeqNo = 0
        End If
    End If
    
    If Trim(txtOrdCode) = "" Then
        txtOrdCode = txtCode
    End If
'    db_BeginTran gLocal
    
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        SQL = "Update EquipExam " & CR & _
              "Set ExamName = '" & Trim(txtName) & "', " & CR & _
              "    ExamName1 = '" & Trim(txtName1) & "', " & vbCrLf & _
              "    OrdCode = '" & Trim(txtOrdCode) & "', " & vbCrLf & _
              "    Seqno = " & Trim(txtSeqNo) & ", " & vbCrLf & _
              "    OrdGubun = '" & lsGubun & "', " & vbCrLf & _
              "    RSGubun = '" & Trim(txtGubun.Text) & "', " & CR & _
              "    PointSize = " & Trim(txtRang) & ", " & CR & _
              "    RefLow = '" & Trim(txtRefLow) & "', " & CR & _
              "    RefHigh = '" & Trim(txtRefHigh) & "', " & CR & _
              "    IndexFlag = '" & Check1.Value & "', " & vbCrLf & _
              "    UnitCode = '" & Trim(txtUnit) & "' " & CR & _
              "Where Equip = '" & gEquip & "' " & CR & _
              "  and EquipCode = '" & Trim(txtEquipCode) & "' " & CR & _
              "  and ExamCode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into EquipExam (Equip,EquipCode,ExamCode,OrdCode, ExamName,ExamName1,SeqNo, " & _
              " OrdGubun,RSGubun, PointSize,RefLow,RefHigh, IndexFlag, UnitCode) " & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtOrdCode) & "', '" & Trim(txtName) & "', '" & Trim(txtName1) & "', " & _
              Trim(txtSeqNo) & ", '" & lsGubun & "', '" & Trim(txtGubun.Text) & "', " & Trim(txtRang) & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "', '" & Check1.Value & "', '" & Trim(txtUnit) & "' ) "
    End If
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        Exit Sub
    End If
    
'    db_Commit gLocal
    
    DisplayList
    
    cmdCancel_Click
    
End Sub

Private Sub Command1_Click()
    Dim lRow As Long
    
    SQL = "Delete from EquipExam " '& CR & _
          "Where Equip = '" & gEquip & "' " & CR & _
          "  and EquipCode = '" & Trim(txtEquipCode) & "' " & CR & _
          "  and ExamCode = '" & Trim(txtCode) & "' "
          
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        Exit Sub
    End If
    
    For lRow = 1 To vasTemp.DataRowCnt
        txtEquipCode = Trim(GetText(vasTemp, lRow, 1))
        txtCode = Trim(GetText(vasTemp, lRow, 2))
        txtName = Trim(GetText(vasTemp, lRow, 3))
        txtSeqNo = Trim(GetText(vasTemp, lRow, 1))
        cmdSave_Click
    Next lRow
End Sub

Private Sub Command2_Click()
    SQL = "Update EquipExam " & CR & _
          "Set Seqno = 999 " & vbCrLf & _
          "Where Equip = '" & gEquip & "' "
    res = SendQuery(gLocal, SQL)
End Sub

Private Sub Command3_Click()
    Dim lsCode, lsOrdcode As String
    Dim lRow As Long
    
    For lRow = 1 To vasTemp.DataRowCnt
        lsCode = Trim(GetText(vasTemp, lRow, 1))
        lsOrdcode = Mid(lsCode, 2)
        
        SQL = "Update equipexam set ordcode = '" & lsOrdcode & "' where examcode = '" & lsCode & "' "
        res = SendQuery(gLocal, SQL)
        
    Next lRow
    
End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
    
    SQL = "select ExamName1 from EquipExam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table EquipExam " & vbCrLf & _
              "add column ExamName1 varchar(100) "
        res = SendQuery(gLocal, SQL)
    End If
    
    ClearText
    DisplayList
    
    txtEquip = gEquip
End Sub

Private Sub Label14_Click()

End Sub

Private Sub SSPanel1_DblClick()
    frmData.Show
End Sub

Private Sub txtCode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If res = 1 Then
    
            txtCode = lsCode
            txtName = lsName
            txtSeqNo = lsSeqNo
            Select Case lsGubun
            Case "I"
                cboGubun.ListIndex = 0
            Case "F"
                cboGubun.ListIndex = 1
            Case "T"
                cboGubun.ListIndex = 2
            Case Else
                cboGubun.ListIndex = -1
            End Select
            txtRang = lsRang
            txtRefLow = lsRefLow
            txtRefHigh = lsRefHigh
            txtUnit = lsUnit
                
        End If
                
        txtName.SetFocus
    End If
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
        res = ExistOfEquipCode(Trim(txtEquipCode))
        If res = 1 Then
    
            txtCode = lsCode
            txtName = lsName
            txtSeqNo = lsSeqNo
            Select Case lsGubun
            Case "I"
                cboGubun.ListIndex = 0
            Case "F"
                cboGubun.ListIndex = 1
            Case "T"
                cboGubun.ListIndex = 2
            Case Else
                cboGubun.ListIndex = -1
            End Select
            txtRang = lsRang
            txtRefLow = lsRefLow
            txtRefHigh = lsRefHigh
            txtUnit = lsUnit
                
        End If
        txtCode.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtName = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        'cboGubun.SetFocus
    End If
    
End Sub

Private Sub txtRang_GotFocus()
    SelectFocus txtRang
End Sub

Private Sub txtRang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRang = "" Then
            txtRang.SetFocus
            Exit Sub
        End If
        
        txtRefLow.SetFocus
    End If
End Sub

Private Sub txtRefHigh_GotFocus()
    SelectFocus txtRefHigh
End Sub

Private Sub txtRefHigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRefHigh = "" Then
            txtRefHigh.SetFocus
            Exit Sub
        End If
        
        txtUnit.SetFocus
    End If
End Sub

Private Sub txtRefLow_GotFocus()
    SelectFocus txtRefLow
End Sub

Private Sub txtRefLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRefLow = "" Then
            txtRefLow.SetFocus
            Exit Sub
        End If
        
        txtRefHigh.SetFocus
    End If
End Sub

Private Sub txtUnit_GotFocus()
    SelectFocus txtUnit
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtUnit = "" Then
            txtUnit.SetFocus
            Exit Sub
        End If

        cmdSave.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)

    Dim i As Integer
    
    If Row = 0 Then
        vasSort vasList, Col
    End If
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtCode = Trim(GetText(vasList, Row, 2))
    txtName = Trim(GetText(vasList, Row, 3))
    txtName1 = Trim(GetText(vasList, Row, 4))
    txtOrdCode = Trim(GetText(vasList, Row, 5))
    txtSeqNo = Trim(GetText(vasList, Row, 6))
    For i = 0 To cboOrdGubun.ListCount - 1
        If Left(cboOrdGubun.List(i), 1) = Trim(GetText(vasList, Row, 7)) Then
            cboOrdGubun.ListIndex = i
            Exit For
        End If
    Next i
    
    txtGubun = Trim(GetText(vasList, Row, 8))
'    Select Case Trim(GetText(vasList, Row, 6))
'    Case "I"
'        cboGubun.ListIndex = 0
'    Case "F"
'        cboGubun.ListIndex = 1
'    Case "T"
'        cboGubun.ListIndex = 2
'    Case Else
'        cboGubun.ListIndex = -1
'    End Select
    
    txtRang = Trim(GetText(vasList, Row, 9))
    
    
    i = InStr(1, Trim(GetText(vasList, Row, 10)), "-")
    If i > 0 Then
        txtRefLow = Trim(Mid(GetText(vasList, Row, 10), 1, i - 1))
        txtRefHigh = Trim(Mid(GetText(vasList, Row, 10), i + 1))
    End If
    
    If Trim(GetText(vasList, Row, 11)) = "1" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    
    txtUnit = Trim(GetText(vasList, Row, 12))
    
    cmdSave.Caption = "수정"
    
End Sub
