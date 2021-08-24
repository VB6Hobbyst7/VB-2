VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOrderCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11700
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSFrame ssfrmExam 
      Height          =   5205
      Left            =   2310
      TabIndex        =   29
      Top             =   6030
      Visible         =   0   'False
      Width           =   6465
      _Version        =   65536
      _ExtentX        =   11404
      _ExtentY        =   9181
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
      Begin FPSpread.vaSpread vasExam 
         Height          =   4905
         Left            =   180
         TabIndex        =   30
         Top             =   180
         Width           =   6225
         _Version        =   393216
         _ExtentX        =   10980
         _ExtentY        =   8652
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
         MaxCols         =   4
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmOrderCode.frx":0000
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7305
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   7245
      _Version        =   393216
      _ExtentX        =   12779
      _ExtentY        =   12885
      _StockProps     =   64
      ColHeaderDisplay=   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":190C
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   11565
      _Version        =   65536
      _ExtentX        =   20399
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "       AX4030 장비 코드 설정"
      ForeColor       =   8388608
      BackColor       =   16056319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   5475
      Left            =   7350
      TabIndex        =   1
      Top             =   630
      Width           =   4275
      Begin IF_AX4030.MDButton cmdClose 
         Height          =   585
         Left            =   3150
         TabIndex        =   38
         Top             =   4770
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "종료"
      End
      Begin IF_AX4030.MDButton cmdCancel 
         Height          =   585
         Left            =   2130
         TabIndex        =   37
         Top             =   4770
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Clear"
      End
      Begin IF_AX4030.MDButton cmdDelete 
         Height          =   585
         Left            =   1110
         TabIndex        =   36
         Top             =   4770
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "삭제"
      End
      Begin IF_AX4030.MDButton cmdSave 
         Height          =   585
         Left            =   90
         TabIndex        =   35
         Top             =   4770
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "저장"
      End
      Begin VB.CheckBox Check2 
         Height          =   285
         Left            =   1380
         TabIndex        =   34
         Top             =   5340
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmOrderCode.frx":3499
         Left            =   1380
         List            =   "frmOrderCode.frx":349B
         TabIndex        =   32
         Top             =   3630
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3720
         Picture         =   "frmOrderCode.frx":349D
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   28
         Top             =   1140
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CheckBox Check1 
         Height          =   285
         Left            =   1410
         TabIndex        =   27
         Top             =   4050
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtSeq 
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
         Top             =   1950
         Width           =   945
      End
      Begin VB.TextBox txtDelta 
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
         TabIndex        =   21
         Top             =   4230
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtPHigh 
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
         Left            =   2700
         TabIndex        =   19
         Top             =   4170
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtPLow 
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
         Left            =   1320
         TabIndex        =   17
         Top             =   4170
         Visible         =   0   'False
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
         Left            =   2760
         TabIndex        =   15
         Top             =   4860
         Visible         =   0   'False
         Width           =   945
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
         Left            =   1380
         TabIndex        =   13
         Top             =   4860
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtMuch 
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
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
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
         TabIndex        =   9
         Top             =   1530
         Width           =   2655
      End
      Begin VB.TextBox txtDec 
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
         Left            =   1290
         TabIndex        =   7
         Top             =   4155
         Visible         =   0   'False
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
         TabIndex        =   5
         Top             =   1104
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
         Top             =   672
         Width           =   2655
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Flag"
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
         Left            =   660
         TabIndex        =   33
         Top             =   5340
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체타입"
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
         Left            =   210
         TabIndex        =   31
         Top             =   3675
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사여부"
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
         Left            =   210
         TabIndex        =   26
         Top             =   4080
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label12 
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
         Left            =   270
         TabIndex        =   25
         Top             =   2010
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "%"
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
         Left            =   2460
         TabIndex        =   23
         Top             =   4350
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "델    타"
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
         TabIndex        =   22
         Top             =   4305
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label9 
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
         Left            =   2490
         TabIndex        =   20
         Top             =   3990
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "패    닉"
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
         Left            =   150
         TabIndex        =   18
         Top             =   4035
         Visible         =   0   'False
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
         Left            =   2460
         TabIndex        =   16
         Top             =   4935
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label3 
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
         Left            =   210
         TabIndex        =   14
         Top             =   4935
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   240
         TabIndex        =   10
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label5 
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
         TabIndex        =   8
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "정 확 도"
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
         Left            =   330
         TabIndex        =   6
         Top             =   4080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label2 
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
         TabIndex        =   4
         Top             =   1170
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
         Top             =   750
         Width           =   1020
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

    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtDec = "1"
    txtRefLow = ""
    txtRefHigh = ""
    txtPLow = ""
    txtPHigh = ""
    txtDelta = ""
    txtSeq = ""
    cmbType = ""
    
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
'    SQL = "SELECT equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue " & CR & _
'          "  From equipexam " & CR & _
'          " WHERE equipno = '" & gEquip & "' " & CR & _
'          " Order by seqno, EquipCode "
          
    SQL = "SELECT equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue, examflag, seqno, examtype " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " Order by examcode "
          
    db_select_Vas gLocal, SQL, vasList
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer
'장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue   " & CR & _
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

Function Select_Suga_Info(asSuga As String) As Integer
    Select_Suga_Info = -1
    
    If Trim(asSuga) = "" Then
        Exit Function
    End If
    
'    If Not Connect_Server Then
'        cn_Server_Flag = False
'        Exit Function
'    Else
'        cn_Server_Flag = True
'    End If
    
    'Connect_Server_Neosoft
    
    SQL = " Select LABM_ID, LABM_NAME " & CR & _
          " from CC_LABM " & CR & _
          " where LABM_ID = '" & Trim(asSuga) & "' "

    res = db_select_Col_Neo(gServer, SQL)
    
'    If cn_Server_Flag Then DisConnect_Server
    
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
    
    txtDec = ""
    txtName = Trim(gReadBuf(1))
    txtRefLow = ""
    txtRefHigh = ""
    txtPLow = ""
    txtPHigh = ""
    
    txtDelta = ""
    
    Select_Suga_Info = 1
End Function

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSeq.SetFocus
    End If
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
        
    db_BeginTran gLocal
    
    SQL = "Delete From equipexam " & CR & _
          "Where equipno = '" & gEquip & "' " & CR & _
          "  and equipcode = '" & Trim(txtEquipCode) & "' " & CR & _
          "  and examcode = '" & Trim(txtCode) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        db_RollBack gLocal
        Exit Sub
    End If
    
    db_Commit gLocal

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
    Dim lsFlag As String
    Dim lsResFlag As String
    Dim liSeqNo As Integer
    
    'gEquip = "0LX20"
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        MsgBox "장비코드를 입력하세요", vbInformation
        Exit Sub
    End If
    
    
'    If Trim(txtCode) = "" Then
'        txtCode.SetFocus
'        MsgBox "검사코드를 입력하세요", vbInformation
'        Exit Sub
'    End If
    
    If Trim(txtDec) = "" Then
        txtDec.Text = 1
'        txtDec.SetFocus
'        Exit Sub
    End If
    
    If IsNumeric(txtSeq) Then
        liSeqNo = CInt(txtSeq)
    Else
        liSeqNo = 0
    End If
    
    If Check1.Value = 1 Then
        lsFlag = "1"
    Else
        lsFlag = "0"
    End If
    
    If Check2.Value = 1 Then
        lsResFlag = "1"
    Else
        lsResFlag = "0"
    End If
    
    
    
    db_BeginTran gLocal
    'examcode, examname, resprec, refmlow, refmhigh, refwlow, refwhigh
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        SQL = "Update equipexam " & CR & _
              "Set resprec = '" & Trim(lsResFlag) & "', " & vbCrLf & _
              "    examname = '" & Trim(txtName) & "', " & vbCrLf & _
              "    reflow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    refhigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    paniclow = '" & Trim(txtPLow) & "', " & vbCrLf & _
              "    panichigh = '" & Trim(txtPHigh) & "', " & vbCrLf & _
              "    deltavalue = '" & Trim(txtDelta) & "', " & vbCrLf & _
              "    examflag = " & lsFlag & ", " & vbCrLf & _
              "    examtype = '" & Trim(cmbType) & "', " & vbCrLf & _
              "    seqno = " & liSeqNo & " " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into equipexam (equipno,equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue, examflag, seqno, examtype ) " & CR & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(lsResFlag) & "', '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "', '" & Trim(txtPLow) & "', '" & Trim(txtPHigh) & "', '" & Trim(txtDelta) & "', " & lsFlag & ", " & liSeqNo & ", '" & Trim(cmbType) & "' ) "
    End If


    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        db_RollBack gLocal
        SaveQuery SQL
        Exit Sub
    End If
    
    db_Commit gLocal
    
    'gEquip = txtMuch
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    Me.Height = 8600
    Me.Width = 11970
            
    ClearText
    DisplayList
    
    cmbType.AddItem "SE", 0
    cmbType.AddItem "TU", 1
    cmbType.AddItem "UR", 2
    cmbType.AddItem "SF", 3
    
    
    txtMuch = gEquip
End Sub

Private Sub txtDelta_GotFocus()
    SelectFocus txtDelta
End Sub

Private Sub txtDelta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSave.SetFocus
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
        txtCode = UCase(txtCode)
        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf res = 0 Then
            cmdSave.Caption = "저장"
            'res = Select_Suga_Info(txtCode)
            'If res <= 0 Then
            '    MsgBox "검사번호가 존재하지 않습니다", vbExclamation
            '    txtCode.SetFocus
            '    Exit Sub
            'End If
            
        ElseIf res = 1 Then
            cmdSave.Caption = "수정"
            txtName = Trim(gReadBuf(2))
            'txtDec = Trim(gReadBuf(3))
            txtRefLow = Trim(gReadBuf(5))
            txtRefHigh = Trim(gReadBuf(6))
        End If
        
        txtName.SetFocus
    End If
End Sub

Private Sub txtPHigh_GotFocus()
    SelectFocus txtPHigh
End Sub

Private Sub txtPHigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDelta.SetFocus
    End If
End Sub

Private Sub txtPLow_GotFocus()
    SelectFocus txtPLow
End Sub

Private Sub txtPLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPHigh.SetFocus
    End If
End Sub

Private Sub txtRefhigh_GotFocus()
    SelectFocus txtRefHigh
End Sub

Private Sub txtRefhigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'txtPLow.SetFocus
        cmdSave.SetFocus
    End If
End Sub

'Private Sub Check2_GotFocus()
'    SelectFocus Check2
'End Sub

'Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        'txtPLow.SetFocus
'        cmdSave.SetFocus
'    End If
'End Sub



Private Sub txtRefLow_GotFocus()
    SelectFocus txtRefLow
End Sub

Private Sub txtRefLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRefHigh.SetFocus
    End If
End Sub

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

Private Sub cmbType_GotFocus()
    SelectFocus cmbType
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(cmbType.Text) = "" Then
            cmbType.SetFocus
            Exit Sub
        End If

        'txtDec.SetFocus
        'cmdSave.SetFocus
        
        Check1.SetFocus
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

        'txtDec.SetFocus
        'cmdSave.SetFocus
        
        cmbType.SetFocus
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

        'txtDec.SetFocus
        txtRefLow.SetFocus
        'txtRefLow.SetFocus
    End If
End Sub

Private Sub vasExam_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasExam, Col
    End If
End Sub

Private Sub vasExam_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or Row > vasExam.DataRowCnt Then Exit Sub
    
    txtCode = Trim(GetText(vasExam, Row, 1))
    txtName = Trim(GetText(vasExam, Row, 2))
    txtSeq = Trim(GetText(vasExam, Row, 4))
    Check1.Value = 1
    
    ssfrmExam.Visible = False
    
    cmdSave.SetFocus
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
'    txtDec = Trim(GetText(vasList, Row, 4))
    txtRefLow = Trim(GetText(vasList, Row, 5))
    txtRefHigh = Trim(GetText(vasList, Row, 6))
    txtPLow = Trim(GetText(vasList, Row, 7))
    txtPHigh = Trim(GetText(vasList, Row, 8))
    txtDelta = Trim(GetText(vasList, Row, 9))
    
    cmbType = Trim(GetText(vasList, Row, 12))
    
    If Trim(GetText(vasList, Row, 4)) = "1" Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
    
    If Trim(GetText(vasList, Row, 10)) = "1" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    txtSeq = Trim(GetText(vasList, Row, 11))
    
    cmdSave.Caption = "수정"
End Sub
