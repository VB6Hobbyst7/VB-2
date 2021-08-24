VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOrderCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   6240
   ClientLeft      =   1170
   ClientTop       =   1515
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10455
   Begin FPSpread.vaSpread vasList 
      Height          =   5535
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   6570
      _Version        =   393216
      _ExtentX        =   11589
      _ExtentY        =   9763
      _StockProps     =   64
      ColHeaderDisplay=   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   525
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   10335
      _Version        =   65536
      _ExtentX        =   18230
      _ExtentY        =   926
      _StockProps     =   15
      Caption         =   "       Coagu Check 장비 코드 설정"
      ForeColor       =   4194304
      BackColor       =   16774393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
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
      Height          =   5535
      Left            =   6675
      TabIndex        =   1
      Top             =   630
      Width           =   3735
      Begin VB.CommandButton cmdClose 
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
         Left            =   2670
         TabIndex        =   15
         Top             =   4680
         Width           =   825
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
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
         Left            =   1830
         TabIndex        =   14
         Top             =   4680
         Width           =   825
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
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
         Left            =   990
         TabIndex        =   13
         Top             =   4680
         Width           =   825
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
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
         Left            =   150
         TabIndex        =   12
         Top             =   4680
         Width           =   825
      End
      Begin VB.TextBox txtSubCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   36
         Top             =   1350
         Width           =   2235
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
         Left            =   1500
         TabIndex        =   34
         Top             =   4410
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboOrdGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmOrderCode.frx":1D35
         Left            =   1410
         List            =   "frmOrderCode.frx":1D45
         TabIndex        =   32
         Top             =   4680
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.CheckBox chkUse 
         Height          =   285
         Left            =   1635
         TabIndex        =   31
         Top             =   4035
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtSeq 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   28
         Top             =   2070
         Width           =   765
      End
      Begin VB.TextBox txtMuch 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Top             =   1710
         Width           =   2235
      End
      Begin VB.TextBox txtDec 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   7
         Top             =   2430
         Width           =   2235
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   990
         Width           =   2235
      End
      Begin VB.TextBox txtEquipCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   615
         Width           =   2235
      End
      Begin VB.TextBox txtRefLow 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   17
         Top             =   2790
         Width           =   885
      End
      Begin VB.TextBox txtRefHigh 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2610
         TabIndex        =   19
         Top             =   2790
         Width           =   885
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
         Left            =   1410
         TabIndex        =   21
         Top             =   5040
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
         Left            =   2790
         TabIndex        =   23
         Top             =   5040
         Visible         =   0   'False
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
         TabIndex        =   25
         Top             =   4770
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "서브코드"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1425
         Width           =   780
      End
      Begin VB.Label Label15 
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
         Left            =   330
         TabIndex        =   35
         Top             =   4485
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label14 
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
         Left            =   210
         TabIndex        =   33
         Top             =   4530
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
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
         Height          =   225
         Left            =   450
         TabIndex        =   30
         Top             =   4080
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    서"
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
         Left            =   240
         TabIndex        =   29
         Top             =   2145
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비구분"
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
         Left            =   240
         TabIndex        =   10
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1785
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "정 확 도"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2505
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
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
         Left            =   240
         TabIndex        =   2
         Top             =   690
         Width           =   780
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
         TabIndex        =   27
         Top             =   4845
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
         TabIndex        =   26
         Top             =   4845
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
         TabIndex        =   24
         Top             =   5115
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
         Left            =   240
         TabIndex        =   22
         Top             =   5115
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
         Left            =   2310
         TabIndex        =   20
         Top             =   2835
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보고범위"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2865
         Width           =   780
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
    cboOrdGubun.Text = ""
    txtGubun = ""
    txtRefLow = ""
    txtRefHigh = ""
    txtPLow = ""
    txtPHigh = ""
    txtDelta = ""
    txtSeq = ""
    txtDec = ""
    txtSubCode = ""
    
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
    '장비코드,검사코드,검사명,정확도,오더구분,참고치_Low,참고치_High,패닉_Low,패닉_High,델타,검사구분,순번
    SQL = "SELECT equipcode, examcode, examname, resprec, OrdGubun, reflow, refhigh, paniclow, panichigh, deltavalue, examflag, seqno, rsgubun, subcode " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " Order by seqno, equipcode "
          
    db_select_Vas gLocal, SQL, vasList
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode_1(asEquipCode As String, Optional asSuga As String = "", Optional asSubCode As String = "") As Integer
'장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode_1 = -1
    
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
    If Trim(asSuga) <> "" Then
        SQL = SQL & CR & _
          "   AND subcode = '" & asSubCode & "' "
    End If
    Res = db_select_Col(gLocal, SQL)
    If Res = 0 Then
        ExistOfEquipCode_1 = 0
        Exit Function
    ElseIf Res = -1 Then
        ExistOfEquipCode_1 = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode_1 = 1
End Function


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
    Res = db_select_Col(gLocal, SQL)
    If Res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf Res = -1 Then
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
    
    Connect_Server_Neosoft
    
    SQL = " Select LABM_ID, LABM_NAME " & CR & _
          " from CC_LABM " & CR & _
          " where LABM_ID = '" & Trim(asSuga) & "' "

    Res = db_select_Col_Neo(gServer, SQL)
    
'    If cn_Server_Flag Then DisConnect_Server
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    ElseIf Res = 0 Then
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

Private Sub cboOrdGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboOrdGubun.ListIndex < 0 Then
            cboOrdGubun.SetFocus
            Exit Sub
        End If
        
        txtSeq.SetFocus
    End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'txtSeq.SetFocus
        cboOrdGubun.SetFocus
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
          "  and examcode = '" & Trim(txtCode) & "' " & CR & _
          "  and subcode = '" & Trim(txtSubCode) & "' "
    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        db_RollBack gLocal
        Exit Sub
    End If
    
    db_Commit gLocal

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
    Dim lsFlag As String
    Dim liSeqNo As Integer
    
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
    
    '소수점
    If Trim(txtDec) = "" Then
        txtDec.Text = 1
'        txtDec.SetFocus
'        Exit Sub
    End If
    
    '순번
    If IsNumeric(txtSeq) Then
        liSeqNo = CInt(txtSeq)
    Else
        liSeqNo = 0
    End If
    
    '사용여부
    If chkUse.Value = 1 Then
        lsFlag = "1"
    Else
        lsFlag = "0"
    End If
    
    db_BeginTran gLocal
    'examcode, examname, resprec, refmlow, refmhigh, refwlow, refwhigh
    Res = ExistOfEquipCode_1(Trim(txtEquipCode), Trim(txtCode), Trim(txtSubCode))
    If Res = 1 Then
        SQL = "Update equipexam " & CR & _
              "Set resprec = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    examname = '" & Trim(txtName) & "', " & vbCrLf & _
              "    rsgubun = '" & Trim(txtGubun.Text) & "', " & CR & _
              "    reflow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    refhigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    paniclow = '" & Trim(txtPLow) & "', " & vbCrLf & _
              "    panichigh = '" & Trim(txtPHigh) & "', " & vbCrLf & _
              "    deltavalue = '" & Trim(txtDelta) & "', " & vbCrLf & _
              "    examflag = " & lsFlag & ", " & vbCrLf & _
              "    seqno = " & liSeqNo & " " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and subcode = '" & Trim(txtSubCode) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(txtCode) & "' "
    ElseIf Res = 0 Then
        SQL = "Insert Into equipexam (equipno,equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue, examflag, seqno, rsgubun, subcode) " & CR & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & CR & _
              "        '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "', '" & Trim(txtPLow) & "', '" & Trim(txtPHigh) & "', '" & Trim(txtDelta) & "', " & CR & _
              "        " & lsFlag & ", " & liSeqNo & ", '" & Trim(txtGubun.Text) & "', '" & Trim(txtSubCode) & "' ) "
    End If

    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        db_RollBack gLocal
        SaveQuery SQL
        Exit Sub
    End If
    
    db_Commit gLocal
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
            
    ClearText
    
    DisplayList
    
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

Private Sub txtsubcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtsubcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSubCode = UCase(txtSubCode)
        Res = ExistOfEquipCode_1(Trim(txtEquipCode), Trim(txtCode), Trim(txtSubCode))
        If Res = -1 Then
            txtSubCode.SetFocus
            Exit Sub
        ElseIf Res = 0 Then
            cmdSave.Caption = "저장"
'            res = Select_Suga_Info(txtCode)
'            If res <= 0 Then
'                MsgBox "검사번호가 존재하지 않습니다", vbExclamation
'                txtCode.SetFocus
'                Exit Sub
'            End If

        ElseIf Res = 1 Then
            cmdSave.Caption = "수정"
            txtName = Trim(gReadBuf(2))
            'txtDec = Trim(gReadBuf(3))
            txtRefLow = Trim(gReadBuf(5))
            txtRefHigh = Trim(gReadBuf(6))
        End If
        
        txtName.SetFocus
    End If
End Sub

Private Sub txtcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCode = UCase(txtCode)
'        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
'        If res = -1 Then
'            txtCode.SetFocus
'            Exit Sub
'        ElseIf res = 0 Then
'            cmdSave.Caption = "저장"
''            res = Select_Suga_Info(txtCode)
''            If res <= 0 Then
''                MsgBox "검사번호가 존재하지 않습니다", vbExclamation
''                txtCode.SetFocus
''                Exit Sub
''            End If
'
'        ElseIf res = 1 Then
'            cmdSave.Caption = "수정"
'            txtName = Trim(gReadBuf(2))
'            'txtDec = Trim(gReadBuf(3))
'            txtRefLow = Trim(gReadBuf(5))
'            txtRefHigh = Trim(gReadBuf(6))
'        End If
        
        txtSubCode.SetFocus
    End If
End Sub


Private Sub txtGubun_GotFocus()
    SelectFocus txtGubun
End Sub

Private Sub txtGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDec.SetFocus
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

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtName.Text) = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        
        txtSeq.SetFocus
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
        
        txtDec.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
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
    txtSubCode = Trim(GetText(vasList, Row, 14))
    
    '오더구분
    cboOrdGubun.Text = ""
    For i = 0 To cboOrdGubun.ListCount - 1
        If Left(cboOrdGubun.List(i), 1) = Trim(GetText(vasList, Row, 5)) Then
            cboOrdGubun.ListIndex = i
            Exit For
        End If
    Next i
    
    txtRefLow = Trim(GetText(vasList, Row, 6))
    txtRefHigh = Trim(GetText(vasList, Row, 7))
    txtPLow = Trim(GetText(vasList, Row, 8))
    txtPHigh = Trim(GetText(vasList, Row, 9))
    txtDelta = Trim(GetText(vasList, Row, 10))
    
    '검사여부
    If Trim(GetText(vasList, Row, 11)) = "1" Then
        chkUse.Value = 1
    Else
        chkUse.Value = 0
    End If
    
    '순번
    txtSeq = Trim(GetText(vasList, Row, 12))
    
    txtGubun = Trim(GetText(vasList, Row, 13))
    
    cmdSave.Caption = "수정"
End Sub
