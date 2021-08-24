VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRemark 
   Caption         =   "Remark 설정"
   ClientHeight    =   5070
   ClientLeft      =   3510
   ClientTop       =   2925
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8805
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3570
      Left            =   9090
      TabIndex        =   16
      Top             =   1305
      Width           =   2310
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
         Left            =   1005
         TabIndex        =   18
         Top             =   1065
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtType 
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
         Left            =   1005
         TabIndex        =   17
         Top             =   1425
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
         Left            =   135
         TabIndex        =   20
         Top             =   1155
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "검사장비      CODE"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   135
         TabIndex        =   19
         Top             =   1470
         Visible         =   0   'False
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   5175
      TabIndex        =   3
      Top             =   675
      Width           =   3525
      Begin VB.TextBox txtEquipCode_TLA 
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
         TabIndex        =   30
         Top             =   630
         Width           =   900
      End
      Begin VB.TextBox txtEquipNo 
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
         TabIndex        =   28
         Top             =   1485
         Width           =   855
      End
      Begin VB.CheckBox chkMain 
         Height          =   330
         Left            =   3015
         TabIndex        =   26
         Top             =   180
         Width           =   285
      End
      Begin VB.Frame fraMo 
         Caption         =   "모검체 루트"
         Height          =   1230
         Left            =   135
         TabIndex        =   21
         Top             =   2340
         Width           =   3165
         Begin VB.TextBox txtRoot2 
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
            Left            =   960
            TabIndex        =   23
            Top             =   690
            Width           =   1215
         End
         Begin VB.TextBox txtRoot1 
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
            Left            =   960
            TabIndex        =   22
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "루 트 2"
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
            Left            =   135
            TabIndex        =   25
            Top             =   765
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "루 트 1"
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
            Left            =   135
            TabIndex        =   24
            Top             =   345
            Width           =   630
         End
      End
      Begin VB.TextBox txtEquipName 
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
         Left            =   1125
         TabIndex        =   11
         Top             =   1035
         Width           =   2115
      End
      Begin VB.TextBox txtValue 
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
         TabIndex        =   10
         Top             =   1935
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
         TabIndex        =   9
         Top             =   210
         Width           =   900
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         Height          =   495
         Left            =   180
         TabIndex        =   8
         Top             =   3600
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   495
         Left            =   990
         TabIndex        =   7
         Top             =   3600
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   1770
         TabIndex        =   6
         Top             =   3600
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         Height          =   495
         Left            =   2550
         TabIndex        =   5
         Top             =   3600
         Width           =   795
      End
      Begin VB.CheckBox chkDivision 
         Height          =   330
         Left            =   3015
         TabIndex        =   4
         Top             =   1485
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "TLA 코드"
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
         TabIndex        =   31
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비번호"
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
         TabIndex        =   29
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "메인장비"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   27
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장 비 명"
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
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "분 주 량"
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
         Top             =   1995
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "OCS 코드"
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
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '투명
         Caption         =   "자검체  생성 유무"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   12
         Top             =   1485
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   90
      ScaleHeight     =   525
      ScaleWidth      =   8565
      TabIndex        =   1
      Top             =   90
      Width           =   8595
      Begin Threed.SSPanel SSPanel1 
         Height          =   585
         Left            =   -435
         TabIndex        =   2
         Top             =   0
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "     Remark 설정"
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
   Begin FPSpread.vaSpread vasList 
      Height          =   4185
      Left            =   90
      TabIndex        =   0
      Top             =   750
      Width           =   4995
      _Version        =   393216
      _ExtentX        =   8811
      _ExtentY        =   7382
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
      MaxCols         =   10
      MaxRows         =   20
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEquipConfig.frx":0000
   End
End
Attribute VB_Name = "frmRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ClearText()
'화면초기화
    txtEquipCode = ""
    txtEquipName = ""
    txtValue = ""
    chkDivision = 0
    txtEquipNo = ""
    chkMain = 0
    
    If fraMo.Enabled = False Then
        fraMo.Enabled = True
        txtRoot1 = ""
        txtRoot2 = ""
    Else
        txtRoot1 = ""
        txtRoot2 = ""
    End If
    
    
    cmdSave.Caption = "저장"

End Sub

Sub DisplayList()
'검사항목 조회
    ClearSpread vasList

    SQL = "SELECT  EQUIPCODE, EQUIPCODE_TLA, EQUIPNAME, USEYN, JA_VALUES, ROOT1, ROOT2, EQUIPNUMBER, EQUIPMAIN " & CR & _
          "  FROM  Division " & CR & _
          " GROUP BY EQUIPCODE, EQUIPCODE_TLA, EQUIPNAME, USEYN, JA_VALUES, ROOT1, ROOT2, EQUIPNUMBER, EQUIPMAIN "
          
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Private Sub chkDivision_Click()
    If chkDivision.Value = 1 Then
        fraMo.Enabled = False
    Else
        fraMo.Enabled = True
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
    
    SQL = ""
    SQL = SQL & vbCrLf & "DELETE FROM DIVISION "
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & Trim(txtEquipCode) & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPCODE_TLA = '" & Trim(txtEquipCode_TLA) & "' "
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
    Dim lsEquipName As String
    Dim lsValue As Integer
    Dim lsType As String
    Dim lsEnable As String
    Dim Root1 As String
    Dim Root2 As String
    Dim EquipMain As String
    Dim EquipNum As String
    
    
    If Trim(txtEquipCode_TLA) = "" Then
        txtEquipCode_TLA.SetFocus
        MsgBox "TLA장비코드를 입력하세요", vbInformation
        Exit Sub
    End If
   
    lsEquipName = Trim(txtEquipName)

    If chkDivision.Value = 1 Then
        lsEnable = "Y"
    Else
        lsEnable = "N"
    End If
    
    
    If IsNumeric(txtValue) Then
        lsValue = CInt(txtValue)
    Else
        lsValue = 0
    End If
    
    If chkMain.Value = 1 Then
        EquipMain = "Y"
    Else
        EquipMain = "N"
    End If
    
    Root1 = txtRoot1
    Root2 = txtRoot2
    
    EquipNum = txtEquipNo
    
'    db_BeginTran gLocal
    'equipno, equipcode, examcode, examname, resprec, seqno, reflow, refhigh
    res = ExistOfEquipCode(Trim(txtEquipCode))
    If res = 1 Then
        SQL = ""
        SQL = SQL & vbCrLf & "UPDATE  DIVISION "
        SQL = SQL & vbCrLf & "   SET  EQUIPCODE = '" & txtEquipCode & "', "
        SQL = SQL & vbCrLf & "        EQUIPCODE_TLA = '" & Trim(txtEquipCode_TLA) & "', "
        SQL = SQL & vbCrLf & "        EQUIPNAME = '" & lsEquipName & "', "
        SQL = SQL & vbCrLf & "        USEYN = '" & lsEnable & "', "
        SQL = SQL & vbCrLf & "        JA_VALUES = '" & lsValue & "', "
        SQL = SQL & vbCrLf & "        EQUIPMAIN = '" & EquipMain & "', "
        SQL = SQL & vbCrLf & "        EQUIPNUMBER = '" & EquipNum & "', "
        SQL = SQL & vbCrLf & "        ROOT1 = '" & Root1 & "', "
        SQL = SQL & vbCrLf & "        ROOT2 = '" & Root2 & "' "
        SQL = SQL & vbCrLf & " WHERE  EQUIPCODE = '" & txtEquipCode & "' "
    ElseIf res = 0 Then
        SQL = ""
        SQL = SQL & vbCrLf & "INSERT INTO DIVISION "
        SQL = SQL & vbCrLf & "       (EQUIPCODE, EQUIPNAME, USEYN, "
        SQL = SQL & vbCrLf & "        EQUIPMAIN,EQUIPNUMBER, EQUIPCODE_TLA , "
        SQL = SQL & vbCrLf & "        JA_VALUES, ROOT1, ROOT2) "
        SQL = SQL & vbCrLf & "VALUES('" & Trim(txtEquipCode) & "', '" & lsEquipName & "', '" & lsEnable & "', "
        SQL = SQL & vbCrLf & "       '" & Trim(EquipMain) & "', '" & Trim(EquipNum) & "', '" & Trim(txtEquipCode_TLA) & "', "
        SQL = SQL & vbCrLf & "       '" & Trim(lsValue) & "', '" & Trim(Root1) & "', '" & Trim(Root2) & "')"
    End If

    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        SaveQuery SQL
        Exit Sub
    End If
    
    DisplayList
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    ClearText
    DisplayList
    frmInterface.tmrSearch.Enabled = False
    frmInterface.StatusBar1.Panels.Item(1) = "  중지 "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmInterface.tmrSearch.Enabled = True
    frmInterface.StatusBar1.Panels.Item(1) = "   "
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
        
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtEquipCode_TLA = Trim(GetText(vasList, Row, 2))
    txtEquipName = Trim(GetText(vasList, Row, 3))
    
    If Trim(GetText(vasList, Row, 4)) = "N" Then
        chkDivision.Value = 0
    Else
        chkDivision.Value = 1
    End If
    
    txtValue = Trim(GetText(vasList, Row, 5))
    If chkDivision.Value = 0 Then
        txtRoot1 = Trim(GetText(vasList, Row, 6))
        txtRoot2 = Trim(GetText(vasList, Row, 7))
    End If
    
    
    txtEquipNo = Trim(GetText(vasList, Row, 8))
    If Trim(GetText(vasList, Row, 9)) = "N" Or Trim(GetText(vasList, Row, 9)) = "" Then
        chkMain.Value = 0
    Else
        chkMain.Value = 1
    End If
    
    cmdSave.Caption = "수정"
    
End Sub


Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer
'장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT  EQUIPCODE, EQUIPNAME, USEYN, JA_VALUES, ROOT1, ROOT2 " & CR & _
          "  FROM  Division " & CR & _
          " WHERE  EQUIPCODE = '" & asEquipCode & "' " & CR & _
          " GROUP BY EQUIPCODE, EQUIPNAME, USEYN, JA_VALUES, ROOT1, ROOT2 "
    res = db_select_Col(gLocal, SQL)
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
        
    ExistOfEquipCode = 1
End Function
