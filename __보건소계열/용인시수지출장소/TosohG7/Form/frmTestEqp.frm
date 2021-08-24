VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmTestEqp 
   Caption         =   " 장비 VS 검사코드 설정"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   12090
   WindowState     =   2  '최대화
   Begin VB.TextBox txtTestNm 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1020
      Width           =   1425
   End
   Begin VB.TextBox txtTestCd 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   8910
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1020
      Width           =   1425
   End
   Begin VB.TextBox txtTstnmEqpDt 
      Appearance      =   0  '평면
      Enabled         =   0   'False
      Height          =   270
      Left            =   8910
      MaxLength       =   20
      TabIndex        =   5
      Top             =   685
      Width           =   1425
   End
   Begin VB.TextBox txtTstcdEqpDt 
      Appearance      =   0  '평면
      Enabled         =   0   'False
      Height          =   270
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   4
      Top             =   685
      Width           =   1425
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   -30
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":0000
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":059A
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Caption         =   " Instruments Test Item Link ."
      SubCaption      =   "검사실 검사항목과 장비 검사항목을 연결 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTstcdEqp 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   0
      Top             =   685
      Width           =   1425
   End
   Begin VB.TextBox txtTstnmEqp 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   45
      TabIndex        =   10
      Top             =   6450
      Width           =   11490
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Print"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   1515
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   2895
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   4260
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
   End
   Begin HSCotrol.CButton cmdEqpItm_Add 
      Height          =   300
      Left            =   3015
      TabIndex        =   3
      Top             =   675
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":0B34
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdEqpItm_Del 
      Height          =   300
      Left            =   3015
      TabIndex        =   9
      Top             =   1005
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Del"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":0C8E
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdAdd 
      Height          =   300
      Left            =   10980
      TabIndex        =   8
      Top             =   675
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":1228
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdDel 
      Height          =   300
      Left            =   10980
      TabIndex        =   19
      Top             =   1005
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Del"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp.frx":1382
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin VB.Frame Frame5 
      Height          =   6015
      Left            =   3990
      TabIndex        =   18
      Top             =   450
      Width           =   30
   End
   Begin FPSpread.vaSpread spdTestListDt 
      Height          =   4980
      Left            =   4260
      TabIndex        =   22
      Top             =   1410
      Width           =   7530
      _Version        =   196608
      _ExtentX        =   13282
      _ExtentY        =   8784
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestEqp.frx":191C
      UserResize      =   2
   End
   Begin FPSpread.vaSpread spdTestList 
      Height          =   4980
      Left            =   180
      TabIndex        =   24
      Top             =   1410
      Width           =   3660
      _Version        =   196608
      _ExtentX        =   6456
      _ExtentY        =   8784
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestEqp.frx":1C5C
      UserResize      =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "검사약어"
      Height          =   180
      Left            =   4800
      TabIndex        =   23
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "검사항목코드"
      Height          =   180
      Left            =   7665
      TabIndex        =   21
      Top             =   1065
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "장비검사명"
      Height          =   180
      Left            =   7860
      TabIndex        =   20
      Top             =   735
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "장비 검사명"
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   1065
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "장비코드(채널)"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   735
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "장비코드(채널)"
      Height          =   180
      Left            =   4290
      TabIndex        =   15
      Top             =   735
      Width           =   1230
   End
End
Attribute VB_Name = "frmTestEqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OBJTAG_EQP    As String = "EQP"
Private Const OBJTAG_TST    As String = "TST"
Private Const AUTO_VEFY     As String = "YES"
Private Const AUTO_VEFN     As String = "NO"

Private Const TLB_TEMP      As String = "TEMPTEABLE"
Private Const TLB_RESULT    As String = "INTERFACE003"

Private mAdoRs              As ADODB.Recordset
Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
        Case 0: Call cmdPrint
        Case 1: Call cmdSave
        Case 2: Call cmdClear
        Case 3: Call cmdClose
        Case Else
    End Select
    
End Sub

Private Sub cmdPrint()

    Call PrintFrom(lvwTestListLab.ListItems)

End Sub

Private Sub cmdAdd_Click()
    
    Dim iRow As Integer
    
    If Trim(txtTstcdEqpDt) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    If Trim(txtTstnmEqpDt) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    If Trim(txtTestCd) = "" Then
        Call ShowMessage("장비검사코드와 연결할 검사코드가 없습니다. 코드를 선택 하시오.   ")
        Exit Sub
    End If
    
    With spdTestListDt
        .Col = 1
        For iRow = 1 To .maxrows
            .Row = iRow
            If Trim(.Text) = "" Then
                .Col = 1: .Text = Trim(txtTstcdEqpDt)
                .Col = 2: .Text = Trim(txtTstnmEqpDt)
                .Col = 3: .Text = Trim(txtTestNm)
                .Col = 4: .Text = Trim(txtTestCd)
                    
                txtTstcdEqp = "":   txtTstnmEqp = ""
                txtTstcdEqpDt = "": txtTstnmEqpDt = "": txtTestCd = ""

                Exit Sub
            End If
            
            If Trim(.Text) = Trim(txtTstcdEqpDt) Then
                If vbYes = MsgBox(Trim(txtTstcdEqpDt) & " 장비검사 코드는 이미 있습니다. 바꾸시겠습니까?", vbExclamation + vbYesNo) Then
                    .Col = 1: .Text = Trim(txtTstcdEqpDt)
                    .Col = 2: .Text = Trim(txtTstnmEqpDt)
                    .Col = 3: .Text = Trim(txtTestNm)
                    .Col = 4: .Text = Trim(txtTestCd)
                    
                    txtTstcdEqp = "":   txtTstnmEqp = ""
                    txtTstcdEqpDt = "": txtTstnmEqpDt = "": txtTestCd = "": txtTestNm = ""

                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        Next iRow
        .maxrows = .maxrows + 1
        .Row = .maxrows
        .Col = 1: .Text = Trim(txtTstcdEqpDt)
        .Col = 2: .Text = Trim(txtTstnmEqpDt)
        .Col = 3: .Text = Trim(txtTestNm)
        .Col = 4: .Text = Trim(txtTestCd)
            
        txtTstcdEqp = "":   txtTstnmEqp = ""
        txtTstcdEqpDt = "": txtTstnmEqpDt = "": txtTestCd = "": txtTestNm = ""
    
    End With
    
End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub cmdClear()
    
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    txtTstcdEqpDt = ""
    txtTstnmEqpDt = ""
    txtTestCd = ""
    txtTestNm = ""

End Sub

Private Sub cmdDel_Click()
    Dim iRow        As Integer
    
    If Trim(txtTstcdEqpDt) = "" Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqpDt) = "" Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTestCd) = "" Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        txtTestCd.SetFocus
        Exit Sub
    End If
    
    With spdTestListDt
        .Col = 1
        For iRow = 1 To .maxrows
            .Row = iRow
            If Trim(.Text) = Trim(txtTstcdEqpDt) Then
                .DeleteRows iRow, 1
                .maxrows = .maxrows - 1
                
                txtTstcdEqp = ""
                txtTstnmEqp = ""
                txtTstcdEqpDt = ""
                txtTstnmEqpDt = ""
                txtTestCd = ""
                txtTestNm = ""
                txtTstcdEqp.SetFocus
                Exit Sub
            End If
        Next iRow
    End With
    
End Sub

Private Sub cmdEqpItm_Add_Click()
    Dim objEqpItem  As clsCommon
    Dim strTemp     As String
    
    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 입력 하시오.   ")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("장비 검사명이 없습니다. 검사명을 입력 하시오.   ")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
    Set objEqpItem = New clsCommon
    
    With objEqpItem
        .SetAdoCn AdoCn_Jet
        If .Let_EqpTestItem(INS_CODE, Trim(txtTstcdEqp), Trim(txtTstnmEqp)) Then
            txtTstcdEqp = ""
            txtTstnmEqp = ""
            txtTstcdEqpDt = ""
            txtTstnmEqpDt = ""
            txtTestCd = ""
            txtTstcdEqp.SetFocus
        Else
            Call ShowMessage("오류가있어 저장 하지 못했습니다.")
        End If
    End With
    
    Set objEqpItem = Nothing
    Call f_subSet_EqpData(INS_CODE)
    txtTstcdEqp.SetFocus

End Sub

Private Sub cmdEqpItm_Del_Click()
    Dim objEqpItem As clsCommon
    
    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하시오.")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
    Set objEqpItem = New clsCommon
    
    With objEqpItem
        .SetAdoCn AdoCn_Jet
        If .Del_EqpTestItem(INS_CODE, Trim(txtTstcdEqp)) Then
            txtTstcdEqp = ""
            txtTstnmEqp = ""
            txtTstcdEqpDt = ""
            txtTstnmEqpDt = ""
            txtTestCd = ""
            txtTstcdEqp.SetFocus
        Else
            Call ShowMessage("오류가있어 삭제 하지 못했습니다.")
        End If
    End With
    
    Set objEqpItem = Nothing
    Call f_subSet_EqpData(INS_CODE)
    
End Sub

Private Sub cmdSave()
    Dim iRow        As Integer
    Dim sqlDoc  As String, sqlRet   As Integer

    On Error GoTo frmTestEqp_Add_Error
    
    
    With spdTestListDt
        For iRow = 1 To .maxrows
            .Row = iRow
                               sqlDoc = "Update INTERFACE002"
            .Col = 2: sqlDoc = sqlDoc + "   set TESTNM_EQP = '" & Trim$(.Text) & "'," & _
                                        "       OUT_SEQ    = 0,"
            .Col = 4: sqlDoc = sqlDoc + "       TESTCD     = '" & Trim$(.Text) & "',"
            .Col = 3: sqlDoc = sqlDoc + "       TESTNM     = '" & Trim$(.Text) & "'," & _
                                        "       AUTOVERIFY = ''," & _
                                        "       REMARK     = ''," & _
                                        "       DELTA      = ''," & _
                                        "       DELTAGBN   = ''," & _
                                        "       PANICL     = ''," & _
                                        "       PANICH     = ''"
                      sqlDoc = sqlDoc + " where EQP_CD     = '" & INS_CODE & "'"
            .Col = 1: sqlDoc = sqlDoc + "   and TESTCD_EQP = '" & Trim$(.Text) & "'"
                     
            AdoCn_Jet.Execute sqlDoc, sqlRet
            If sqlRet = 0 Then
                                   sqlDoc = "Insert into INTERFACE002(" & _
                                            "            EQP_CD, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD," & _
                                            "            TESTNM, AUTOVERIFY, REMARK,   DELTA,   DELTAGBN," & _
                                            "            PANICL, PANICH)" & _
                                            "    values( '" & INS_CODE & "', "
                .Col = 1: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                .Col = 2: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                          sqlDoc = sqlDoc + "             0,"
                .Col = 3: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                .Col = 4: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                          sqlDoc = sqlDoc + "            '', '', '', '', '', '')"
                AdoCn_Jet.Execute sqlDoc, sqlRet
            End If
            
        Next
    End With
    
    Call f_subSet_EqpData(INS_CODE)

    Exit Sub
frmTestEqp_Add_Error:

    Call ErrMsgProc("frmTestEqp - Private Sub cmdSave()")

End Sub

Private Sub cmdSerch_Click()

    Dim objTestItem As clsCommon
    
    Set objTestItem = New clsCommon
    
    With objTestItem
        Call .SetAdoCn(AdoCn_SQL)
        Set mAdoRs = .Get_TestItem("")
    End With
    
    Set objTestItem = Nothing
    
    Call PopUp_List.ListItems.Clear
    If Not mAdoRs Is Nothing Then
        If Not mAdoRs.EOF Then
            Call DataLoadLvw(PopUp_List, vbCr, vbTab, mAdoRs.GetString)
            Call PopUp_List.ListItems.Remove(PopUp_List.ListItems.Count)
            
            With pnlTestitem
                .Visible = True
                .ZOrder
            End With
            PopUp_List.SetFocus
        End If
    Else
        Call ShowMessage("등록된 검사항목이 없습니다.")
    End If
    
    Set mAdoRs = Nothing
End Sub

Private Sub Form_Load()
    
    CaptionBar1.Caption = INS_NAME & " Instruments Test Item Link ."
    Call cmdClear
    Call f_subSet_EqpData(INS_CODE)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PopUp_List Is Nothing Then Set PopUp_List = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
'        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Private Sub Image1_DblClick()
    If lvwTstListEqp.Top > txtTstnmEqp.Top Then
        Call lvwTstListEqp.Move(Label1.left, CaptionBar1.Height, lvwTstListEqp.Width, ScaleHeight - (CaptionBar1.Height + fraCmdBar.Height + 30))
        txtTstcdEqp.Enabled = False
        txtTstnmEqp.Enabled = False
        cmdEqpItm_Add.Enabled = False
        cmdEqpItm_Del.Enabled = False
        Call lvwTstListEqp.ZOrder
    Else
        Call lvwTstListEqp.Move(Label1.left, lvwTestListLab.Top, lvwTstListEqp.Width, lvwTestListLab.Height)
        txtTstcdEqp.Enabled = True
        txtTstnmEqp.Enabled = True
        cmdEqpItm_Add.Enabled = True
        cmdEqpItm_Del.Enabled = True
    End If
End Sub

Private Sub f_subSet_EqpData(ByVal strEqp_Cd As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim iRow    As Integer, iRow1   As Integer
    
    sqlDoc = "select TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM," & _
             "       AUTOVERIFY, REMARK,     REFL,    REFH,   DELTA," & _
             "       DELTAGBN,   PANICL,     PANICH" & _
             "  from INTERFACE002" & _
             " where EQP_CD = '" & INS_CODE & "'" & _
             " order by TESTCD_EQP, TESTCD"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: iRow = 0
    Do While Not adoRS.EOF
        iRow = iRow + 1
        With spdTestList
            If iRow >= .maxrows Then .maxrows = .maxrows + 1
            .Row = iRow
            .SetText 1, iRow, Trim(adoRS("TESTCD_EQP"))
            .SetText 2, iRow, Trim(adoRS("TESTNM_EQP"))
            .SetText 3, iRow, Trim(adoRS("TESTNM"))
        End With
        
        With spdTestListDt
            If Trim(adoRS("TESTCD")) <> "" Then
                If iRow1 >= .maxrows Then .maxrows = .maxrows + 1
                .Row = iRow1
                .SetText 1, iRow, Trim(adoRS("TESTCD_EQP"))
                .SetText 2, iRow, Trim(adoRS("TESTNM_EQP"))
                .SetText 3, iRow, Trim(adoRS("TESTNM"))
                .SetText 4, iRow, Trim(adoRS("TESTCD"))
                iRow1 = iRow + 1
            End If
        End With
        
        adoRS.MoveNext
    Loop
    spdTestList.maxrows = spdTestList.maxrows - 1
'    spdTestListDt.maxrows = spdTestListDt.maxrows - 1
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subClear_Form()
    
    With spdTestList
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdTestListDt
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    txtTstcdEqpDt = ""
    txtTstnmEqpDt = ""
    txtTestCd = ""
    
End Sub


Private Sub PopUp_List_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PopUp_List_DblClick
        KeyAscii = 0
    End If
End Sub

Private Sub spdTestList_Click(ByVal Col As Long, ByVal Row As Long)

    With spdTestList
        .Row = Row
        .Col = 1
        If Trim(.Text) <> "" Then
            .Col = 1: txtTstcdEqp = Trim(.Text): txtTstcdEqpDt = Trim(.Text)
            .Col = 2: txtTstnmEqp = Trim(.Text): txtTstnmEqpDt = Trim(.Text)
            .Col = 3: txtTestNm = Trim(.Text)
            txtTestCd.Text = ""
            
            txtTestCd.SetFocus
        Else
            Exit Sub
        End If
    End With
    
End Sub

Private Sub spdTestListDt_Click(ByVal Col As Long, ByVal Row As Long)

    With spdTestListDt
        .Row = Row
        .Col = 1
        If Trim(.Text) <> "" Then
            .Col = 1: txtTstcdEqpDt = Trim(.Text)
            .Col = 2: txtTstnmEqpDt = Trim(.Text)
            .Col = 3: txtTestNm.Text = Trim(.Text)
            .Col = 4: txtTestCd.Text = Trim(.Text)
            txtTestCd.SetFocus
        Else
            Exit Sub
        End If
    End With
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
    
End Sub

Private Sub txtTestNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtTstcdEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtTstnmEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdEqpItm_Add_Click
        KeyAscii = 0
        Exit Sub
    End If
End Sub

