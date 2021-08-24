VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm602OnlineHelp 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사항목 On-Line 지침서"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   7140
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "128"
      Top             =   7215
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   1
      Left            =   5490
      TabIndex        =   38
      Top             =   1575
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 세부사항입력"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   0
      Left            =   1170
      TabIndex        =   37
      Top             =   1575
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 검사정보"
      LeftGab         =   100
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   5340
      Left            =   5490
      TabIndex        =   16
      Top             =   1830
      Width           =   4320
      Begin VB.TextBox txtTel 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1005
         TabIndex        =   1
         Top             =   2745
         Width           =   3150
      End
      Begin VB.TextBox txtRemark 
         Height          =   1770
         Left            =   135
         TabIndex        =   2
         Top             =   3510
         Width           =   4035
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   420
         Left            =   1005
         TabIndex        =   24
         Top             =   1155
         Width           =   3165
         Begin VB.OptionButton optDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "매일"
            Height          =   255
            Index           =   0
            Left            =   165
            TabIndex        =   26
            Top             =   120
            Width           =   1065
         End
         Begin VB.OptionButton optDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "선택"
            Height          =   255
            Index           =   1
            Left            =   1395
            TabIndex        =   25
            Top             =   120
            Width           =   1065
         End
      End
      Begin MedControls1.LisLabel lblSpcCd 
         Height          =   330
         Left            =   1005
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   330
         Left            =   1005
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblWorkarea 
         Height          =   330
         Left            =   1005
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   825
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DBE6E6&
         Height          =   840
         Left            =   1005
         TabIndex        =   27
         Top             =   1485
         Width           =   3165
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "월요일"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   34
            Top             =   150
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "화요일"
            Height          =   180
            Index           =   1
            Left            =   1035
            TabIndex        =   33
            Top             =   150
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "수요일"
            Height          =   180
            Index           =   2
            Left            =   1995
            TabIndex        =   32
            Top             =   150
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "목요일"
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   31
            Top             =   360
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "금요일"
            Height          =   180
            Index           =   4
            Left            =   1035
            TabIndex        =   30
            Top             =   360
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "토요일"
            Height          =   180
            Index           =   5
            Left            =   1995
            TabIndex        =   29
            Top             =   360
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "일요일"
            Height          =   180
            Index           =   6
            Left            =   75
            TabIndex        =   28
            Top             =   570
            Width           =   870
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   1005
         TabIndex        =   41
         Top             =   2250
         Width           =   3165
         Begin VB.OptionButton optTest 
            BackColor       =   &H00DBE6E6&
            Caption         =   "No"
            Height          =   180
            Index           =   1
            Left            =   1410
            TabIndex        =   43
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton optTest 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Yes"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   42
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.Label lblItemCd 
         BackStyle       =   0  '투명
         Caption         =   "예약항목    여부"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   8
         Left            =   90
         TabIndex        =   40
         Tag             =   "35215"
         Top             =   2340
         Width           =   930
      End
      Begin VB.Label lblItemCd 
         BackStyle       =   0  '투명
         Caption         =   "특이사항"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   90
         TabIndex        =   36
         Tag             =   "35215"
         Top             =   3255
         Width           =   930
      End
      Begin VB.Label lblItemCd 
         BackStyle       =   0  '투명
         Caption         =   "담당부서   연락처"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   105
         TabIndex        =   35
         Tag             =   "35215"
         Top             =   2775
         Width           =   930
      End
      Begin VB.Label lblItemCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일시"
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
         Index           =   5
         Left            =   105
         TabIndex        =   23
         Tag             =   "35215"
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label lblItemCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "담당부서"
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
         Index           =   4
         Left            =   105
         TabIndex        =   21
         Tag             =   "35215"
         Top             =   930
         Width           =   780
      End
      Begin VB.Label lblItemCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체명"
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
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Tag             =   "35215"
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblItemCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체코드"
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
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Tag             =   "35215"
         Top             =   225
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   990
      Left            =   1170
      TabIndex        =   9
      Top             =   1830
      Width           =   4305
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   930
         TabIndex        =   0
         Top             =   150
         Width           =   1440
      End
      Begin VB.CommandButton cmdPopupList 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2385
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frm602OnlineHelp.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   150
         Width           =   300
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00EBEBEB&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2715
         Style           =   1  '그래픽
         TabIndex        =   12
         Tag             =   "124"
         Top             =   150
         Width           =   750
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00EBEBEB&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3465
         Style           =   1  '그래픽
         TabIndex        =   11
         Tag             =   "124"
         Top             =   150
         Width           =   750
      End
      Begin MedControls1.LisLabel lblTestName 
         Height          =   330
         Left            =   930
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   555
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.Label lblItemCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사명"
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
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Tag             =   "35215"
         Top             =   615
         Width           =   585
      End
      Begin VB.Label lblItemCd 
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
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Tag             =   "35215"
         Top             =   195
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   8475
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "128"
      Top             =   7200
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   5775
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "128"
      Top             =   7215
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   2
      Left            =   1170
      TabIndex        =   39
      Top             =   2820
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 검체정보"
      LeftGab         =   100
   End
   Begin VB.Frame fraSpcList 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   1170
      TabIndex        =   6
      Tag             =   "35209"
      Top             =   3060
      Width           =   4305
      Begin FPSpread.vaSpread tblSpcList 
         Height          =   3000
         Left            =   120
         TabIndex        =   7
         Tag             =   "35220"
         Top             =   780
         Width           =   4080
         _Version        =   196608
         _ExtentX        =   7197
         _ExtentY        =   5292
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   9
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "frm602OnlineHelp.frx":058A
         VirtualRows     =   7
      End
      Begin VB.Label Label22 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체리스트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00DDF0F5&
         FillStyle       =   0  '단색
         Height          =   390
         Index           =   0
         Left            =   90
         Shape           =   4  '둥근 사각형
         Top             =   255
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm602OnlineHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private Const Indicator = "▶"

Private Sub cmdClear_Click()
    Call ClearData
End Sub
Private Sub ClearData(Optional ByVal blnClear As Boolean = False)
    Dim ii As Integer
    
    If blnClear = False Then
        txtTestCd.Text = "":   lblTestName.Caption = ""
        Call medClearTable(tblSpcList)
    End If
    
    lblSpcCd.Caption = "": lblSpcNm.Caption = ""
    txtRemark.Text = "": txtTel.Text = ""
    lblWorkarea.Caption = ""
    optTest(1).Value = True
'    optDiv(0).Value = True
    For ii = 0 To 6
        If optDiv(0).Value Then
            chkDay(ii).Value = 1
        Else
            chkDay(ii).Value = 0
        End If
    Next
    
    
End Sub
Private Sub cmdExit_Click()
    Unload Me
    Set objCodeList = Nothing
End Sub

Private Sub cmdPopupList_Click()
    Dim tmpSql As String

    Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = dbconn
        .FormCaption = "Test Code List.."
        .Tag = "TestCode"
        .ColumnHeaderText = "검사코드;검사명"
        tmpSql = SqlLAB001CodeList
        Call .LoadPopUp(tmpSql) ' (, Me.Top + txtTestCd.Top + txtTestCd.Height, Me.Left + txtTestCd.Left, lstItemList)
    End With
End Sub



Private Sub Form_Activate()
    Call ClearData
End Sub

Private Sub Form_Load()
    Dim MyItem As New clsItem
    Call MyItem.GetItemList(lstItemList): DoEvents

End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    If pSelectedItem = "" Then Exit Sub
    DoEvents
    txtTestCd.Text = medShift(pSelectedItem, ";")
    lblTestName.Caption = medShift(pSelectedItem, ";")
    
    Call txtTestCd_KeyPress(vbKeyReturn)

    Me.Enabled = True
End Sub

Private Sub optDiv_Click(Index As Integer)
    Dim lngValue As Integer
    Dim ii       As Integer
    
    lngValue = 0
    Frame3.Enabled = False
    If Index = 0 Then
        lngValue = 1
    Else
        Frame3.Enabled = True
    End If
    For ii = 0 To 6
        chkDay(ii).Value = lngValue
    Next

End Sub
Private Function SqlLAB001CodeList() As String
    SqlLAB001CodeList = " Select a.testcd, a.abbrnm10 as testnm " & _
                        " From  " & T_LAB001 & " a " & _
                        " Where a.applydt = ( select max(applydt) from " & T_LAB001 & "  " & _
                        "                              where testcd = a.testcd ) " & _
                        " Order by a.testcd "
End Function
'Private Sub objCodeList_SendCode(ByVal SelString As String)
'    If SelString = "" Then Exit Sub
'    DoEvents
'    txtTestCd.Text = medShift(SelString, ";")
'    lblTestName.Caption = medShift(SelString, ";")
'
'    Call txtTestCd_KeyPress(vbKeyReturn)
'
'    Me.Enabled = True
'
'End Sub
Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objCodeList Is Nothing Then Call cmdPopupList_Click
'        Call objCodeList.SetFocus(2)
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    Dim tmpSql As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtTestCd.Text = "" Then Exit Sub
    
        If lstItemList.Exists(Trim(txtTestCd.Text)) Then
            lstItemList.KeyChange (Trim(txtTestCd.Text))
            lblTestName.Caption = lstItemList.Fields("testnm")
            
        Else
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        Call LabSpecimenLoad(SqlSpecimenRead(txtTestCd.Text))
        If tblSpcList.MaxRows > 0 Then
            Call tblSpcList_Click(1, 1)
        End If
    End If
    
End Sub
Private Function SqlSpecimenRead(ByVal TestCd As String) As String
    Dim SSQL As String
    
    SSQL = " select a.workarea,b.spccd,c.field3 as spcnm,d.field3 as workareanm,b.seq " & _
           " from " & T_LAB001 & " a," & T_LAB004 & " b," & T_LAB032 & " c," & T_LAB032 & " d" & _
           " where  " & DBW("a.testcd=", TestCd) & _
           " and a.applydt=(select max(applydt) from s2lab001 where testcd=a.testcd)" & _
           " and (a.expdt='' or a.expdt is null)" & _
           " and a.testcd=b.testcd" & _
           " and " & DBW("c.cdindex=", LC3_Specimen) & _
           " and c.cdval1=b.spccd" & _
           " and b.applydt=(select max(applydt) from s2lab004 where testcd=b.testcd and spccd=b.spccd)" & _
           " and (b.expdt='' or b.expdt is null)" & _
           " and " & DBW("d.cdindex=", LC3_WorkArea) & _
           " and d.cdval1=a.workarea" & _
           " order by seq"
    SqlSpecimenRead = SSQL
End Function

'% Sub Routine 3 : LabSpecimenLoad
'%                 지정검체명들을 Tab에 Display

Private Sub LabSpecimenLoad(ByVal SqlStmt As String)

    Dim objRs As Recordset
    Dim i As Integer

    Set objRs = New Recordset   'Sql 실행
    objRs.Open SqlStmt, dbconn

    With tblSpcList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxRows = 0
    
        .Row = 0
        While (objRs.EOF = False)
            If .Row = .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .Row + 1
            .Col = 2: .Value = Trim("" & objRs.Fields("seq").Value)
            .Col = 3: .Value = Trim("" & objRs.Fields("spccd").Value)
            .Col = 4: .Value = Trim("" & objRs.Fields("spcnm").Value)
            .Col = 5: .Value = Trim("" & objRs.Fields("workarea").Value & Space(5) & objRs.Fields("workareanm").Value & "")
            objRs.MoveNext
        Wend
        .RowHeight(-1) = 13.3
    End With

    Set objRs = Nothing

End Sub

'% 검체리스트에서 한 검체를 선택(클릭)하면 상세 정보 Display

Private Sub tblSpcList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim RS       As Recordset
    Dim SSQL     As String
    Dim aryTmp() As String
    Dim ii       As Integer
    
    If tblSpcList.DataRowCnt < 1 Then Exit Sub
    If Row < 1 Then Exit Sub
    
    Call ClearData(True)
    
    With tblSpcList
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = 1
        .BlockMode = True
        .Value = ""
        .BlockMode = False
        .Row = Row: .Col = 1
        .Value = Indicator: .ForeColor = &HFF&
        .Col = 3: lblSpcCd.Caption = .Value
        .Col = 4: lblSpcNm.Caption = .Value
        .Col = 5: lblWorkarea.Caption = .Value
    End With
    If lblSpcCd.Caption = "" Then Exit Sub
    Set RS = New Recordset
    
    
    SSQL = " select * from " & T_LAB031 & _
           " where " & _
                     DBW("cdindex=", LC4_TestItemComment) & _
           " and " & DBW("cdval1=", txtTestCd.Text) & _
           " and " & DBW("cdval2=", lblSpcCd.Caption)
    
    RS.Open SSQL, dbconn
    If Not RS.EOF Then
        
        If Trim(RS.Fields("field4").Value & "") <> "" Then
            If RS.Fields("field4").Value & "" = "7" Then
                optDiv(0).Value = True
            Else
                optDiv(1).Value = True
                aryTmp = Split(RS.Fields("field4").Value & "", COL_DIV)
                For ii = 0 To UBound(aryTmp)
                    chkDay(aryTmp(ii)).Value = 1
                Next
            End If
        End If
        txtRemark.Text = RS.Fields("text1").Value & ""
        txtTel.Text = RS.Fields("field5").Value & ""
        optTest(0).Value = IIf(RS.Fields("text2").Value & "" = "1", True, False)
    End If
    Set RS = Nothing
    
End Sub

Private Sub cmdSave_Click()
    Dim RS        As Recordset
    Dim strField4 As String
    Dim SSQL      As String
    Dim strTest   As String
    Dim sTestNm   As String
    Dim sSpcNm    As String
    Dim ii        As Integer
    
    
    If txtTestCd.Text = "" Then Exit Sub
    If lblSpcCd.Caption = "" Then Exit Sub
    
    strTest = "0"
    If optTest(0).Value Then strTest = "1"
    
    If optDiv(0).Value Then
        strField4 = "7"
    Else
        For ii = 0 To 6
            If chkDay(ii).Value = 1 Then
                If strField4 = "" Then
                    strField4 = ii
                Else
                    strField4 = strField4 & COL_DIV & ii
                End If
            End If
        Next
    End If
    
    If lstItemList.Exists(txtTestCd.Text) Then
        lstItemList.KeyChange txtTestCd.Text
        sTestNm = lstItemList.Fields("testnm")
    End If
    
    
    On Error GoTo SAVE_ERROR
    dbconn.BeginTrans
    
    SSQL = "delete from " & T_LAB031 & " where " & _
                DBW("cdindex=", LC4_TestItemComment) & _
         " and " & DBW("cdval1=", txtTestCd.Text) & _
         " and " & DBW("cdval2=", lblSpcCd.Caption)
    dbconn.Execute SSQL
    
    SSQL = " insert into " & T_LAB031 & " (cdindex,cdval1,cdval2,field1,field2,field3,field4,field5,text1,text2) " & _
         " values (" & _
        DBV("cdindex", LC4_TestItemComment, 1) & DBV("cdval1", txtTestCd.Text, 1) & DBV("cdval2", lblSpcCd.Caption, 1) & _
        DBV("field1", lblTestName.Caption, 1) & DBV("field2", lblSpcNm.Caption, 1) & DBV("field3", lblWorkarea.Caption, 1) & _
        DBV("field4", strField4, 1) & DBV("field5", txtTel.Text, 1) & DBV("text1", txtRemark.Text, 1) & DBV("text2", strTest) & _
        " )"
    dbconn.Execute SSQL
    dbconn.CommitTrans
    Exit Sub
    
SAVE_ERROR:
    dbconn.RollbackTrans

End Sub
Private Sub cmdFind_Click(Index As Integer)
    Dim i As Integer
    
    If txtTestCd.Text = "" Then Exit Sub

    If Not lstItemList.Exists(txtTestCd.Text) Then Exit Sub
    Call lstItemList.KeyChange(txtTestCd.Text)

    Select Case Index
        Case 0:
            lstItemList.MovePrevious
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
        Case 1:
            lstItemList.MoveNext
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
    End Select
    Call txtTestCd_KeyPress(vbKeyReturn)
End Sub
