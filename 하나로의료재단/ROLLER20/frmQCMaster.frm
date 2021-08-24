VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmQCMaster 
   Caption         =   "QC 설정"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   17445
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdDSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   3750
      Width           =   735
   End
   Begin VB.CommandButton cmdHSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4590
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdQCPathSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   9960
      TabIndex        =   11
      Top             =   7260
      Width           =   645
   End
   Begin VB.TextBox txtQCPath 
      Height          =   315
      Left            =   1950
      TabIndex        =   10
      Top             =   7260
      Width           =   7935
   End
   Begin VB.CommandButton cmdDAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   3750
      Width           =   345
   End
   Begin VB.CommandButton cmdHAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4230
      TabIndex        =   5
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   525
      Left            =   15960
      TabIndex        =   0
      Top             =   7200
      Width           =   1275
   End
   Begin FPSpread.vaSpread spdHeader 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   5175
      _Version        =   393216
      _ExtentX        =   9128
      _ExtentY        =   5530
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayRowHeaders=   0   'False
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
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmQCMaster.frx":0000
   End
   Begin FPSpread.vaSpread spdDetail 
      Height          =   6465
      Left            =   5370
      TabIndex        =   3
      Top             =   510
      Width           =   11895
      _Version        =   393216
      _ExtentX        =   20981
      _ExtentY        =   11404
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmQCMaster.frx":062E
   End
   Begin FPSpread.vaSpread spdQCID 
      Height          =   2835
      Left            =   120
      TabIndex        =   4
      Top             =   4140
      Width           =   5175
      _Version        =   393216
      _ExtentX        =   9128
      _ExtentY        =   5001
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmQCMaster.frx":0CDD
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "QC결과 저장경로"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   210
      TabIndex        =   9
      Top             =   7320
      Width           =   1545
   End
   Begin VB.Label Label2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "상세정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   780
   End
   Begin VB.Label lblQCID 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "QC ID 리스트 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   3870
      Width           =   1305
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "물질 리스트 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   1125
   End
End
Attribute VB_Name = "frmQCMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDAdd_Click()
    
    With spdQCID
        .MaxRows = .MaxRows + 1
        Call SetText(spdQCID, mGetP(lblQCID.Caption, 2, ","), spdQCID.MaxRows, 1)
    End With

End Sub

Private Sub cmdDSave_Click()

    Call SetQCList_Detail(mGetP(lblQCID.Caption, 2, ","))
    
End Sub

Private Sub cmdHAdd_Click()
    
    With spdHeader
        .MaxRows = .MaxRows + 1
        
    End With
    
End Sub

Private Sub cmdHSave_Click()

    Call SetQCList_Header
    
End Sub

Private Sub cmdQCPathSave_Click()
    
    Call WritePrivateProfileString("HOSP", "QCPATH", txtQCPath.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub


Private Sub Form_Load()
    
    Call frmClear

    Call GetQCList_Header

End Sub


Private Sub frmClear()

    spdHeader.MaxRows = 0
    spdDetail.MaxRows = 0
    spdQCID.MaxRows = 0
    
    txtQCPath.Text = gHOSP.QCPATH
    
End Sub

Private Sub spdHeader_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strInst     As String
    
    If Row = 0 Then
        Exit Sub
    End If
    
    strInst = Trim(GetText(spdHeader, Row, 4))
    lblQCID.Caption = "QC ID 리스트," & strInst
    
    '-- 673018,
    Call GetQCList_Detail(Trim(GetText(spdHeader, Row, 1)), Trim(GetText(spdHeader, Row, 2)), Trim(GetText(spdHeader, Row, 3)))

    '-- 673018
    Call GetQCList_QCID(strInst)

End Sub

Private Sub spdHeader_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strLot  As String
    Dim iRow    As Integer
    
    iRow = spdHeader.ActiveRow
    
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > spdHeader.DataRowCnt Then
            Exit Sub
        End If
        strLot = Trim(GetText(spdHeader, iRow, 2))

        If MsgBox(strLot & " 를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM QCHEADER " & vbCr
        SQL = SQL & " WHERE LOTID = '" & strLot & "' " & vbCr
        
        Call DBExec(AdoCn_Local, SQL)
        
        Call spdHeader.DeleteRows(iRow, 1)
        spdHeader.MaxRows = spdHeader.MaxRows - 1
        
    End If
    
End Sub



Private Sub spdQCID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strQCID  As String
    Dim iRow    As Integer
    
    iRow = spdQCID.ActiveRow
    
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > spdQCID.DataRowCnt Then
            Exit Sub
        End If
        strQCID = Trim(GetText(spdQCID, iRow, 3))

        If MsgBox(strQCID & " 를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM QCDETAIL " & vbCr
        SQL = SQL & " WHERE ID = '" & strQCID & "' " & vbCr
        
        Call DBExec(AdoCn_Local, SQL)
        
        Call spdQCID.DeleteRows(iRow, 1)
        spdQCID.MaxRows = spdQCID.MaxRows - 1
        
    End If
    
End Sub
