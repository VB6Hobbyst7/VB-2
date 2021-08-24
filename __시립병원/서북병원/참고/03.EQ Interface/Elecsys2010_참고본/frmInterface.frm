VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   " Elecsys 2010 Interface Program"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15225
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows 기본값
   Begin IF_Elecsys2010_진주의료원.MDButton btnClear 
      Height          =   675
      Left            =   12810
      TabIndex        =   34
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
   End
   Begin IF_Elecsys2010_진주의료원.MDButton btnClose 
      Height          =   675
      Left            =   13965
      TabIndex        =   35
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "종료"
   End
   Begin IF_Elecsys2010_진주의료원.MDButton btnSetup 
      Height          =   675
      Left            =   11655
      TabIndex        =   33
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "코드설정"
   End
   Begin IF_Elecsys2010_진주의료원.MDButton btnConfig 
      Height          =   675
      Left            =   10500
      TabIndex        =   32
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "통신설정"
   End
   Begin IF_Elecsys2010_진주의료원.MDButton btnTrans 
      Height          =   675
      Left            =   9345
      TabIndex        =   31
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "선택전송"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   5085
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtBuff 
      Height          =   990
      Left            =   1125
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5010
      Visible         =   0   'False
      Width           =   3645
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   375
      Left            =   300
      TabIndex        =   29
      Top             =   1140
      Width           =   465
      _Version        =   65536
      _ExtentX        =   820
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "순번"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.74
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin VB.CommandButton cmdUp 
      Height          =   525
      Left            =   330
      Picture         =   "frmInterface.frx":0442
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   9750
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdDown 
      Height          =   525
      Left            =   1080
      Picture         =   "frmInterface.frx":0571
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   9750
      Visible         =   0   'False
      Width           =   705
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   9255
      Left            =   8100
      TabIndex        =   23
      Top             =   1110
      Width           =   7005
      _Version        =   196613
      _ExtentX        =   12356
      _ExtentY        =   16325
      _StockProps     =   64
      ColHeaderDisplay=   1
      EditModePermanent=   -1  'True
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
      GridColor       =   16777215
      MaxCols         =   13
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmInterface.frx":06A3
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   900
      TabIndex        =   24
      Top             =   1230
      Width           =   165
   End
   Begin VB.TextBox txtMsg 
      ForeColor       =   &H000000C0&
      Height          =   765
      Left            =   1050
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   4050
      Visible         =   0   'False
      Width           =   6345
   End
   Begin Threed.SSPanel sspMode 
      Height          =   675
      Left            =   8520
      TabIndex        =   21
      Top             =   120
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "전송모드"
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.26
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin VB.CommandButton cmd_Trans 
      Caption         =   "선택전송"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9360
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12810
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   13950
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "코드설정"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11670
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "통신설정"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10530
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   180
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   4096
      InputLen        =   1
      RThreshold      =   1
      EOFEnable       =   -1  'True
   End
   Begin VB.TextBox txtUID 
      Appearance      =   0  '평면
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6870
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtToday 
      Appearance      =   0  '평면
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6870
      TabIndex        =   11
      Text            =   "2002/02/18"
      Top             =   120
      Width           =   1455
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   150
      TabIndex        =   6
      Top             =   90
      Width           =   5460
      _Version        =   65536
      _ExtentX        =   9631
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "     Elecsys 2010 INTERFACE"
      ForeColor       =   16777215
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4410
         Picture         =   "frmInterface.frx":2288
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   60
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   9225
      Left            =   240
      TabIndex        =   22
      Top             =   1110
      Width           =   7725
      _Version        =   196613
      _ExtentX        =   13626
      _ExtentY        =   16272
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   21
      Protect         =   0   'False
      RowHeaderDisplay=   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmInterface.frx":2B52
   End
   Begin VB.Frame Frame1 
      Height          =   9600
      Left            =   150
      TabIndex        =   20
      Top             =   870
      Width           =   15030
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1995
      Left            =   4290
      TabIndex        =   5
      Top             =   2730
      Width           =   2805
      _Version        =   196613
      _ExtentX        =   4948
      _ExtentY        =   3519
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmInterface.frx":4A1A
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1500
      Width           =   2055
   End
   Begin VB.TextBox txtAll 
      Height          =   375
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2610
      Width           =   2055
   End
   Begin VB.TextBox txtDate 
      Height          =   405
      Left            =   5190
      TabIndex        =   4
      Top             =   1950
      Width           =   2325
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   2175
      Left            =   8880
      TabIndex        =   25
      Top             =   3030
      Width           =   2895
      _Version        =   196613
      _ExtentX        =   5106
      _ExtentY        =   3836
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
      SpreadDesigner  =   "frmInterface.frx":8F21
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검 사 자"
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
      Left            =   5730
      TabIndex        =   14
      Top             =   525
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사일자"
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
      Left            =   5715
      TabIndex        =   12
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "오늘 검사수"
      Height          =   195
      Left            =   6135
      TabIndex        =   10
      Top             =   225
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblToday 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "100,000"
      Height          =   195
      Left            =   7395
      TabIndex        =   9
      Top             =   225
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "현재 검사수"
      Height          =   195
      Left            =   6570
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   195
      Left            =   8025
      TabIndex        =   7
      Top             =   555
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pp"
      Visible         =   0   'False
      Begin VB.Menu subUp 
         Caption         =   "검체번호 변경"
      End
      Begin VB.Menu subDel 
         Caption         =   "검체번호 삭제"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colRack = 3
Const colPos = 4
Const colPID = 5
Const colPName = 6
Const colJumin = 7
Const colPSex = 8
Const colPAge = 9
Const colOCnt = 10
Const colState = 11
Const colExamDate = 12
Const colReceNo = 13        '2004/07/28 이상은 - 접수번호 추가

Const colSampleType = 14         'Sample / QC 구분

Const colEquipExam = 3
Const colExamCode = 4
Const colExamName = 5
Const colResult = 6
Const colRCheck = 7
Const colPCheck = 8
Const colDCheck = 9
Const colUnit = 10
Const colRef = 11
Const colPanic = 12
Const colResult1 = 13

Dim ConfirmData As String

Dim gBarCode As String
Dim sBarCode As String
Dim sSeqNo As String
Dim sDiskNo As String
Dim sPosNo As String
Dim sSampleType As String
Dim sOrder As String
Dim llRow As Long
Dim sResDateTime As String

Sub Var_Clear()
    gOrderMessage = ""
    
    gBarCode = ""
    sBarCode = ""
    sSeqNo = ""
    sDiskNo = ""
    sPosNo = ""
    sSampleType = ""
    
    llRow = -1
End Sub

Private Sub btnClear_Click()
Dim iRow As Integer
    
    gWorkFlag = -1
    
    txtMsg.Text = ""
    
'    ClearSpread vasID, 1, 1
'    vasID.MaxRows = 0

    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
        
        If vasID.Value = 1 Then
            vasDeleteRow vasID, iRow
            
            iRow = iRow - 1
        End If
    Next iRow
    
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
End Sub

Private Sub btnClose_Click()
    Unload Me
    
    End
End Sub

Private Sub btnConfig_Click()
    frmConfig.SSPanel_machine.Caption = "Elecsys 2010"
    frmConfig.Show 1
End Sub

Private Sub btnSetup_Click()
    frmEquipExam.SSPanel1.Caption = "  Elecsys 2010 장비 코드 설정"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub btnTrans_Click()
'선택전송
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

    If txtUID.Text = "" Then
        MsgBox "사용자 확인을 해 주십시오"
        txtUID.SetFocus
        Exit Sub
    End If
    
    If (vasID.DataRowCnt < 1) Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
    
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        
        If vasID.Value = 1 Then
            liRet = -1
            If Left(Trim(GetText(vasID, vasIDRow, colBarCode)), 1) = "Q" Then
                liRet = Insert_Data_QC(vasIDRow)
            Else
                If Trim(GetText(vasID, vasIDRow, colPID)) <> "" Then
                    liRet = Insert_Data(vasIDRow)
                End If
            End If
            
            If liRet = 1 Then
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", vasIDRow, colState
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", vasIDRow, colState
            End If
            
            vasID.Row = vasIDRow
            vasID.Col = 1
            vasID.Value = 0
        Else
        
        End If
    Next vasIDRow
        
    db_Commit gServer
End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmd_Trans_Click()
'선택전송
Dim vasIDRow As Integer
Dim vasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

    If txtUID.Text = "" Then
        MsgBox "사용자 확인을 해 주십시오"
        txtUID.SetFocus
        Exit Sub
    End If
    
    If (vasID.DataRowCnt < 1) Or (vasRes.DataRowCnt < 1) Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
    
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        
        '2005/05/07 이상은 - 체크된 열 저장되도록 함
        'If vasID.Value <> 1 Then '체크된 열은 저장이 안됨
        If vasID.Value = 1 Then
            liRet = -1
'            If Trim(GetText(vasID, vasIDRow, colSeqNo)) = "QC" Then
'                liRet = Insert_Data_QC(vasIDRow)
'            Else
                If Trim(GetText(vasID, vasIDRow, colPID)) <> "" Then
                    liRet = Insert_Data(vasIDRow)
                End If
'            End If
            
            If liRet = 1 Then
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", vasIDRow, colState
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", vasIDRow, colState
            End If
            
            vasID.Row = vasIDRow
            vasID.Col = 1
            vasID.Value = 0
        Else
        
        End If
    Next vasIDRow
        
    db_Commit gServer
    
End Sub

Function Insert_Data_QC(argSpcRow As Integer) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow As Integer
    Dim i As Integer
    
    Dim sExamCode As String
    Dim sResult As String
    
    Insert_Data_QC = -1
    
    ClearSpread vasResTemp
    
'2004/05/28 이상은
'MDB에서는 convert함수 안 됨
'    SQL = " Select equipcode, examcode, result, refflag, mid(resdate,1,10), mid(resdate,11) " & vbCrLf & _
'          " From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(Trim(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' "
          
    SQL = " Select equipcode, examcode, result, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where examdate = '" & Format(Trim(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' "
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    vasResTemp.MaxRows = vasResTemp.DataRowCnt + 1
    
    If vasResTemp.DataRowCnt < 1 Then
        Exit Function
    End If

    For i = 1 To vasResTemp.DataRowCnt
        sExamCode = ""
        sResult = ""
        
        sExamCode = Trim(GetText(vasResTemp, i, 2))
        sResult = Trim(GetText(vasResTemp, i, 3))
        
        If sExamCode <> "" And sResult <> "" Then
            'ExamRes에 저장
            SQL = "Update ExamRes " & vbCrLf & _
                  "Set " & vbCrLf & _
                  "    Result = '" & Trim(GetText(vasResTemp, i, 3)) & "', " & vbCrLf & _
                  "    Decision = '" & Trim(GetText(vasResTemp, i, 4)) & "', " & vbCrLf & _
                  "    PanicFlag = '', " & vbCrLf & _
                  "    DeltaFlag = '', " & vbCrLf & _
                  "    ExamDate = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss'), " & vbCrLf & _
                  "    ExamUID = '" & txtUID & "', " & vbCrLf & _
                  "    EquipCode = '" & gEquip & "', " & vbCrLf & _
                  "    ExamState = 'D', " & vbCrLf & _
                  "    Input_UID = '" & txtUID & "', " & vbCrLf & _
                  "    Input_DateTime =TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') "
        
            SQL = SQL & CR & _
              " Where HID = '117' " & vbCrLf & _
              "  and PID = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' " & vbCrLf & _
              "  and ReceNo = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "' " & vbCrLf & _
              "  and SpecimenID = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' " & vbCrLf & _
              "  and ExamCode = '" & sExamCode & "'  "
    
            res = SendQuery(gServer, SQL)
            If res = -1 Then
                db_RollBack gServer
                Exit Function
            End If
            
            'QCRes에 저장
            SQL = "UPDATE QCRES SET " & vbCrLf & _
                  " RESULT = '" & sResult & "', " & vbCrLf & _
                  " EXAMSTATE = 'B', " & vbCrLf & _
                  " EDIT_ID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                  " EXAM_DT = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
                  "WHERE SPC_NO = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' " & vbCrLf & _
                  "  AND RECENO = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "' " & vbCrLf & _
                  "  AND EXAMCODE = '" & sExamCode & "' "
            res = SendQuery(gServer, SQL)
            If res = -1 Then
                db_RollBack gServer
                Exit Function
            End If
        End If
    Next i
    
    'ExamReq에업데이트
    SQL = "UPDATE ExamReq SET " & vbCrLf & _
          " ExamState = 'B' " & vbCrLf & _
          "WHERE HID = '117' " & vbCrLf & _
          " And PID = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' " & vbCrLf & _
          " And SpecimenID = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gServer, SQL)
    If res = -1 Then
        db_RollBack gServer
        Exit Function
    End If
    
    'QCReq에 업데이트
    SQL = "UPDATE QCREQ SET " & vbCrLf & _
          " QCFLAG = 'B' " & vbCrLf & _
          "WHERE SPC_NO = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' " & vbCrLf & _
          "  AND RECENO = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "' "
    res = SendQuery(gServer, SQL)
    If res = -1 Then
        db_RollBack gServer
        Exit Function
    End If
            
    Insert_Data_QC = 1
End Function

Function Insert_Data(argSpcRow As Integer) As Integer
'서버의 데이타 베이스에 저장

    Dim iRow As Integer
    Dim i As Integer
    Dim sCnt As String
    
    Insert_Data = -1
       
    '2005/06/16 이상은
    sCnt = "0"
    SQL = " Select count(ExamState) From ExamRes " & vbCrLf & _
          " Where HID = '117' " & vbCrLf & _
          " And PID = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' " & vbCrLf & _
          " And ReceNo = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "' " & vbCrLf & _
          " And SpecimenID = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' " & vbCrLf & _
          " And ExamCode in (" & gAllExam & ") " & vbCrLf & _
          " And ExamState = 'D' "
    res = db_select_Var(gServer, SQL, sCnt)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    If sCnt = "" Then
        sCnt = "0"
    ElseIf sCnt > "0" Then
        If MsgBox("해당 환자의 결과가 완료되었습니다." & CR & _
                "그래도 전송하시겠습니까?", vbExclamation + vbYesNo, "확인") = vbNo Then
            Exit Function
        End If
    End If
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select equipcode, examcode, result, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where examdate = '" & Format(Trim(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' "
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    vasResTemp.MaxRows = vasResTemp.DataRowCnt + 1
    
    '서버로 결과값 저장하기
    For i = 1 To vasResTemp.DataRowCnt
'        If Len(Trim(GetText(vasID, argSpcRow, colBarCode))) = 12 Then
            SQL = "Update ExamRes " & vbCrLf & _
                  "Set " & vbCrLf & _
                  "    Result = '" & Trim(GetText(vasResTemp, i, 3)) & "', " & vbCrLf & _
                  "    Decision = '" & Trim(GetText(vasResTemp, i, 4)) & "', " & vbCrLf & _
                  "    PanicFlag = '" & Trim(GetText(vasResTemp, i, 5)) & "', " & vbCrLf & _
                  "    DeltaFlag = '" & Trim(GetText(vasResTemp, i, 6)) & "', " & vbCrLf & _
                  "    ExamDate = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss'), " & vbCrLf & _
                  "    ExamUID = '" & txtUID & "', " & vbCrLf & _
                  "    EquipCode = '" & gEquip & "', " & vbCrLf & _
                  "    ExamState = 'D', " & vbCrLf & _
                  "    Input_UID = '" & txtUID & "', " & vbCrLf & _
                  "    Input_DateTime =TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') "
        
            SQL = SQL & CR & _
              " Where HID = '117' " & vbCrLf & _
              "  and PID = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' " & vbCrLf & _
              "  and ReceNo = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "' " & vbCrLf & _
              "  and SpecimenID = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' " & vbCrLf & _
              "  and ExamCode = '" & Trim(GetText(vasResTemp, i, 2)) & "'  "
    
            res = SendQuery(gServer, SQL)
'        ElseIf Len(Trim(GetText(vasID, argSpcRow, colBarCode))) = 13 Then
'            SQL = "Update ExamRes " & vbCrLf & _
'                  "Set " & vbCrLf & _
'                  "    Result = '" & Trim(GetText(vasResTemp, i, 3)) & "', " & vbCrLf & _
'                  "    Decision = '" & Trim(GetText(vasResTemp, i, 4)) & "', " & vbCrLf & _
'                  "    PanicFlag = '" & Trim(GetText(vasResTemp, i, 5)) & "', " & vbCrLf & _
'                  "    DeltaFlag = '" & Trim(GetText(vasResTemp, i, 6)) & "', " & vbCrLf & _
'                  "    ExamDate = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss'), " & vbCrLf & _
'                  "    ExamUID = '" & txtUID & "', " & vbCrLf & _
'                  "    EquipCode = '" & gEquip & "', " & vbCrLf & _
'                  "    ExamState = 'D', " & vbCrLf & _
'                  "    Input_UID = '" & txtUID & "', " & vbCrLf & _
'                  "    Input_DateTime =TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') "
'
'            SQL = SQL & CR & _
'              " Where HID = '117' " & vbCrLf & _
'              "  and PID = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' " & vbCrLf & _
'              "  and ReceNo = '" & Trim(GetText(vasID, argSpcRow, colBarCode)) & "' " & vbCrLf & _
'              "  and ExamCode = '" & Trim(GetText(vasResTemp, i, 2)) & "'  "
'            res = SendQuery(gServer, SQL)
'
'        End If
        If res = -1 Then
            db_RollBack gServer
            Exit Function
        End If
    Next i
    
    Insert_Data = 1
End Function

Private Sub cmdClear_Click()
Dim iRow As Integer
    
    gWorkFlag = -1
    
    txtMsg.Text = ""
    
'    ClearSpread vasID, 1, 1
'    vasID.MaxRows = 0

    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
        
        If vasID.Value = 1 Then
            vasDeleteRow vasID, iRow
            
            iRow = iRow - 1
        End If
    Next iRow
    
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConfig_Click()
    frmConfig.SSPanel_machine.Caption = "Elecsys 2010"
    frmConfig.Show 1
End Sub

Private Sub cmdSetup_Click()
    frmEquipExam.SSPanel1.Caption = "  Elecsys 2010 장비 코드 설정"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub cmdWorkList_Click()
    frmWorkList.Left = 0
    frmWorkList.Top = 0
    gWorkFlag = -1
    frmWorkList.Show
End Sub

Private Sub Command1_Click()
    ELECSYS2010 txtBuff.Text
    
    txtBuff.Text = ""
End Sub

Private Sub Form_Activate()
    vasActiveCell vasID, 1, colBarCode
    vasID.SetFocus
End Sub

Private Sub Form_Load()
    Dim sDate As String
    '1. 화면 및 변수 초기화
    '2. 데이타베이스에 Connect 하기 - Local - Server
    '3. Ini 내용 불러오기    GetSetup
    '4. Comport Open
    
On Error GoTo errFind

    Me.Left = 0
    Me.Top = 0
    
    cmdClear_Click
        
    ClearSpread vasID, 1, 1
    vasID.MaxRows = 1
    
    GetSetup    'ini에서 DB정보 불러오기
        
    If Not Connect_Server Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
    gEquip = "00050"   '강제설정
    GetEquipInfo    'Server에서 장비정보 불러오기
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
    
    Me.txtUID = gExamUID
    
    raw_data = ""
    
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    txtToday = Format(CDate(GetDateFull), "yyyy/mm/dd")
    
    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", CDate(txtToday.Text), -30), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
    '===================================================================
    
    '검사코드 가져오기
    GetExamCode
        
    'MultiSelect Mode
    vasRes.OperationMode = 1
    
    'fontsize
    vasRes.FontSize = 9
    
    llRow = -1
    
    gHeader = "H|\^&|||ASTM-Host" & chrCR & chrETX
    gMsgEnd = "L|1" & chrCR & chrETX
    
errFind:
'2005/06/16 이상은
    If Err = 8002 Then      'Port
        MsgBox "통신 포트를 확인하세요!", vbExclamation, "알림"
        cmdConfig_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    DisConnect_Server
    DisConnect_Local
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    
    gAllExam = ""
    
    For i = 1 To 24
        gArr_Exam(i, 1) = ""
        gArr_Exam(i, 2) = ""
        gArr_Exam(i, 3) = ""
    Next i
    
    ClearSpread vasTemp
    
    SQL = "Select EquipCode, ExamCode, ExamName From EquipExam where Equip = '" & gEquip & "' " & vbCrLf & _
          " Order by EquipCode"
          
    db_select_Vas gServer, SQL, vasTemp
    
    For i = 1 To vasTemp.DataRowCnt
        If IsNumeric(Trim(GetText(vasTemp, i, 1))) = True Then
            gArr_Exam(i, 1) = Trim(GetText(vasTemp, i, 1))
            gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '검사코드
            gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사항목코드
            
            '2005/02/23 이상은
            If gAllExam = "" Then
                gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
            Else
                gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
            End If
        End If
    Next i
    
End Sub

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    Dim sSendData
    
    lsChar = MSComm1.Input
    
    'raw_data = raw_data & lsChar
    
    Select Case lsChar
    Case chrENQ
        Save_Raw_Data "[RX]" & lsChar
        txtBuff = ""
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
        
    Case chrSTX
        txtBuff.Text = ""
        
    Case chrETX
        txtBuff = txtBuff & lsChar
        
        Save_Raw_Data "[RX]" & txtBuff
        
        ELECSYS2010 txtBuff
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
    
    Case chrACK
        Save_Raw_Data "[RX]" & lsChar
        
        Select Case gTxMsgFlag
        Case ""  'ENQ 보낸 다음 처음 받은 ACK
            'Header 보내기
            sSendData = chrSTX & CStr(gCurTxCnt) & gHeader & ASTM_CSum(CStr(gCurTxCnt) & gHeader) & chrCR & chrLF
            
            gTxMsgFlag = "H"
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "H"
            'patient 보내기
            sSendData = chrSTX & CStr(gCurTxCnt) & gPatient & ASTM_CSum(CStr(gCurTxCnt) & gPatient) & chrCR & chrLF
            gTxMsgFlag = "P"
            
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "P"
            'Test Order 보내기
'                gTxOrder = CStr(gCurTxCnt) & "O|1|" & GetText(Form_in.vaSpread_Ref, gBarRow, 3)
            sSendData = chrSTX & CStr(gCurTxCnt) & gOrder & ASTM_CSum(CStr(gCurTxCnt) & gOrder) & chrCR & chrLF
            gTxMsgFlag = "O"
            
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "O"
            'Message Terminator 보내기
            sSendData = chrSTX & CStr(gCurTxCnt) & gMsgEnd & ASTM_CSum(CStr(gCurTxCnt) & gMsgEnd) & chrCR & chrLF
            gTxMsgFlag = "L"
            
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "L"
            sSendData = chrEOT
            gPreData = sSendData
            
            MSComm1.Output = sSendData
            gTxMsgFlag = ""
            
            'gTimerMode = -1
                        
        End Select
        Save_Raw_Data "[TX]" & sSendData
        
    Case chrEOT
        txtBuff = lsChar
        Save_Raw_Data "[RX]" & txtBuff
        txtBuff = ""
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
        
        If gOrderMessage = "Q" Then
            gTxMsgFlag = ""
            gCurTxCnt = 1
            
            gPatient = "P|1||" & gBarCode & chrCR & chrETX
            gOrder = "O|1|" & gBarCode & "|" & sSeqNo & "^" & sDiskNo & "^" & sPosNo & "|" & sOrder & "|R||||||N||||||||||||||O" & chrCR & chrETX
            MSComm1.Output = chrENQ
            Save_Raw_Data "[TX]" & chrENQ
        End If
        
    Case Else
        txtBuff.Text = txtBuff.Text & lsChar
    End Select
End Sub

'Sub Save_Raw_Data()
''argSQL의 내용을 파일로 저장
'    Dim FilNum
'    Dim sFileName As String
'
'    FilNum = FreeFile
'
'    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
'        MkDir (App.Path & "\Result")
'    End If
'
'    sFileName = Format(CDate(txtToday.Text), "yyyymmdd")
'
''    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
'    Open App.Path & "\Result\" & sFileName & ".txt" For Append As FilNum
'    Print #FilNum, raw_data
'    Close FilNum
'End Sub

Sub Save_Raw_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
        MkDir (App.Path & "\Result")
    End If
    
    sFileName = gEquip & "_" & Format(CDate(txtToday.Text), "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Result\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Sub ELECSYS2010(asData As String)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim ResultTbl(1 To 40) As String        'Array에 담기
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim sCnt As String
    
    Dim sDate As String
    Dim sRefFlag As String
    Dim sResEnd As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sReceNo As String
    Dim sReceDate  As String
    Dim sPID As String
    Dim sPName As String
    Dim sJumin As String
    Dim sPSex As String
    Dim sPage As String
    Dim sTestID As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sExamDate As String
    
    Dim sSampleType As String
    
    Dim sLevelNo As String
    
    Dim lsTemp1 As String
    Dim jRow As Integer
    
    If asData = "" Then
        Exit Sub
    End If
    
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then           'Header Record
        Var_Clear
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "Q" Then           'Query Record
        gOrderMessage = "Q"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sBarCode = Left(sTmp, i - 1)
        gBarCode = Trim(sBarCode)
        
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sSeqNo = Left(sTmp, i - 1)
        
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sDiskNo = Left(sTmp, i - 1)
        
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sPosNo = Left(sTmp, i - 1)
        
        llRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = gBarCode Then
                llRow = i
                Exit For
            End If
        Next i
        If llRow = -1 Then
            llRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < llRow Then
                vasID.MaxRows = llRow
            End If
        End If
        
        SetText vasID, gBarCode, llRow, colBarCode
        SetText vasID, sSeqNo, llRow, 0
        SetText vasID, sDiskNo, llRow, colRack
        SetText vasID, sPosNo, llRow, colPos
                            
        '샘플의 환자 정보 가져오기
        If Trim(GetText(vasID, llRow, colPID)) = "" Then
            Get_Sample_Info llRow
        End If
        
        'Order 만들기
        sOrder = Make_Order(gBarCode, llRow)
        
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "O") Then         'Test Order Record
        sBarCode = Trim(ResultTbl(3))
        
        sResDateTime = Trim(ResultTbl(7))
        
        sTmp = ResultTbl(4)
        i = InStr(1, sTmp, "^")
        sSeqNo = Left(sTmp, i - 1)
        
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sDiskNo = Left(sTmp, i - 1)

        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sPosNo = Left(sTmp, i - 1)
        
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        sSampleType = Left(sTmp, i - 1)
        sTmp = Mid(sTmp, i + 1)
        
        If sSampleType = "SAMPLE" Then
            sSampleType = "P"
        ElseIf sSampleType = "CONTROL" Then
            sSampleType = "Q"
        End If
        
        llRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = sBarCode Then
                llRow = i
                Exit For
            End If
        Next i
        If llRow = -1 Then
            llRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < llRow Then
                vasID.MaxRows = llRow
            End If
        End If
        
        SetText vasID, sBarCode, llRow, colBarCode
        SetText vasID, sSeqNo, llRow, 0
        SetText vasID, sDiskNo, llRow, colRack
        SetText vasID, sPosNo, llRow, colPos
                            
        If Trim(GetText(vasID, llRow, colPID)) = "" Then
            '샘플의 환자 정보 가져오기
            Get_Sample_Info llRow
        End If
        
        ClearSpread vasRes
        
    End If

    If (Mid(ResultTbl(1), 2, 1) = "R") Then
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            sTestID = Left(sTmp, i - 1)
        Else
            sTestID = ""
        End If
        sTmp = ResultTbl(4)
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            sResult = Mid(sTmp, i + 1)
        Else
            sResult = sTmp
        End If
        
        If Not IsNumeric(sResult) And Trim(sResult) <> "" Then
            Select Case sTestID
            Case "400", "900"    'HBs Ag
                sTmp = ""
                For i = 1 To Len(sResult)
                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
                        sTmp = sTmp & Mid(sResult, i, 1)
                    End If
                Next i
                If CCur(sTmp) > 1 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negatvie"
                End If
                
            Case "410"    'HBs Ab
                sTmp = ""
                For i = 1 To Len(sResult)
                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
                        sTmp = sTmp & Mid(sResult, i, 1)
                    End If
                Next i
                If CCur(sTmp) >= 10 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negatvie"
                End If
                
            Case "430"    'Anti-Hbe
                sTmp = ""
                For i = 1 To Len(sResult)
                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
                        sTmp = sTmp & Mid(sResult, i, 1)
                    End If
                Next i
                If CCur(sTmp) > 1 Then
                    sRefFlag = "Negative"
                Else
                    sRefFlag = "Positive"
                End If
                
                
            Case "440"    'HBe Ag
                sTmp = ""
                For i = 1 To Len(sResult)
                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
                        sTmp = sTmp & Mid(sResult, i, 1)
                    End If
                Next i
                If CCur(sTmp) >= 1 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negative"
                End If

'            Case 460    'HBc IgM
'                sTmp = ""
'                For i = 1 To Len(sResult)
'                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
'                        sTmp = sTmp & Mid(sResult, i, 1)
'                    End If
'                Next i
'                If CCur(sTmp) >= 1 Then
'                    sRefFlag = "Pos"
'                Else
'                    sRefFlag = "Neg"
'                End If

            Case "560"    'HIV
                sTmp = ""
                For i = 1 To Len(sResult)
                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
                        sTmp = sTmp & Mid(sResult, i, 1)
                    End If
                Next i
                If CCur(sTmp) >= 1 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negatvie"
                End If
            Case Else
                sTmp = ""
                For i = 1 To Len(sResult)
                    If Mid(sResult, i, 1) <> "<" And Mid(sResult, i, 1) <> ">" And Mid(sResult, i, 1) <> "=" Then
                        sTmp = sTmp & Mid(sResult, i, 1)
                    End If
                Next i
            End Select

        ElseIf IsNumeric(sResult) Then
            Select Case sTestID
            Case "400", "900"    'HBs Ag
                If CCur(sResult) > 1 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negative"
                End If
                
            Case "410"    'HBs Ab
                If CCur(sResult) >= 10 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negative"
                End If
            Case "430"    'Anti-Hbe
                If CCur(sResult) > 1 Then
                    sRefFlag = "Negative"
                Else
                    sRefFlag = "Positive"
                End If

             Case "440"   'HBe Ag
                If CCur(sResult) >= 1 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negative"
                End If

             Case "560"   'HIV
                If CCur(sResult) >= 1 Then
                    sRefFlag = "Positive"
                Else
                    sRefFlag = "Negative"
                End If
                
'             Case 460   'HBc IgM
'                If CCur(sResult) >= 1 Then
'                    sRefFlag = "Pos"
'                Else
'                    sRefFlag = "Neg"
'                End If
            End Select
        End If
        
        '최종결과
        sResEnd = ""
        
        If sRefFlag <> "" Then
            sResEnd = sRefFlag
        Else
            sResEnd = sResult
        End If
            
        'sResDateTime = ResultTbl(13)    'result time
        
        '검사코드만큼 Row의 갯수를 설정
        gReadBuf(0) = ""
        
        SQL = "Select count(ExamCode) From EquipExam" & vbCrLf & _
                  " Where Equip = '" & gEquip & "' "
        res = db_select_Col(gServer, SQL)
        vasRes.MaxRows = gReadBuf(0)

        ClearSpread vasTemp
        
        SQL = "Select ExamCode From EquipExam" & vbCrLf & _
              " Where Equip = '" & gEquip & "' " & vbCrLf & _
              "  And EquipCode = '" & sTestID & "'"
        res = db_select_Vas(gServer, SQL, vasTemp)
        
        lsTemp1 = ""
        For jRow = 1 To vasTemp.DataRowCnt
            If lsTemp1 = "" Then
                lsTemp1 = "'" & Trim(GetText(vasTemp, jRow, 1)) & "'"
            Else
                lsTemp1 = lsTemp1 & ", '" & Trim(GetText(vasTemp, jRow, 1)) & "'"
            End If
        Next jRow
            
            
        If Len(Trim(GetText(vasID, llRow, colBarCode))) = 13 Then       '메뉴얼
            gReadBuf(0) = ""
            SQL = " Select a.ExamCode, b.ExamAlias From ExamRes a , ExamMaster b " & CR & _
                  " Where a.HID = '117' " & CR & _
                  " And a.PID = '" & Trim(GetText(vasID, llRow, colPID)) & "' " & CR & _
                  " And a.ReceNo = '" & Trim(GetText(vasID, llRow, colBarCode)) & "' " & CR & _
                  " And a.ExamCode in ( " & lsTemp1 & " ) " & CR & _
                  " And a.HID = b.HID " & CR & _
                  " And a.ExamCode = b.ExamCode "
        Else
            gReadBuf(0) = ""
            SQL = " Select a.ExamCode, b.ExamAlias From ExamRes a , ExamMaster b " & CR & _
                  " Where a.HID = '117' " & CR & _
                  " And a.PID = '" & Trim(GetText(vasID, llRow, colPID)) & "' " & CR & _
                  " And a.ReceNo = '" & Trim(GetText(vasID, llRow, colReceNo)) & "' " & CR & _
                  " And a.SpecimenID = '" & Trim(GetText(vasID, llRow, colBarCode)) & "' " & CR & _
                  " And a.ExamCode in ( " & lsTemp1 & " ) " & CR & _
                  " And a.HID = b.HID " & CR & _
                  " And a.ExamCode = b.ExamCode "
        End If
        res = db_select_Col(gServer, SQL)
        
        If res < 1 Then
            SQL = " Select ExamCode, ExamName From EquipExam " & CR & _
                  " Where Equip = '" & gEquip & "' " & vbCrLf & _
                  " And EquipCode = '" & sTestID & "'"
            res = db_select_Col(gServer, SQL)
            
        End If
        
        If (res = 1) And (gReadBuf(0) <> "") Then
            If vasRes.DataRowCnt = 0 Then
                vasRes.MaxRows = 1
            Else
                vasRes.MaxRows = vasRes.DataRowCnt + 1
            End If

            j = vasRes.DataRowCnt + 1
            
            If sResult <> "" Then
'                If gRType = "C" Then
'                    SetText vasRes, gSpec, j, colBarCode
'                Else
                    SetText vasRes, sBarCode, j, colBarCode         '검체번호
'                End If
                
                SetText vasRes, Trim(sTestID), j, colEquipExam      '장비코드
                SetText vasRes, gReadBuf(0), j, colExamCode         '검사코드
                SetText vasRes, gReadBuf(1), j, colExamName         '검사명
                SetText vasRes, sResEnd, j, colResult               '검사결과
                SetText vasRes, sResEnd, j, colResult1              '검사결과
                sExamCode = gReadBuf(0)
                
                '판정
                If Left(Trim(GetText(vasID, llRow, colBarCode)), 1) = "Q" Then
                    QC_Result Trim(GetText(vasID, llRow, colBarCode)), sExamCode, sResult, j
                Else
                    Check_Result Trim(GetText(vasID, llRow, colBarCode)), Trim(GetText(vasID, llRow, colPID)), sExamCode, sResEnd, j, Trim(GetText(vasID, llRow, colPSex))
                End If
                
                Save_Local_One llRow, j, "A"
            End If
        Else
            SetText vasRes, "", j, colResult
        End If
            
        SetText vasID, "수신완료", llRow, colState
        SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
        '==============================================================================================
    End If

End Sub

Function Make_Order(argBarCode As String, asRow As Long) As String
    Dim i As Integer
    Dim sOrder As String
    Dim sEquipCode As String
    
    Make_Order = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    sOrder = ""
    
'    sOrder = "^^^400^0\^^^410^0\^^^430^0\^^^440^0"
'    MakeOrder = 1
'
'    Exit Function
    
    ClearSpread vasTemp
        
    SQL = " Select ExamCode From ExamRes" & vbCrLf & _
          " Where HID = '117' " & vbCrLf & _
          " And PID = '" & Trim(GetText(vasID, asRow, colPID)) & "' " & vbCrLf & _
          " And ReceNo = '" & Trim(GetText(vasID, asRow, colReceNo)) & "' " & vbCrLf & _
          " And SpecimenID = '" & Trim(argBarCode) & "'" & vbCrLf & _
          " And ExamCode in (" & gAllExam & ") " & vbCrLf & _
          " And NVL(ExamState, ' ') <> 'D' "
    res = db_select_Vas(gServer, SQL, vasTemp)

    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If

    If vasTemp.DataRowCnt > 0 Then
        For i = 1 To vasTemp.DataRowCnt
            If Trim(GetText(vasTemp, i, 1)) <> "" Then
                '검사코드로 장비코드 찾기
                sEquipCode = GetEquip_ExamCode(Trim(GetText(vasTemp, i, 1)))
                
                If sOrder = "" Then
                    sOrder = "^^^" & Trim(sEquipCode) & "^0"
                Else
                    sOrder = sOrder & "\^^^" & Trim(sEquipCode) & "^0"
                End If
    
            End If
        Next i
        
        Make_Order = 1
        SetText vasID, "Order", asRow, colState
    Else
        Make_Order = 0
        SetText vasID, "없음", asRow, colState
    End If
    
    Make_Order = sOrder
    
End Function

Function Get_QC_Info(ByVal asRow As Long) As Integer
    Dim sID, lsPID As String

    '샘플 환자 장보 가져오기
    sID = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
    '환자번호, 환자이름, 주민번호, 성별, 나이, 처방개수
    SQL = "select distinct QCBarcode, LotNo, QCInLevel, LabCode   " & vbCrLf & _
          "from QCInItem " & vbCrLf & _
          "where EquipCode = '" & gEquip & "' " & vbCrLf & _
          "  and QCBarcode = '" & sID & "' " & vbCrLf & _
          "  and AppDate <= '" & Trim(txtToday.Text) & "' "
    res = db_select_Col(gServer, SQL)
    If res = 1 Then
        SetText vasID, gReadBuf(1), asRow, colPID
        SetText vasID, gReadBuf(2), asRow, colPName
        SetText vasID, gReadBuf(3), asRow, colJumin
    End If
    
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    
End Function

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsBarCode As String
    Dim lsPID As String
    Dim lsReceNo As String
    
    '샘플 환자 정보 가져오기
    lsBarCode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
'    SQL = " Select /*+ Index_Desc(ExamRes,ExamRes_Index) */  PID, ReceNo " & vbCrLf & _
'          " From ExamRes " & vbCrLf & _
'          " Where HID = '117' " & vbCrLf & _
'          " And SpecimenID = '" & lsBarCode & "'" & vbCrLf & _
'          " Group by PID, ReceNo "
          
    SQL = " Select PID, ReceNo " & vbCrLf & _
          " From ExamRes " & vbCrLf & _
          " Where HID = '117' " & vbCrLf & _
          " And SpecimenID = '" & lsBarCode & "'" & vbCrLf & _
          " Group by PID, ReceNo "
    res = db_select_Col(gServer, SQL)
    
    If res = 1 Then
        lsPID = Trim(gReadBuf(0))
        SetText vasID, lsPID, asRow, colPID
        
        lsReceNo = Trim(gReadBuf(1))
        SetText vasID, lsReceNo, asRow, colReceNo
    Else
        lsPID = ""
        lsReceNo = ""
    End If
    
    If lsPID <> "" Then
        '챠트번호, 환자이름, 성별, 주민번호
        SQL = " Select pid, pname, jumin1 || jumin2 " & vbCrLf & _
              " From patient " & vbCrLf & _
              " Where pid = '" & lsPID & "'"
        res = db_select_Col(gServer, SQL)
    
        If res = 1 Then
            SetText vasID, gReadBuf(1), asRow, colPName
            SetText vasID, gReadBuf(2), asRow, colJumin
        
            CalAgeSex gReadBuf(2), txtToday.Text
            SetText vasID, gPatGen.Sex, asRow, colPSex
            SetText vasID, gPatGen.Age, asRow, colPAge
        End If
    End If
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""

    'Order 갯수 나타내기
    SQL = " Select count(ExamCode) " & vbCrLf & _
          " From ExamRes " & vbCrLf & _
          " Where HID = '117' " & vbCrLf & _
          " And PID = '" & Trim(lsPID) & "' " & vbCrLf & _
          " And ReceNo = '" & Trim(lsReceNo) & "' " & vbCrLf & _
          " And SpecimenID = '" & lsBarCode & "'" & vbCrLf & _
          " And ExamCode In (" & gAllExam & ") "
    res = db_select_Col(gServer, SQL)
    If res = 1 Then
        SetText vasID, gReadBuf(0), asRow, colOCnt
    End If

    
End Function

Function SetResult(asResult As String, aiItem As Integer) As String
''DB에서 불러오기
'    Dim iFloat As Integer
'
'    If Not IsNumeric(asResult) Then
'        Exit Function
'    End If
'
'    Select Case aiItem
'    Case 7, 16
'        iFloat = 2
'    Case 14
'        iFloat = 0
'    Case Else
'        iFloat = 1
'    End Select
'
'    If iFloat = 0 Then
'        SetResult = CStr(CCur(asResult))
'    Else
'        SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
'    End If
 
    Dim iFloat As Integer
   
    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    SQL = " Select PointSize From EquipExam " & vbCrLf & _
          " Where EquipCode = '" & Trim(gArr_ExamCode(aiItem, 1)) & "' " & vbCrLf & _
          " And Equip = '" & gEquip & "' "
          
    res = db_select_Col(gServer, SQL)
    
    iFloat = Trim(gReadBuf(0))

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
            SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
        Else
            SetResult = CStr(CCur(Left(asResult, 4 - iFloat)) & "." & Right(asResult, iFloat))
        End If
    End If
End Function



Private Sub sspMode_Click()
    If sspMode.Caption = "수정모드" Then
        sspMode.Caption = "전송모드"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "전송모드" Then
        sspMode.Caption = "수정모드"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 0
        
        vasActiveCell vasRes, 1, colResult
        vasRes.SetFocus
    End If

End Sub

Private Sub subDel_Click()
    Dim i As Long
    
    i = vasID.ActiveRow
    
    SQL = " Delete From pat_res " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And examdate = '" & Format(txtToday.Text, "yyyymmdd") & "' " & CR & _
          " And barcode = '" & Trim(GetText(vasID, i, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
    
    vasID.DeleteRows i, 1
    If i > vasID.DataRowCnt Then
        i = vasID.DataRowCnt
    End If
    vasID.MaxRows = vasID.DataRowCnt
    vasActiveCell vasID, i, colBarCode
    vasID.SetFocus
End Sub

Private Sub subUp_Click()
Dim sValue As String
Dim sTmp As String
Dim i As Integer
Dim j As Integer

    sTmp = ""
    
    vasID.Row = vasID.ActiveRow
    vasID.Col = vasID.ActiveCol
    
    sTmp = vasID.Text
    
    sValue = InputBox("변경할 검체번호를 입력하세요")
        
    If Trim(sValue) <> "" Then
        If MsgBox("" & sTmp & "를 " & sValue & "로 수정하시겠습니까?", vbYesNo, "확인") = vbYes Then
            SetText vasID, sValue, vasID.Row, vasID.Col
            
            If Trim(GetText(vasID, vasID.Row, colBarCode)) <> "" Then
                Get_Sample_Info vasID.Row
                            
                For i = 1 To vasRes.DataRowCnt
                    Save_Local_One vasID.Row, i, "A"
                Next
            End If
        End If
    End If

End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim lsTmpID As String
    
    Dim i As Integer
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarCode))

    ClearSpread vasRes
    vasRes.MaxRows = 0
    
    SQL = "select '', barcode, equipcode,  examcode, examname, result, refflag, panicflag, deltaflag, unit, refvalue, panicvalue, result " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND Barcode = '" & Trim(GetText(vasID, vasID.Row, colBarCode)) & "' " & vbCrLf & _
          "  order by equipcode"
          
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For i = 1 To vasRes.DataRowCnt
        '참조치
        Select Case Trim(GetText(vasRes, i, colRCheck))
        Case "H"
            vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(205, 55, 0)
        Case "L"
            vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(65, 105, 225)
        Case ""
             vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(255, 255, 255)
        End Select
        
        'Panic
        Select Case Trim(GetText(vasRes, i, 8))
        Case "H"
            vasRes.Row = i
            vasRes.Col = 8
            vasRes.ForeColor = RGB(205, 55, 0)
        Case "L"
            vasRes.Row = i
            vasRes.Col = 8
            vasRes.ForeColor = RGB(65, 105, 225)
        Case ""
             vasRes.Row = i
            vasRes.Col = 8
            vasRes.ForeColor = RGB(255, 255, 255)
        End Select
            
        'Delta
        Select Case Trim(GetText(vasRes, i, 9))
        Case "D"
            vasRes.Row = i
            vasRes.Col = 9
            vasRes.ForeColor = RGB(205, 55, 0)
        Case "L"
            vasRes.Row = i
            vasRes.Col = 9
            vasRes.ForeColor = RGB(65, 105, 225)
        Case ""
             vasRes.Row = i
            vasRes.Col = 9
            vasRes.ForeColor = RGB(255, 255, 255)
        End Select
    Next i
End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    If Trim(GetText(vasID, asRow1, colSampleType)) = "QC" Then
        sExamDate = Trim(GetText(vasID, asRow1, colExamDate))

        '2004/05/28 이상은
        'sExamDate = Left(sExamDate, 4) & "-" & Mid(sExamDate, 5, 2) & "-" & Mid(sExamDate, 7, 2) & " " & Mid(sExamDate, 9, 2) & ":" & Mid(sExamDate, 11, 2) & ":00"
    Else
        sExamDate = GetDateFull
    End If
    
    sCnt = ""
    SQL = "delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasRes, asRow2, colBarCode)) & "' "
    SaveQuery SQL
    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, receno, pid, " & _
          "pname, pjumin, page, psex, resdate, " & _
          "equipcode, examcode, examtype, result, sendflag, examname, " & _
          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(txtToday.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasID, asRow1, colBarCode)) & "', '', " & _
          "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colJumin)) & "', " & _
          "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
          "'" & sExamDate & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', '', " & _
          "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colPCheck)) & "', " & _
          "'" & Trim(GetText(vasRes, asRow2, colDCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colUnit)) & "', " & _
          "'" & Trim(GetText(vasRes, asRow2, colRef)) & "', '" & Trim(GetText(vasRes, asRow2, colPanic)) & "') "
    
    SaveQuery SQL
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim lsID As String
'    Dim lsTmpID As String
'
'    Dim i As Integer
'
'    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
'
'    If Row < 1 Or Row > vasID.DataRowCnt Then
'        Exit Sub
'    End If
'
'    lsID = Trim(GetText(vasID, Row, colBarCode))
'
'    ClearSpread vasRes
'    vasRes.MaxRows = 0
'
'    SQL = "select '', barcode, equipcode,  examcode, examname, result, refflag, panicflag, deltaflag, unit, refvalue, panicvalue, result " & vbCrLf & _
'          "FROM pat_res " & vbCrLf & _
'          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'          "  AND Barcode = '" & Trim(GetText(vasID, vasID.Row, colBarCode)) & "' " & vbCrLf & _
'          "  order by equipcode"
'
'    res = db_select_Vas(gLocal, SQL, vasRes)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    For i = 1 To vasRes.DataRowCnt
'        '참조치
'        Select Case Trim(GetText(vasRes, i, colRCheck))
'        Case "H"
'            vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'
'        'Panic
'        Select Case Trim(GetText(vasRes, i, 8))
'        Case "H"
'            vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'
'        'Delta
'        Select Case Trim(GetText(vasRes, i, 9))
'        Case "D"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'    Next i
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    
    Dim iRow As Integer
    Dim lRow As Long

    iRow = vasID.ActiveRow
    lRow = iRow
    
    If KeyCode = vbKeyReturn Then
        If Trim(GetText(vasID, iRow, colBarCode)) <> "" Then
            Get_Sample_Info lRow

            For i = 1 To vasRes.DataRowCnt
                Save_Local_One lRow, i, "A"
            Next i
        End If
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop
End Sub

Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Response, Help
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    If KeyCode = vbKeyReturn Then
        vasIDRow = vasID.ActiveRow
        If vasResCol = colResult And _
           Trim(GetText(vasRes, vasResRow, colResult)) <> Trim(GetText(vasRes, vasResRow, colResult1)) Then
            
            Response = MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!", Help, 100)
            If Response = vbYes Then
                '판정, 델타, 패닉 수정
                Check_Result Trim(GetText(vasID, vasIDRow, colBarCode)), _
                             Trim(GetText(vasID, vasIDRow, colPID)), _
                             Trim(GetText(vasRes, vasResRow, colExamCode)), _
                             Trim(GetText(vasRes, vasResRow, colResult)), _
                             vasResRow, Trim(GetText(vasID, vasIDRow, colPSex))

                SQL = " Update pat_res " & vbCrLf & _
                      " Set result = '" & Trim(GetText(vasRes, vasResRow, colResult)) & "', " & vbCrLf & _
                      " refFlag = '" & Trim(GetText(vasRes, vasResRow, colRCheck)) & "', " & vbCrLf & _
                      " panicFlag = '" & Trim(GetText(vasRes, vasResRow, colPCheck)) & "', " & vbCrLf & _
                      " deltaFlag = '" & Trim(GetText(vasRes, vasResRow, colDCheck)) & "' " & vbCrLf & _
                      " WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipExam)) & "'" & vbCrLf & _
                      "  AND barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                
            End If
        End If
        
    End If
End Sub

Public Function QC_Result(argBarCode As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '결과종류
    Dim sLow        As String       '참조치
    Dim sHigh       As String
    Dim RefRet      As String
    
    Dim sPart       As String
    Dim sEquip      As String
    Dim sLevel      As String
    Dim sLotNo      As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    QC_Result = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        QC_Result = -1
        Exit Function
    End If
    sPart = Trim(GetText(vasID, argRow, colJumin))
    sEquip = gEquip
    sLevel = Trim(GetText(vasID, argRow, colPName))
    sLotNo = Trim(GetText(vasID, argRow, colPID))
    
    SQL = "Select Max(q.AppDate), e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh   " & vbCrLf & _
          "From QCInItem q, ExamMaster e " & vbCrLf & _
          "Where q.LabCode = '" & sPart & "' " & vbCrLf & _
          "  and q.EquipCode = '" & sEquip & "' " & vbCrLf & _
          "  and q.QCInLevel = '" & sLevel & "' " & vbCrLf & _
          "  and q.LotNo = '" & sLotNo & "' " & vbCrLf & _
          "  and q.QCBarcode = '" & argBarCode & "' " & vbCrLf & _
          "  and q.ExamCode = '" & argExamCode & "' " & vbCrLf & _
          "  and q.AppDate >= '1900-01-01' " & vbCrLf & _
          "  and e.AppDate = (select Max(c.AppDate) from ExamMaster c Where c.AppDate >= '1900-01-01' and c.ExamCode = q.ExamCode)" & vbCrLf & _
          "  and e.ExamCode = q.ExamCode " & vbCrLf & _
          "Group by e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh"
    res = db_select_Col(gServer, SQL)
    sResClassCode = Trim(gReadBuf(1))
    
    If sResClassCode = "1" Then '숫자
'참조치 체크
        sLow = ""
        sHigh = ""
        
        '숫자인지 아닌지 확인
        If IsNumeric(sDiffRet) = False Then
           MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
           QC_Result = -1
           Exit Function
        End If
        
        If IsNumeric(gReadBuf(2)) Then
            If CInt(gReadBuf(2)) > 0 Then
                sTmpStr = "#0."
                For i = 1 To CInt(gReadBuf(2))
                    sTmpStr = sTmpStr & "0"
                Next i
            Else
                sTmpStr = "#0"
            End If
            sDiffRet = Format(sDiffRet, sTmpStr)
            SetText vasRes, sDiffRet, argRow, colResult
            SetText vasRes, sDiffRet, argRow, colResult1
        End If
        
        sLow = Trim(gReadBuf(3))
        sHigh = Trim(gReadBuf(4))
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   '이상
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   '이하
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


        
    ElseIf sResClassCode = "2" Then '문자
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 이상은 수정
'        '검사 항목 결과 참조 코드 체크에서 1 이상일 경우만 판정되게
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002년 3월 12일 +-에서 +/-로 수정
'        '2002년 5월 13일 NON-REACTIVE 판정 안돼서 추가
'        '2003년 2월 4일 이상은 수정 - 0-1로 참조치는 1이나 판정됨
'        '=================================================================================
'        '2002년 5월 13일 1 : 40 미만 판정 안됨
'        '2002년 6월 11일 (결과참조가 1:로 시작하면 판정체크 안하게 수정)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "양성", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "양성", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
    End If
    
    SetText vasRes, RefRet, argRow, colRCheck
    
    If RefRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    QC_Result = 1

End Function

Public Function Check_Result(argBarCode As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '결과종류
    Dim sLow        As String       '참조치
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " Select ResClassCode, Res_M_Low, Res_M_High, Res_F_Low, Res_F_High, " & CR & _
          "        PanicValueGubun, Panic_M_Low, Panic_M_High, Panic_F_Low, Panic_F_High, " & CR & _
          "        DeltaValueGubun, DeltaLow, DeltaHigh, Point " & CR & _
          "From ExamMaster " & CR & _
          " Where HID = '117' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_select_Col(gServer, SQL)
    
    sResClassCode = Trim(gReadBuf(0))
    
    If sResClassCode = "1" Then '숫자
'참조치 체크
        sLow = ""
        sHigh = ""
        
        '숫자인지 아닌지 확인
        If IsNumeric(sDiffRet) = False Then
           'MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
           Check_Result = -1
           Exit Function
        End If
        
        If IsNumeric(gReadBuf(13)) Then
            If CInt(gReadBuf(13)) > 0 Then
                sTmpStr = "#0."
                For i = 1 To CInt(gReadBuf(13))
                    sTmpStr = sTmpStr & "0"
                Next i
            Else
                sTmpStr = "#0"
            End If
            sDiffRet = Format(sDiffRet, sTmpStr)
            SetText vasRes, sDiffRet, argRow, colResult
            SetText vasRes, sDiffRet, argRow, colResult1
        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(1))
            sHigh = Trim(gReadBuf(2))
        Case "F"
            sLow = Trim(gReadBuf(3))
            sHigh = Trim(gReadBuf(4))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   '이상
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   '이하
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


'Panic 체크
        sPanicLow = ""
        sPanicHigh = ""
        
        sPanicGubun = Trim(gReadBuf(5))
        
        Select Case asSex
        Case "M", ""
            sPanicLow = Trim(gReadBuf(6))
            sPanicHigh = Trim(gReadBuf(7))
        Case "F"
            sPanicLow = Trim(gReadBuf(8))
            sPanicHigh = Trim(gReadBuf(9))
        End Select
        
        If sPanicGubun = "0" Then '상한/하한
            If sPanicLow = "" Or sPanicHigh = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) > CCur(sDiffRet) Then
                    PanicRet = "L"
                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
                    PanicRet = "H"
                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
                    PanicRet = ""
                End If
            End If
        ElseIf sPanicGubun = "1" Then 'percent
            If sPanicLow = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "L"
                    Else
                        PanicRet = ""
                    End If
                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "H"
                    Else
                        PanicRet = ""
                    End If
                Else
                    PanicRet = ""
                End If
            End If
        End If
        

'Delta 체크
        sDeltaLow = ""
        sDeltaHigh = ""
                
        sTmpRece1 = ""
        sTmpRet1 = ""
        sTmpRece2 = ""
        sTmpRet2 = ""
        PreResult = ""
        
        sMax_ReceNo = ""
'        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
        sReceNo = argBarCode
        
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where HID = '117' " & CR & _
'              " And PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBarCode & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result"
              
'2004/12/30 이상은 - 정렬부분 추가
        If Len(Trim(argBarCode)) = 13 Then
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where HID = '117' " & CR & _
'              " And PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBarCode & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result" & CR & _
'              " Order by 2 desc "
        res = 0
        gReadBuf(0) = ""
        ElseIf Len(Trim(argBarCode)) = 12 Then
        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
              " Where HID = '117' " & CR & _
              " And PID = '" & Trim(argPID) & "' " & CR & _
              " And SpecimenID < '" & argBarCode & "' " & CR & _
              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
              " Group By Result" & CR & _
              " Order by 2 desc "
        res = db_select_Col(gServer, SQL)
        End If
              
        If res > 0 And gReadBuf(0) <> "" Then
            PreResult = gReadBuf(0)
        Else
            PreResult = ""
        End If
      
        If PreResult <> "" And IsNumeric(PreResult) = True Then
          'PreResult = Trim(gReadBuf(0))
          sDeltaGubun = Trim(gReadBuf(10))
          
          sDeltaLow = Trim(gReadBuf(11))
          sDeltaHigh = Trim(gReadBuf(12))
          
            '이전결과에서 현재결과 뺀값이 sDiffRet임 (2002년 3월 15일 수정)
'            sDiffRet = PreResult - sDiffRet
            sDiffRet1 = sDiffRet - PreResult
            If sDeltaGubun = "0" Then '상한/하한
                If sDeltaLow = "" Or sDeltaHigh = "" Then
                    DeltaRet = ""
                Else
                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
                        DeltaRet = "L"
                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
                        DeltaRet = "H"
                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
                        DeltaRet = ""
                    End If
                End If
              
            ElseIf sDeltaGubun = "1" Then 'percent
               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
                  DeltaRet = ""
               Else
                   If sDeltaLow = "" Then
                        DeltaRet = ""
                    Else
                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
                            DeltaRet = "D"
                        Else
                            DeltaRet = ""
                        End If
                    End If
               End If
            End If
        End If
        
    ElseIf sResClassCode = "2" Then '문자
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 이상은 수정
'        '검사 항목 결과 참조 코드 체크에서 1 이상일 경우만 판정되게
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002년 3월 12일 +-에서 +/-로 수정
'        '2002년 5월 13일 NON-REACTIVE 판정 안돼서 추가
'        '2003년 2월 4일 이상은 수정 - 0-1로 참조치는 1이나 판정됨
'        '=================================================================================
'        '2002년 5월 13일 1 : 40 미만 판정 안됨
'        '2002년 6월 11일 (결과참조가 1:로 시작하면 판정체크 안하게 수정)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "양성", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "양성", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "양성"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End Select
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
    End If
    
    SetText vasRes, RefRet, argRow, colRCheck
    SetText vasRes, PanicRet, argRow, colPCheck
    SetText vasRes, DeltaRet, argRow, colDCheck
    

    '2002년 2월 15일 수정 (판정시 H, L 일때 글자 색깔 변화)
    '2002년 3월 14일 수정 (판정시 L일때는 파란색 그 외는 빨간색)
    If RefRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If PanicRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If DeltaRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    ElseIf DeltaRet = "D" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    Check_Result = 1

End Function

