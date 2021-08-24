VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS401 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "주    소 2"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frmBBS401.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6360
   StartUpPosition =   1  '소유자 가운데
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   225
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "우편 번호"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   225
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1980
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "주    소 1"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   2
      Left            =   225
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "주    소 2"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   240
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "전화 번호"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3060
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "직       업"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   1035
      Index           =   5
      Left            =   240
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1826
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "혈  액  형"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lbldt 
      Height          =   315
      Left            =   240
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   240
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "성   명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   600
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "주민 번호"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   6
      Left            =   240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   960
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "생년 월일"
      Appearance      =   0
   End
   Begin VB.ComboBox cboDonor 
      Height          =   300
      ItemData        =   "frmBBS401.frx":076A
      Left            =   2850
      List            =   "frmBBS401.frx":076C
      Style           =   2  '드롭다운 목록
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton cmdZipCd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2475
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1620
      Width           =   360
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1395
      TabIndex        =   7
      Text            =   "txtTelNo"
      Top             =   2700
      Width           =   2010
   End
   Begin VB.TextBox txtAddr2 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1395
      TabIndex        =   6
      Text            =   "txtAddr2"
      Top             =   2340
      Width           =   4890
   End
   Begin VB.TextBox txtAddr1 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1395
      TabIndex        =   4
      Text            =   "txtAddr1"
      Top             =   1980
      Width           =   4890
   End
   Begin VB.TextBox txtZipCd 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1395
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "txtZipCd"
      ToolTipText     =   "검색 버튼을 이용하여 주소를 입력하시오."
      Top             =   1620
      Width           =   1050
   End
   Begin VB.TextBox txtSSN 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1410
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "123456-1234567"
      Top             =   600
      Width           =   1410
   End
   Begin VB.TextBox txtDonorNm 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      Text            =   "txtDonorNm"
      ToolTipText     =   "이름 사이를 띄우지 마시오."
      Top             =   240
      Width           =   1410
   End
   Begin VB.TextBox txtTotVol 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   4665
      TabIndex        =   16
      Text            =   "txtTotVol"
      Top             =   6015
      Width           =   1050
   End
   Begin VB.TextBox txtCnt 
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1410
      TabIndex        =   15
      Text            =   "txtCnt"
      Top             =   6015
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   1410
      TabIndex        =   21
      Top             =   4710
      Width           =   3795
      Begin VB.OptionButton optABO 
         BackColor       =   &H00DBE6E6&
         Caption         =   "A"
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   195
         Width           =   855
      End
      Begin VB.OptionButton optABO 
         BackColor       =   &H00DBE6E6&
         Caption         =   "B"
         Height          =   315
         Index           =   1
         Left            =   1020
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   195
         Width           =   855
      End
      Begin VB.OptionButton optABO 
         BackColor       =   &H00DBE6E6&
         Caption         =   "O"
         Height          =   315
         Index           =   2
         Left            =   1920
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   195
         Width           =   855
      End
      Begin VB.OptionButton optABO 
         BackColor       =   &H00DBE6E6&
         Caption         =   "AB"
         Height          =   315
         Index           =   3
         Left            =   2820
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   4005
      Style           =   1  '그래픽
      TabIndex        =   19
      Tag             =   "128"
      Top             =   6600
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   2542
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "124"
      Top             =   6600
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   1080
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "0"
      Top             =   6600
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      Height          =   330
      Left            =   1410
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   66060291
      CurrentDate     =   36868
   End
   Begin MSComctlLib.ListView lvwJob 
      Height          =   1455
      Left            =   1395
      TabIndex        =   8
      ToolTipText     =   "하나의 직업만 선택가능합니다. 없는 경우 기타를 선택하시오."
      Top             =   3060
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2566
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MedControls1.LisLabel lblDonorCd 
      Height          =   315
      Left            =   2850
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "lblDonorCd"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblSex 
      Height          =   315
      Left            =   4020
      TabIndex        =   26
      Top             =   945
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "lblSex"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   2865
      TabIndex        =   30
      Top             =   945
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "성별/나이"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   1410
      TabIndex        =   22
      Top             =   5235
      Width           =   2055
      Begin VB.OptionButton optRh 
         BackColor       =   &H00DBE6E6&
         Caption         =   "RH+"
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   195
         Width           =   855
      End
      Begin VB.OptionButton optRh 
         BackColor       =   &H00DBE6E6&
         Caption         =   "RH-"
         Height          =   315
         Index           =   1
         Left            =   1020
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   195
         Width           =   855
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   7
      Left            =   240
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "헌혈 횟수"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   8
      Left            =   3495
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "총 헌혈량"
      Appearance      =   0
   End
   Begin VB.Line Line6 
      X1              =   180
      X2              =   6060
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   6060
      Y1              =   6495
      Y2              =   6495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6120
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   6120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6160
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "CC"
      Height          =   180
      Left            =   5760
      TabIndex        =   24
      Top             =   6150
      Width           =   270
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "회"
      Height          =   180
      Left            =   2490
      TabIndex        =   23
      Top             =   6120
      Width           =   180
   End
End
Attribute VB_Name = "frmBBS401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'헌혈자 등록
'Coding By Legends

Private objMySQL As clsBBSSQLStatement
Private blnModify As Boolean
Private Sub cboDonor_Click()
    Dim strDonorNm As String
    Dim strSSN As String
    
    strSSN = Mid(cboDonor.Text, 1, 14)  '주민번호
    strSSN = Mid(strSSN, 1, 6) & Mid(strSSN, 8, 14)
    strDonorNm = Mid(cboDonor.Text, 18) '헌혈자 이름
    
    
    '헌혈자의 정보를 보여준다.
    Call ShowDonorValue(strDonorNm, strSSN)
    
    cmdSave.tag = "1"
    cmdSave.Caption = "수정"
End Sub

Private Sub cmdClear_Click()
    Call FormInitialize
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Function Save_Check() As Boolean
    '정보 누락 체크
    If Trim(txtDonorNm.Text) = "" Then
        MsgBox "성명을 입력하세요.", vbInformation, "정보확인"
        txtDonorNm.SetFocus
        Exit Function
    End If
    
    If Trim(txtSSN.Text) = "" Then
        MsgBox "주민 번호를 입력하세요.", vbInformation, "정보확인"
        txtSSN.SetFocus
        Exit Function
    End If
        
'    If Trim(txtZipCd.Text) = "" Then
'        MsgBox "우편 번호를 입력하세요.", vbInformation, "정보확인"
'        txtZipCd.SetFocus
'        Exit Function
'    End If
    
    If Trim(txtAddr1.Text) = "" Then
        MsgBox "주소를 입력하세요.", vbInformation, "정보확인"
        txtAddr1.SetFocus
        Exit Function
    End If
    
    If Trim(txtAddr2.Text) = "" Then
        MsgBox "상세 주소를 입력하세요.", vbInformation, "정보확인"
        txtAddr2.SetFocus
        Exit Function
    End If
    
'    If Trim(txtTelNo.Text) = "" Then
'        MsgBox "전화 번호를 입력하세요.", vbInformation, "정보확인"
'        txtTelNo.SetFocus
'        Exit Function
'    End If

'    If Val(lvwJob.Tag) < 0 Then
'        MsgBox "직업을 선택하세요.", vbInformation, "정보확인"
'        lvwJob.SetFocus
'        Exit Function
'    End If
    
    If optABO(0).value = False And optABO(1).value = False And _
       optABO(2).value = False And optABO(3).value = False Then
        MsgBox "혈액형을 선택하세요.", vbInformation, "정보확인"
        Exit Function
    End If
    
    If optRh(0).value = False And optRh(1).value = False Then
        MsgBox "Rh+, Rh- 를 선택하세요.", vbInformation, "정보확인"
        Exit Function
    End If
    Save_Check = True
End Function
Private Sub cmdSave_Click()
    Dim strSeq      As String
    Dim blnUpdateFg As Boolean
    Dim arySSN()    As String
    Dim strSSN      As String
    Dim aryZipCd()  As String
    Dim strZipCd    As String
    
    Dim SSQL        As String
    Dim objBg       As clsBeginTrans
    
    
    If Save_Check = False Then Exit Sub
    
    Set objBg = New clsBeginTrans

On Error GoTo SAVE_ERROR

    DBConn.BeginTrans
    
    arySSN = Split(txtSSN.Text, "-")
    strSSN = arySSN(0) & arySSN(1)
    If txtZipCd.Text <> "" Then
        aryZipCd = Split(txtZipCd.Text, "-")
        strZipCd = aryZipCd(0) & aryZipCd(1)
    End If
    
    '디비에 넣기전에. Seq 값 세팅
    blnUpdateFg = IIf(GetNoGiveInfo, True, False)
    
    If cmdSave.tag = "1" Then
        '헌혈자마스터 수정...............
        SSQL = objBg.SetDonorMST(DonorMstUpdateChk(lblDonorCd.Caption), lblDonorCd.Caption, Trim(txtDonorNm.Text), _
                                 strSSN, Format(dtpDOB.value, PRESENTDATE_FORMAT), Mid(lblSex.Caption, 1, 1), _
                                 strZipCd, Trim(txtAddr1.Text), Trim(txtAddr2.Text), Trim(txtTelNo.Text), _
                                 GetJobCd, GetABO, GetRh, Val(Trim(txtCnt.Text)), Val(Trim(txtTotVol.Text)))
        DBConn.Execute SSQL
        
    Else
    '저장모드인 경우
        
        strSeq = GetNoGiveSeq
        '헌혈자마스터저장.............
        SSQL = objBg.SetDonorMST(DonorMstUpdateChk(lblDonorCd.Caption), strSeq, Trim(txtDonorNm.Text), _
                                 strSSN, Format(dtpDOB.value, PRESENTDATE_FORMAT), Mid(lblSex.Caption, 1, 1), _
                                 strZipCd, Trim(txtAddr1.Text), Trim(txtAddr2.Text), Trim(txtTelNo.Text), _
                                 GetJobCd, GetABO, GetRh, Val(Trim(txtCnt.Text)), Val(Trim(txtTotVol.Text)))
        DBConn.Execute SSQL
        '번호부여정보수정............
        SSQL = objBg.SetNoGiveInfo(blnUpdateFg, BN_DONOR_ID, Val(strSeq))
        DBConn.Execute SSQL
        
    End If
    
    DBConn.CommitTrans
    Call FormInitialize
    Exit Sub
    
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리 되지 않았습니다.", vbInformation, "정보확인"
    Set objBg = Nothing
End Sub


Private Function DonorMstUpdateChk(ByVal donorid As String) As Boolean
'헌혈자 등록마스터에서 업데이트 체크
   
    Dim Rs As New Recordset
    Dim objDonor As clsBBSSQLStatement
    
    Set objDonor = New clsBBSSQLStatement
    With objDonor
'        .setDbConn DBConn
        Rs.Open .GetDonorMst(lblDonorCd.Caption), DBConn
    End With
    
    If Rs.EOF Then DonorMstUpdateChk = False: Exit Function
    Do Until Rs.EOF
        If lblDonorCd.Caption = Rs.Fields("donorid").value & "" Then
            DonorMstUpdateChk = True
            Set Rs = Nothing
            Set objDonor = Nothing
            Exit Function
        Else
            DonorMstUpdateChk = False
        End If
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    Set objDonor = Nothing
End Function

Private Function GetNoGiveInfo() As Boolean
'번호부여 정보 업데이트 체크
    
    Dim Rs          As Recordset
    Dim objNoGive   As clsBBSSQLStatement
    Dim arySQL(1)   As String
    
    Set objNoGive = New clsBBSSQLStatement
    Set Rs = New Recordset
    With objNoGive
        Rs.Open .GetNoGiveInfo(BN_DONOR_ID), DBConn
    End With
           
    If Rs.EOF Then
        '필드가 존재하지 않는 경우 Insert 실행
        arySQL(0) = objNoGive.SetNoGiveInfo(False, BN_DONOR_ID, 0)
        Call InsertData(arySQL, False)
    End If
    
    GetNoGiveInfo = True
    
    Set Rs = Nothing
    Set objNoGive = Nothing
End Function

Private Function GetNoGiveSeq() As String
'번호부여 정보에서 최고값을 얻어온다.

    Dim Rs As New Recordset
    Dim objMaxSeq As clsBBSSQLStatement
    
    Set objMaxSeq = New clsBBSSQLStatement
    With objMaxSeq
'        .setDbConn DBConn
        Rs.Open .GetNoGiveMaxSeq(BN_DONOR_ID), DBConn
    End With
    
    If Rs.EOF Then
        GetNoGiveSeq = 1
    Else
        GetNoGiveSeq = Val(Rs.Fields("maxseq").value & "") + 1
    End If
    
    Set Rs = Nothing
    Set objMaxSeq = Nothing
End Function

Private Function GetABO() As String
'디비에 넣을수 있는 혈액형으로 변환

    If optABO(0) Then GetABO = "A": Exit Function
    If optABO(1) Then GetABO = "B": Exit Function
    If optABO(2) Then GetABO = "O": Exit Function
    If optABO(3) Then GetABO = "AB": Exit Function
End Function

Private Function GetRh() As String
'디비에 넣을수 있는 Rh로 변환
    
    If optRh(0) Then GetRh = "+": Exit Function
    If optRh(1) Then GetRh = "-": Exit Function
End Function

Private Function GetJobCd() As String
'디비에 넣을 수 있는 JobCode로 변환
    
    Dim Rs As New Recordset
    Dim objJobCd As clsBBSSQLStatement
    Dim strJobNm As String
    Dim iTmx As ListItem
    
    For Each iTmx In lvwJob.ListItems
        If iTmx.Checked Then
            strJobNm = iTmx.Text
            
            Set objJobCd = New clsBBSSQLStatement
            With objJobCd
'                .setDbConn DBConn
                Set Rs = Nothing
                Set Rs = New Recordset
                Rs.Open .GetJobCd(strJobNm), DBConn
            End With
            
            GetJobCd = Rs.Fields("cdval1").value & ""
            
            Set Rs = Nothing
            Set objJobCd = Nothing
            
            Exit For
        End If
    Next
End Function

Private Sub cmdZipCd_Click()
'우편번호찾기
    Dim objZipCd As New clsZipCdFind
    
    blnModify = False
    With objZipCd
'        .setDbConn DBCONN
        Call .FormShow
        txtZipCd.Text = .ZIPCD
        txtAddr1.Text = .Province & Space(3) & .District & Space(3) & .Village
        txtAddr2.Text = .AddrNo
    End With
    blnModify = True
    Set objZipCd = Nothing
    
End Sub

Private Sub dtpDOB_Change()
    Dim lngAge As Long
       
    lngAge = DateDiff("yyyy", Format(dtpDOB.value, "yyyy-MM-dd"), Format(GetSystemDate, "yyyy-MM-dd"))
    
    If lngAge < 0 Then
        '메시지가 두번뜨는 버그 수정해야된다.
        MsgBox "날짜 설정을 다시하세요.", vbInformation, "정보확인"
'        dtpDOB.Value = Mid(txtSSN.Text, 1, 2) & "-" & Mid(txtSSN.Text, 3, 2) & "-" & Mid(txtSSN.Text, 5, 2)
        dtpDOB.SetFocus
        Exit Sub
    End If
    
    lblSex.Caption = Mid(lblSex.Caption, 1, 2) & lngAge
End Sub

Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
    blnModify = False
End Sub

Private Sub Form_Load()
    Call FormInitialize
End Sub

Private Sub FormInitialize()
    Dim i As Long
    
    blnModify = False
    txtDonorNm.Enabled = True
    txtDonorNm.Text = ""
    lblDonorCd.Caption = ""
    cboDonor.Clear
    cboDonor.Visible = False
    txtSSN.Text = ""
    dtpDOB.value = GetSystemDate
    lblSex.Caption = ""
    txtZipCd.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtTelNo.Text = ""
    
    Call GetJobNm
  
    For i = 0 To 3
        optABO(i).value = False
    Next
    optRh(0).value = False: optRh(1).value = False
    txtCnt.Text = ""
    txtTotVol.Text = ""
    cmdSave.Caption = "저장(&S)"
    cmdSave.tag = ""
End Sub

Private Sub GetJobNm()
'직업명을 얻어온다.

    Dim Rs As New Recordset
    Dim objJobNm As clsBBSSQLStatement
    Dim iTmx As ListItem
    
    
    lvwJob.tag = -1
    
    Set objJobNm = New clsBBSSQLStatement
    
    With objJobNm
'        .setDbConn DBConn
        Rs.Open .GetJobCd, DBConn
    End With
        
    lvwJob.ListItems.Clear
    Do Until Rs.EOF
        Set iTmx = lvwJob.ListItems.Add(, , Rs.Fields("field1").value & "")
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    Set objJobNm = Nothing
End Sub

Private Sub lvwJob_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Item.Selected = True
    
    If Val(lvwJob.tag) > 0 And Val(lvwJob.tag) <> Item.Index Then
        lvwJob.ListItems(Val(lvwJob.tag)).Checked = False
    End If
    
    If Item.Checked Then
        lvwJob.tag = Item.Index
    Else
        lvwJob.tag = "-1"
    End If

End Sub

Private Sub lvwJob_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = Not (Item.Checked)
    Call lvwJob_ItemCheck(Item)
End Sub

Private Sub txtDonorNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtDonorNm.Text) = "" Then Exit Sub
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDonorNm_LostFocus()
'

    Dim rs1 As New Recordset
    Dim rs2 As New Recordset
    Dim rs3 As New Recordset
    Dim objDonor As clsBBSSQLStatement
    Dim lngCnt As Long  '같은 이름을 가진 헌혈자의 수
    Dim strTmp As String
    Dim strSSN As String
    Dim strMsg As VbMsgBoxResult
    
    If txtDonorNm.Text = "" Then Exit Sub
    If blnModify = True Then Exit Sub
    Set objDonor = New clsBBSSQLStatement

    With objDonor

        rs1.Open .GetDonorMst(, Trim(txtDonorNm.Text)), DBConn
    End With

    If rs1.EOF = False Then
    '동명의 헌혈자가 있는 경우
    
        strMsg = MsgBox("이미 등록된 헌혈자입니다. 같은 이름으로 새 헌혈자를 등록합니까?", vbQuestion + vbYesNo, "정보확인")
        
        If strMsg = vbYes Then
        '새 헌혈자 등록
            cmdSave.tag = ""
            cmdSave.Caption = "저장(&S)"
        Else
            rs2.Open objDonor.GetDonorCnt(Trim(txtDonorNm.Text)), DBConn
            lngCnt = rs2.Fields("cnt").value & ""
            
            If lngCnt > 1 Then
            '동명이인인 경우 ComboBox에 보여준다.
                cboDonor.Visible = True
                cboDonor.Clear
    '            Set rs3 = OpenRecordSet(objDonor.GetDonorMst(, Trim(txtDonorNm.Text)))
                Do Until rs1.EOF
                    strSSN = Mid(rs1.Fields("ssn").value & "", 1, 6) & "-" & Mid(rs1.Fields("ssn").value & "", 7)
                    strTmp = strSSN & Space(3) & rs1.Fields("donornm").value & ""
                    cboDonor.AddItem strTmp
                    rs1.MoveNext
                Loop
                MsgBox "동명의 헌혈자가 2인 이상입니다." & vbNewLine & "수정할 헌혈자를 리스트에서 확인하세요.", vbInformation, "정보확인"
                txtDonorNm.Enabled = False
                cboDonor.SetFocus
                cboDonor.ListIndex = 0
            Else
            '같은 이름의 헌혈자가 없는 경우 헌혈자 정보를 화면에 보여준다.
                Call ShowDonorValue(Trim(txtDonorNm.Text))
                cmdSave.tag = 1
                cmdSave.Caption = "수정"
            End If
        End If
    End If
    
    Set rs2 = Nothing
    Set rs1 = Nothing
    Set objDonor = Nothing
End Sub

Private Sub ShowDonorValue(ByVal DonorNm As String, Optional ByVal SSN As String = "")
'화면에 헌혈자에 대한 정보를 보여준다.

    Dim Rs As New Recordset
    Dim rs1 As New Recordset
    Dim objDonorInfo As clsBBSSQLStatement
    Dim strSSN As String
    Dim strZipCd As String
    Dim lngAge As Long
    Dim strJobCd As String
    Dim itmFound As ListItem
    Dim iTmx As ListItem
    
    Set objDonorInfo = New clsBBSSQLStatement
    
    With objDonorInfo
'        .setDbConn DBConn
        If SSN = "" Then
            Rs.Open .GetDonorMst(, DonorNm), DBConn
        Else
            Rs.Open .GetDonorMstBySSN(DonorNm, SSN), DBConn
        End If
    End With
    
    'txtDonorNm.Enabled = False
    txtDonorNm.Text = DonorNm: txtDonorNm.Enabled = True
    lblDonorCd.Caption = Rs.Fields("donorid").value & ""
    strSSN = Mid(Rs.Fields("ssn").value & "", 1, 6) & "-" & _
             Mid(Rs.Fields("ssn").value & "", 7)
    txtSSN.Text = strSSN
    dtpDOB.value = Format(Rs.Fields("dob").value & "", "####-##-##")
    lngAge = DateDiff("yyyy", Format(dtpDOB.value, "yyyy-MM-dd"), Format(GetSystemDate, "yyyy-MM-dd"))
    lblSex.Caption = Rs.Fields("sex").value & "" & "/" & lngAge
    strZipCd = Mid(Rs.Fields("zipcd").value & "", 1, 3) & "-" & Mid(Rs.Fields("zipcd").value & "", 4)
    txtZipCd.Text = strZipCd
    txtAddr1.Text = Rs.Fields("addr1").value & ""
    txtAddr2.Text = Rs.Fields("addr2").value & ""
    txtTelNo.Text = Rs.Fields("telno").value & ""
    
    '직업코드
        
    strJobCd = Trim(Rs.Fields("jobcd").value & "")

    rs1.Open objDonorInfo.GetJobNm(strJobCd), DBConn

    
    If rs1.EOF Then
        Set rs1 = Nothing
        Set objDonorInfo = Nothing
        Exit Sub
    End If
    
    
    Set itmFound = lvwJob.FindItem(Trim(rs1.Fields("field1").value & ""))
    
    For Each iTmx In lvwJob.ListItems
        iTmx.Checked = False
    Next
    
    If Not itmFound Is Nothing Then
        itmFound.Checked = True
        lvwJob.tag = itmFound.Index
    End If

    Set rs1 = Nothing
    
    Select Case Trim(Rs.Fields("abo").value & "")
        Case "A"
            optABO(0).value = True
        Case "B"
            optABO(1).value = True
        Case "O"
            optABO(2).value = True
        Case "AB"
            optABO(3).value = True
    End Select
    
    Select Case Rs.Fields("rh").value & ""
        Case "+"
            optRh(0).value = True
        Case "-"
            optRh(1).value = True
    End Select
    
    txtCnt.Text = Rs.Fields("cnt").value & ""
    txtTotVol.Text = Rs.Fields("totvol").value & ""
    
    Set Rs = Nothing
    Set objDonorInfo = Nothing
End Sub

Private Sub txtSSN_Change()
    Dim lngLen As Long
    
    With txtSSN
        lngLen = Len(Trim(.Text))
        If lngLen = 6 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtSSN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSSN.Text) = "" Then Exit Sub
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSSN_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then
        With txtSSN
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtSSN_LostFocus()
    Dim strDOB As String
    Dim lngAge As Long
    
    If Trim(txtSSN.Text) = "" Then Exit Sub
    strDOB = Mid(txtSSN, 1, 6)
    If IsDate(Format(strDOB, "##-##-##")) = False Then GoTo InputError
    
    dtpDOB = Format(strDOB, "##-##-##")
    lngAge = DateDiff("yyyy", Format(dtpDOB.value, "yyyy-MM-dd"), Format(GetSystemDate, "yyyy-MM-dd"))
    Select Case Mid(txtSSN.Text, 8, 1)
        Case 1, 3
            lblSex.Caption = "M" & "/" & lngAge
        Case 2, 4
            lblSex.Caption = "F" & "/" & lngAge
    End Select
    Exit Sub
    
InputError:
    MsgBox "유효한 주민번호가 아닙니다. 다시 입력하세요.", vbInformation, "정보확인"
    With txtSSN
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtZipCd_Change()

    Dim lngLen As Long
    
    With txtZipCd
        lngLen = Len(Trim(.Text))
        If lngLen = 3 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtZipCd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then
        With txtZipCd
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If

End Sub
