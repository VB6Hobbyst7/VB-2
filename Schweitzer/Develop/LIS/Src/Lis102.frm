VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Begin VB.Form frm102Warning 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Wardning/Infection 등록"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ComboBox cboWarning 
      Height          =   300
      Left            =   6120
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   420
      Width           =   3135
   End
   Begin VB.TextBox txtPtId 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   5
      Top             =   435
      Width           =   1485
   End
   Begin VB.CommandButton cmdHelpList 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2790
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "PtID"
      Top             =   420
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "종류 (&X)"
      Height          =   405
      Left            =   8175
      TabIndex        =   1
      Top             =   885
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "저장 (&S)"
      Height          =   405
      Left            =   7035
      TabIndex        =   0
      Top             =   885
      Width           =   1095
   End
   Begin MedControls1.LisLabel lblDob 
      Height          =   315
      Left            =   4170
      TabIndex        =   3
      Top             =   810
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      BackColor       =   15857140
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblPtNm 
      Height          =   300
      Left            =   4170
      TabIndex        =   4
      Top             =   435
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BackColor       =   15857140
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ Warning / Infection"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00553755&
      Height          =   180
      Left            =   6105
      TabIndex        =   14
      Top             =   180
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   990
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   315
      Width           =   5865
   End
   Begin VB.Label lblAgeDiv 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2700
      TabIndex        =   12
      Top             =   885
      Width           =   60
   End
   Begin VB.Label lblDOB1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "생년월일"
      Height          =   180
      Left            =   3300
      TabIndex        =   11
      Tag             =   "101"
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lblAge 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2250
      TabIndex        =   10
      Top             =   885
      Width           =   345
   End
   Begin VB.Label lblSex 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1365
      TabIndex        =   9
      Top             =   885
      Width           =   690
   End
   Begin VB.Label lblSexAge 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "성별/나이"
      Height          =   255
      Left            =   255
      TabIndex        =   8
      Tag             =   "108"
      Top             =   900
      Width           =   945
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "성    명"
      Height          =   180
      Left            =   3390
      TabIndex        =   7
      Tag             =   "103"
      Top             =   495
      Width           =   600
   End
   Begin VB.Label lblPtId 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "환자 ID"
      Height          =   180
      Left            =   510
      TabIndex        =   6
      Tag             =   "105"
      Top             =   495
      Width           =   585
   End
   Begin VB.Label Label22 
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F4&
      Caption         =   "             /"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1275
      TabIndex        =   13
      Top             =   825
      Width           =   1785
   End
End
Attribute VB_Name = "frm102Warning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Private objMyList As clsS2DLP
'
'Private Sub cmdHelpList_Click(Index As Integer)
'
'    Set objMyList = New clsS2DLP
'
'    With objMyList
'        Select Case Index
'            Case 0
'                Set frmPtFind = frmPtInfo
'                frmPtFind.Show 1
'            Case 1
'                 .Caption = "처방의 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "처방의ID, 처방의명"
'                 Call .ListPop(, 1640, framPtInfo.Left + cmdHelpList(Index).Left, ObjLISComCode.doct)
'
'            Case 2
'                 .Caption = "진료과 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "진료과코드, 진료과명"
'                 Call .ListPop(, 1640, framPtInfo.Left + cmdHelpList(Index).Left, ObjLISComCode.DeptCd)
'                 If txtPtId <> "" Then
''                    PtInfoEnable False
''                    cmdHelpList(0).Enabled = False
'                 End If
'
'            Case 3
'                 .Caption = "병동 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "병동코드,병동명"
'                 Call .ListPop(, 1640, 10550, ObjLISComCode.WardId)
'
'        End Select
'    End With
'
'    Set objMyList = Nothing
'
'End Sub
