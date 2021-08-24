VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmSlideDelete 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "이미지 삭제"
   ClientHeight    =   4125
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   7335
   Icon            =   "frmSlideDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7335
   StartUpPosition =   1  '소유자 가운데
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1395
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "접수번호 "
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   330
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "환자   ID"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblReceptNo 
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "이미지경로"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   7
      Left            =   2640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   690
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "성      명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1755
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "선택이미지"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "취소(&U)"
      Height          =   510
      Left            =   3825
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   3375
      Width           =   1320
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "확인(&O)"
      Height          =   510
      Left            =   2490
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   3375
      Width           =   1320
   End
   Begin VB.PictureBox picSlide 
      BackColor       =   &H00DBE6E6&
      Height          =   2520
      Left            =   150
      ScaleHeight     =   2460
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   330
      Width           =   2475
      Begin VB.Image imgDel 
         Height          =   2265
         Left            =   90
         Picture         =   "frmSlideDelete.frx":0442
         Stretch         =   -1  'True
         Top             =   90
         Width           =   2265
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00DBE6E6&
      Caption         =   "삭제하실려면 확인 버튼을 눌러 주십시요."
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3420
      TabIndex        =   10
      Top             =   2610
      Width           =   3435
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSexAge 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   4965
      TabIndex        =   9
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "주의) 이미지를 삭제합니다."
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3420
      TabIndex        =   8
      Top             =   2340
      Width           =   3435
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAccNo 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   3780
      TabIndex        =   7
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label lblPicName 
      BackColor       =   &H00DBE6E6&
      Height          =   285
      Left            =   3780
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblimgPath 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   3780
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblPtId 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   3735
      TabIndex        =   4
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label lblPtNm 
      BackColor       =   &H00DBE6E6&
      Height          =   300
      Left            =   3780
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmSlideDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnFirst As Boolean
Public Event ImageDeleteLoad()
Public Event ImageDelete()

Private Sub cmdClear_Click()
   '
   Unload Me
   '
End Sub

Private Sub cmdDelete_Click()
   '
   RaiseEvent ImageDelete
   Unload Me
   '
End Sub

Private Sub Form_Activate()
   '
   If blnFirst = True Then
      Me.MousePointer = 13
      DoEvents
      LockWindowUpdate (Me.hwnd)
      RaiseEvent ImageDeleteLoad
      LockWindowUpdate (0&)
      Me.MousePointer = 1
      blnFirst = False
   End If
   '
End Sub

Private Sub Form_Load()
Dim ii As Long
   '
   blnFirst = True
   Set imgDel.Picture = Nothing
   For ii = 1 To 3
      Beep
   Next
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSlideDelete = Nothing
End Sub
