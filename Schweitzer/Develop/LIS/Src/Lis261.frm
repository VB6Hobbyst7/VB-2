VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm261MDefDate 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "시간"
      Height          =   3030
      Left            =   1875
      TabIndex        =   4
      Top             =   615
      Width           =   1230
      Begin VB.OptionButton Option12 
         BackColor       =   &H00DBE6E6&
         Caption         =   "14:00:00"
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   1965
         Width           =   1000
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00DBE6E6&
         Caption         =   "10:00:00"
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   1620
         Width           =   1095
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00DBE6E6&
         Caption         =   "(N-2h)"
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   1290
         Width           =   1000
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00DBE6E6&
         Caption         =   "(N-1h)"
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   975
         Width           =   1000
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00DBE6E6&
         Caption         =   "정각(N)"
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   1065
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00DBE6E6&
         Caption         =   "현재"
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   300
         Width           =   1000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "날짜"
      Height          =   3030
      Left            =   180
      TabIndex        =   3
      Top             =   615
      Width           =   1620
      Begin VB.OptionButton optDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일년전"
         Height          =   240
         Index           =   5
         Left            =   165
         TabIndex        =   11
         Top             =   1995
         Width           =   1125
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "한달전"
         Height          =   240
         Index           =   4
         Left            =   165
         TabIndex        =   10
         Top             =   1650
         Width           =   1125
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일주전(T-7)"
         Height          =   240
         Index           =   3
         Left            =   165
         TabIndex        =   9
         Top             =   1320
         Width           =   1305
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "그제(T-2)"
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   8
         Top             =   1005
         Width           =   1125
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "어제(T-1)"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   7
         Top             =   660
         Width           =   1125
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "오늘(T)"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   5
         Top             =   330
         Width           =   1125
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   165
      TabIndex        =   2
      Top             =   135
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      BackColor       =   14676956
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "취소"
      Height          =   510
      Left            =   1680
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   3870
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F4F0F2&
      Caption         =   "결정"
      Height          =   510
      Left            =   360
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   3870
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   150
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BackColor       =   16709613
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Appearance      =   0
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   3150
      Y1              =   3780
      Y2              =   3780
   End
End
Attribute VB_Name = "frm261MDefDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fF As Form
Dim fTxtDt As Object
Dim fTxtTm As Object

Private Sub cmdCancel_Click()
   Unload Me
End Sub

'public sub SetPosition(byval p
Private Sub Form_Load()

End Sub

Public Sub SetInitValue(pF As Form, pTxtDt As Object, pTxtTm As Object, _
                        ByVal pIniDt As Integer, ByVal pIniTm As Integer)
   Set fF = pF
   Set fTxtDt = pTxtDt
   Set fTxtTm = pTxtTm
   
   Me.Top = 2175
   Me.Left = 6255
      
End Sub
                        

