VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "산소프트 SANIF 정보"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows 기본값
   Begin HSCotrol.CButton cmdOk 
      Height          =   375
      Left            =   4230
      TabIndex        =   3
      ToolTipText     =   "조회된 워크리스트를 일괄등록한다."
      Top             =   3750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BackColor       =   16777215
      Caption         =   "확인"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   0
      HoverColor      =   16744576
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "s@n l@b interf@ce"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   525
      Left            =   4800
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Label labMaker 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "sansoft.kr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4560
      MousePointer    =   1  '화살표
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "office : 0505-299-1544 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   3
      Left            =   315
      TabIndex        =   4
      Top             =   3870
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   $"frmAbout.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      TabIndex        =   2
      Top             =   1770
      Width           =   4140
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   330
      Picture         =   "frmAbout.frx":095A
      Top             =   1290
      Width           =   705
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "SANSOFT INTERFACE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Top             =   1260
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   300
      X2              =   5380
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label lblMachNm 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "SANIF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   525
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   300
      X2              =   5385
      Y1              =   870
      Y2              =   870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub labMaker_Click()
    
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & labMaker.Caption
 
End Sub

Private Sub labMaker_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    labMaker.ForeColor = vbBlue
    
End Sub

