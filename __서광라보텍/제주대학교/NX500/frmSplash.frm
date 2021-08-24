VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer2 
      Left            =   1380
      Top             =   1590
   End
   Begin VB.Timer Timer1 
      Left            =   1830
      Top             =   1590
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료"
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      Top             =   2460
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Interface Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   540
      TabIndex        =   4
      Top             =   900
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMachNm 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "SANSOFT"
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
      Height          =   495
      Left            =   2100
      TabIndex        =   3
      Top             =   420
      Width           =   2085
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "sansoft.kr"
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
      Left            =   540
      TabIndex        =   2
      Top             =   1260
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   2895
      Left            =   30
      Top             =   30
      Width           =   4455
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "프로그램 로딩중입니다."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   360
      TabIndex        =   0
      Top             =   2310
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   420
      Picture         =   "frmSplash.frx":0000
      Top             =   360
      Width           =   705
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        End
    End If
    
End Sub

Private Sub Timer1_Timer()
    
    
    frmSplash.labMsg.Caption = "통신포트 연결중입니다."
    DoEvents
    
    Timer1.Enabled = False
    
    Timer2.Interval = 500
    Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
    frmSplash.labMsg.Caption = "통신포트 연결중입니다."
    DoEvents
    
    Timer2.Enabled = False
    Call frmInterface.Show
    Unload frmSplash

End Sub
