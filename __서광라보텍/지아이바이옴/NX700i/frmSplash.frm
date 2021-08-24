VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer2 
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Left            =   4050
      Top             =   120
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   2895
      Left            =   30
      Top             =   30
      Width           =   5205
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   2
      Top             =   1830
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMachNm 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "LAB INTERFACE PROGRAM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1335
      Left            =   2700
      TabIndex        =   1
      Top             =   600
      Width           =   2085
      WordWrap        =   -1  'True
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
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   360
      Picture         =   "frmSplash.frx":0000
      Top             =   300
      Width           =   2280
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
