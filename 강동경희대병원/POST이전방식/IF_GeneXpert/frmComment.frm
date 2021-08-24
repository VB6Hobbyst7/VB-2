VERSION 5.00
Begin VB.Form frmComment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "코멘트 설정"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   Icon            =   "frmComment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7665
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2970
      Width           =   5000
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      TabIndex        =   9
      Top             =   3750
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6060
      TabIndex        =   8
      Top             =   3750
      Width           =   1095
   End
   Begin VB.TextBox txtSGCarbaR 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2310
      Width           =   5000
   End
   Begin VB.TextBox txtSGMTBRIF 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1710
      Width           =   5000
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7665
      TabIndex        =   1
      Top             =   0
      Width           =   7665
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "코멘트 설정"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.TextBox txtSGCovid19 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1110
      Width           =   5000
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "결과경로"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   3090
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Carba-R 코멘트"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "MTB/RIF 코멘트"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Covid19 코멘트"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   1290
      Width           =   1260
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirm_Click()
    Dim strSGCovid19    As String
    Dim strSDCovid19    As String
    Dim strSGRV16       As String
    Dim strSGPB5        As String
    

    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        strSGCovid19 = Replace(txtSGCovid19.Text, vbCrLf, "CHR(10)CHR(13)")
        strSGRV16 = Replace(txtSGMTBRIF.Text, vbCrLf, "CHR(10)CHR(13)")
        strSGPB5 = Replace(txtSGCarbaR.Text, vbCrLf, "CHR(10)CHR(13)")
        
        Call WritePrivateProfileString("COMMENT", "SGCOVID", strSGCovid19, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "MTBRIF", strSGRV16, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "CARBAR", strSGPB5, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "PATH", txtPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
                
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    txtSGCovid19.Text = Replace(gCFXCmnt.SGCOVID, "CHR(10)CHR(13)", vbCrLf)
    txtSGMTBRIF.Text = Replace(gCFXCmnt.MTBRIF, "CHR(10)CHR(13)", vbCrLf)
    txtSGCarbaR.Text = Replace(gCFXCmnt.CARBAR, "CHR(10)CHR(13)", vbCrLf)
    txtPath.Text = gCFXCmnt.PATH
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

