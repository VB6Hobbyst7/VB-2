VERSION 5.00
Begin VB.Form frmComment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÄÚ¸àÆ® ¼³Á¤"
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
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      TabIndex        =   11
      Top             =   2970
      Width           =   5000
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ÀúÀå"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      TabIndex        =   10
      Top             =   3750
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Ãë¼Ò"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6060
      TabIndex        =   9
      Top             =   3750
      Width           =   1095
   End
   Begin VB.TextBox txtSGRV16 
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.TextBox txtSDCovid19 
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Top             =   1110
      Width           =   5000
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'À§ ¸ÂÃã
      BackColor       =   &H00808000&
      BorderStyle     =   0  '¾øÀ½
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
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "ÄÚ¸àÆ® ¼³Á¤"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
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
         Name            =   "±¼¸²"
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
      Top             =   1710
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°á°ú°æ·Î"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   3090
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¾¾Á¨ RV16 ÄÚ¸àÆ®"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Covid19 ÄÚ¸àÆ®"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¾¾Á¨ Covid19 ÄÚ¸àÆ®"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1680
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirm_Click()
    Dim strSDCovid19    As String
    

    If MsgBox("¼³Á¤À» ÀúÀåÇÏ½Ã°Ú½À´Ï±î?", vbCritical + vbOKCancel + vbDefaultButton2, "È®ÀÎ!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        strSDCovid19 = Replace(txtSDCovid19.Text, vbCrLf, "CHR(10)CHR(13)")
        
        Call WritePrivateProfileString("COMMENT", "SDCOVID", strSDCovid19, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "PATH", txtPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
                
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    txtSDCovid19.Text = Replace(gCFXCmnt.SDCOVID, "CHR(10)CHR(13)", vbCrLf)
    txtPath.Text = gCFXCmnt.PATH

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub
