VERSION 5.00
Begin VB.Form frmResultTrans 
   BackColor       =   &H00FFFFFF&
   Caption         =   "°Ë»ç°á°úº¯È¯"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   Icon            =   "frmResultTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4380
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1890
      TabIndex        =   21
      Text            =   "Positive(+++++)"
      Top             =   3540
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "POS5"
      Top             =   3540
      Width           =   1035
   End
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1890
      TabIndex        =   18
      Text            =   "Positive(+)"
      Top             =   3060
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "POS4"
      Top             =   3060
      Width           =   1035
   End
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1890
      TabIndex        =   15
      Text            =   "Positive(+)"
      Top             =   2580
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "POS3"
      Top             =   2580
      Width           =   1035
   End
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1890
      TabIndex        =   12
      Text            =   "Positive(+)"
      Top             =   2100
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "POS2"
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1890
      TabIndex        =   9
      Text            =   "Positive(+)"
      Top             =   1620
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "POS"
      Top             =   1620
      Width           =   1035
   End
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1890
      TabIndex        =   6
      Text            =   "Equipvocal"
      Top             =   1140
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "EQV"
      Top             =   1140
      Width           =   1035
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
      Height          =   435
      Left            =   2880
      TabIndex        =   4
      Top             =   4380
      Width           =   1095
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
      Height          =   435
      Left            =   1710
      TabIndex        =   3
      Top             =   4380
      Width           =   1095
   End
   Begin VB.TextBox txtDest 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1890
      TabIndex        =   1
      Text            =   "Negative"
      Top             =   660
      Width           =   2085
   End
   Begin VB.TextBox txtSrc 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "NEG"
      Top             =   660
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   22
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   16
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "frmResultTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
    Dim strBP   As String
    Dim strCP   As String
    Dim strLP   As String
    Dim strMP   As String
    
    Dim strBPN   As String
    Dim strCPN   As String
    Dim strLPN   As String
    Dim strMPN   As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("¼³Á¤À» ÀúÀåÇÏ½Ã°Ú½À´Ï±î?", vbCritical + vbOKCancel + vbDefaultButton2, "È®ÀÎ!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        Call WritePrivateProfileString("COMMENT", "NEG", txtDest(0).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "EQV", txtDest(1).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "POS", txtDest(2).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "POS2", txtDest(3).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "POS3", txtDest(4).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "POS4", txtDest(5).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "POS5", txtDest(6).Text, App.PATH & "\INI\" & gMACH & ".ini")
                
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("°æ·Î°¡ ¸ÂÁö ¾Ê½À´Ï´Ù", vbCritical + vbOKCancel + vbDefaultButton2, "Á¾·á¹öÆ°") = vbCancel Then
        Unload Me
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
            
    txtDest(0).Text = gCmnt.NEG
    txtDest(1).Text = gCmnt.EQV
    txtDest(2).Text = gCmnt.POS
    txtDest(3).Text = gCmnt.POS2
    txtDest(4).Text = gCmnt.POS3
    txtDest(5).Text = gCmnt.POS4
    txtDest(6).Text = gCmnt.POS5

End Sub
