VERSION 5.00
Begin VB.Form frmResultTrans 
   BackColor       =   &H00FFFFFF&
   Caption         =   "°Ë»ç°á°úº¯È¯"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   Icon            =   "frmResultTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   12435
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " PB 6 "
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8220
      TabIndex        =   16
      Top             =   150
      Width           =   3975
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
         Left            =   1710
         TabIndex        =   20
         Text            =   "Positive"
         Top             =   1110
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "POS"
         Top             =   1110
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
         Left            =   1710
         TabIndex        =   18
         Text            =   "Negative"
         Top             =   630
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "NEG"
         Top             =   630
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
         Index           =   5
         Left            =   1260
         TabIndex        =   22
         Top             =   1170
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
         Left            =   1260
         TabIndex        =   21
         Top             =   690
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " RP 19 "
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4170
      TabIndex        =   9
      Top             =   150
      Width           =   3975
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
         Left            =   1710
         TabIndex        =   13
         Text            =   "Positive"
         Top             =   1110
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "POS"
         Top             =   1110
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
         Left            =   1710
         TabIndex        =   11
         Text            =   "Negative"
         Top             =   630
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "NEG"
         Top             =   630
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
         Index           =   3
         Left            =   1260
         TabIndex        =   15
         Top             =   1170
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
         Left            =   1260
         TabIndex        =   14
         Top             =   690
         Width           =   375
      End
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
      Left            =   11070
      TabIndex        =   8
      Top             =   2340
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
      Left            =   9900
      TabIndex        =   7
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Frame fraMTB 
      BackColor       =   &H00FFFFFF&
      Caption         =   " MTB "
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3975
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "NEG"
         Top             =   630
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
         Index           =   0
         Left            =   1710
         TabIndex        =   3
         Text            =   "Negative"
         Top             =   630
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "POS"
         Top             =   1110
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
         Left            =   1710
         TabIndex        =   1
         Text            =   "Positive"
         Top             =   1110
         Width           =   2085
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
         Left            =   1260
         TabIndex        =   6
         Top             =   690
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
         Left            =   1260
         TabIndex        =   5
         Top             =   1170
         Width           =   375
      End
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
        Call WritePrivateProfileString("COMMENT", "MTBNEG", txtDest(0).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "MTBPOS", txtDest(1).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "RP19NEG", txtDest(2).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "RP19POS", txtDest(3).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "PB6NEG", txtDest(4).Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMMENT", "PB6POS", txtDest(5).Text, App.PATH & "\INI\" & gMACH & ".ini")
                
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
            
    txtDest(0).Text = gCmnt.MTBNEG
    txtDest(1).Text = gCmnt.MTBPOS
    txtDest(2).Text = gCmnt.RP19NEG
    txtDest(3).Text = gCmnt.RP19POS
    txtDest(4).Text = gCmnt.PB6NEG
    txtDest(5).Text = gCmnt.PB6POS

End Sub
