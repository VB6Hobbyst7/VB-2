VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4650
   StartUpPosition =   1  '소유자 가운데
   Begin VB.OptionButton optType 
      Caption         =   "Client"
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   11
      Top             =   570
      Width           =   975
   End
   Begin VB.OptionButton optType 
      Caption         =   "Server"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   570
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Text            =   "5050"
      Top             =   1500
      Width           =   1815
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Text            =   "127.0.0.1"
      Top             =   990
      Width           =   1815
   End
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2010
      TabIndex        =   5
      Top             =   4260
      Width           =   975
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
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4890
      Width           =   1095
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
      Height          =   375
      Left            =   1950
      TabIndex        =   0
      Top             =   4890
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "일 저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   8
      Left            =   3120
      TabIndex        =   6
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "로컬저장기간"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   7
      Left            =   540
      TabIndex        =   4
      Top             =   4320
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   285
      TabIndex        =   3
      Top             =   1035
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   585
      Width           =   1305
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click()
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
    
        gComm.TCPIP = txtIP.Text
        gComm.TCPPORT = txtPort.Text
        
        If optType(0).Value = True Then
            gComm.TCPTYPE = "SERVER"
        Else
            gComm.TCPTYPE = "CLIENT"
        End If

        Call WritePrivateProfileString("COMM", "TCPTYPE", gComm.TCPTYPE, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "TCPIP", gComm.TCPIP, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "TCPPORT", gComm.TCPPORT, App.PATH & "\INI\" & gMACH & ".ini")
        
        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    If gComm.TCPTYPE = "SERVER" Then
        optType(0).Value = True
    Else
        optType(1).Value = True
    End If
    
    txtIP.Text = gComm.TCPIP
    txtPort.Text = gComm.TCPPORT
    
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub
