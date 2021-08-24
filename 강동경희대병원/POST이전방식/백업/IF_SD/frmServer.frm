VERSION 5.00
Begin VB.Form frmServer 
   Caption         =   " 서버 설정 "
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows 기본값
   Begin VB.OptionButton optAPIURL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "개발서버"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton optAPIURL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "운영서버"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   5
      Top             =   1170
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtDEVURL 
      Appearance      =   0  '평면
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
      Left            =   1470
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   1590
      Width           =   7095
   End
   Begin VB.TextBox txtSTDURL 
      Appearance      =   0  '평면
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
      Left            =   1470
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   1080
      Width           =   7095
   End
   Begin VB.TextBox txtAPIURL 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1485
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   390
      Width           =   7095
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
      Height          =   435
      Left            =   7380
      TabIndex        =   1
      Top             =   2400
      Width           =   1125
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
      Height          =   435
      Left            =   6180
      TabIndex        =   0
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "◈ 사용서버"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   465
      Width           =   1065
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click()
    Dim strUseUrl  As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        If optAPIURL(0).Value = True Then
            strUseUrl = "STD"
        Else
            strUseUrl = "DEV"
        End If
        
        Call WritePrivateProfileString("HOSP", "USEURL", strUseUrl, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "APIURL", txtAPIURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "STDURL", txtSTDURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "EDUURL", txtDEVURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("경로가 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    txtAPIURL.Text = gHOSP.APIURL
    txtSTDURL.Text = gHOSP.STDURL
    txtDEVURL.Text = gHOSP.DEVURL
    
    If gHOSP.USEURL = "STD" Then
        optAPIURL(0).Value = True
    Else
        optAPIURL(1).Value = True
    End If
        
End Sub

Private Sub optAPIURL_Click(Index As Integer)
    
    Select Case Index
        Case 0:     txtAPIURL.Text = txtSTDURL.Text
        Case 1:     txtAPIURL.Text = txtDEVURL.Text
    End Select
    
End Sub

