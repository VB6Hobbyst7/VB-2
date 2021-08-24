VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8505
   StartUpPosition =   1  '소유자 가운데
   Begin VB.OptionButton optAPIURL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "운영서버"
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   16
      Top             =   1800
      Width           =   1245
   End
   Begin VB.OptionButton optAPIURL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "교육서버"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   15
      Top             =   1380
      Width           =   1245
   End
   Begin VB.OptionButton optAPIURL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "기준서버"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   14
      Top             =   990
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.TextBox txtOPRURL 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Text            =   "127.0.0.1"
      Top             =   1740
      Width           =   6615
   End
   Begin VB.TextBox txtEDUURL 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Text            =   "127.0.0.1"
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox txtSTDURL 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   "127.0.0.1"
      Top             =   930
      Width           =   6615
   End
   Begin VB.TextBox txtAPIURL 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1575
      TabIndex        =   9
      Text            =   "127.0.0.1"
      Top             =   450
      Width           =   6615
   End
   Begin VB.TextBox txtOrderPath 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2370
      TabIndex        =   6
      Text            =   "127.0.0.1"
      Top             =   2340
      Width           =   3315
   End
   Begin VB.TextBox txtResultPath 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2370
      TabIndex        =   5
      Text            =   "5050"
      Top             =   2850
      Width           =   3315
   End
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   3360
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
      Left            =   7020
      TabIndex        =   1
      Top             =   4050
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
      Left            =   5850
      TabIndex        =   0
      Top             =   4050
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사용서버"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   525
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "CFX96 오더경로"
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
      Left            =   705
      TabIndex        =   8
      Top             =   2385
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "CFX96 결과경로"
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
      Left            =   660
      TabIndex        =   7
      Top             =   2910
      Width           =   1485
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
      Left            =   3450
      TabIndex        =   4
      Top             =   3420
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
      Left            =   870
      TabIndex        =   2
      Top             =   3420
      Width           =   1170
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
    
        Call WritePrivateProfileString("HOSP", "APIURL", txtAPIURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "STDURL", txtSTDURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "EDUURL", txtEDUURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "OPRURL", txtOPRURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        Call WritePrivateProfileString("COMM", "ORDPATH", txtOrderPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "RSTPATH", txtResultPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\INI\" & gMACH & ".ini")
                
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
    txtEDUURL.Text = gHOSP.EDUURL
    txtOPRURL.Text = gHOSP.OPRURL
    
    txtOrderPath.Text = gComm.ORDPATH
    txtResultPath.Text = gComm.RSTPATH
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub

Private Sub optAPIURL_Click(Index As Integer)
    
    Select Case Index
        Case 0:     txtAPIURL.Text = txtSTDURL.Text
        Case 1:     txtAPIURL.Text = txtEDUURL.Text
        Case 2:     txtAPIURL.Text = txtOPRURL.Text
    End Select
    
End Sub
