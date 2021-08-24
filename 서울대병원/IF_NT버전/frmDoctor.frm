VERSION 5.00
Begin VB.Form frmDoctor 
   Caption         =   "담당의 설정"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7875
   Icon            =   "frmDoctor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdRet 
      Caption         =   "적용"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6660
      TabIndex        =   11
      Top             =   1530
      Width           =   855
   End
   Begin VB.TextBox txtUseDoctorName 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5580
      TabIndex        =   9
      Top             =   2220
      Width           =   1935
   End
   Begin VB.CheckBox chkUse 
      Caption         =   "사용"
      Height          =   405
      Left            =   5610
      TabIndex        =   7
      Top             =   1500
      Width           =   765
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "삭제"
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
      Left            =   6390
      TabIndex        =   4
      Top             =   2910
      Width           =   1125
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "추가"
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
      Left            =   5250
      TabIndex        =   3
      Top             =   2910
      Width           =   1125
   End
   Begin VB.TextBox txtDoctorCode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5580
      TabIndex        =   2
      Top             =   1110
      Width           =   1935
   End
   Begin VB.TextBox txtDoctorName 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5580
      TabIndex        =   1
      Top             =   690
      Width           =   1935
   End
   Begin VB.ListBox lstDoctor 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   210
      TabIndex        =   0
      Top             =   660
      Width           =   4005
   End
   Begin VB.Label Label4 
      Caption         =   "등록된 담당의 리스트"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   10
      Top             =   330
      Width           =   2355
   End
   Begin VB.Label Label3 
      Caption         =   "현재 담당의"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   8
      Top             =   2250
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "담당의 사번"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4380
      TabIndex        =   6
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "담당의 이름"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   5
      Top             =   720
      Width           =   1185
   End
End
Attribute VB_Name = "frmDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strDoctor   As String

Private Sub cmdConfirm_Click()
    
    lstDoctor.AddItem txtDoctorName.Text & Space(20 - LenB(txtDoctorName.Text)) & "|" & txtDoctorCode.Text

    Call SetDoctList

End Sub

Private Sub cmdExit_Click()
    
    lstDoctor.RemoveItem lstDoctor.ListIndex
    
    Call SetDoctList
    
End Sub

Private Sub cmdRet_Click()
    
    strDoctor = txtDoctorName.Text & "|" & txtDoctorCode.Text
    
    If chkUse.Value = "1" Then
        Call WritePrivateProfileString("DOCTOR", "DOCTOR", strDoctor, App.PATH & "\INI\" & gMACH & ".ini")
    End If

End Sub

Private Sub Form_Load()
    
    lstDoctor.Clear
    
    Call GetDoctList
    
End Sub

Public Sub SetDoctList()
    Dim i   As Integer
    Dim strDoctor   As String
    
    strDoctor = txtDoctorName.Text & "|" & txtDoctorCode.Text
    
    If chkUse.Value = "1" Then
        Call WritePrivateProfileString("DOCTOR", "DOCTOR", strDoctor, App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    Call WritePrivateProfileString("DOCTOR", "DOCTORCOUNT", CStr(lstDoctor.ListCount), App.PATH & "\INI\" & gMACH & ".ini")
    
    For i = 0 To lstDoctor.ListCount - 1
        strDoctor = lstDoctor.List(i)
        strDoctor = Trim(mGetP(strDoctor, 1, "|")) & "|" & Trim(mGetP(strDoctor, 2, "|"))
        Call WritePrivateProfileString("DOCTOR", "DOCTOR" & CStr(i + 1), strDoctor, App.PATH & "\INI\" & gMACH & ".ini")
    Next
        
End Sub


Public Sub GetDoctList()
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    
    
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DOCTOR", "DOCTOR", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    strDoctor = Trim(strSetUp1)
    
    txtUseDoctorName.Text = strDoctor
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DOCTOR", "DOCTORCOUNT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    J = Trim(strSetUp1)
    
    If IsNumeric(J) Then
        For i = 1 To J
            strSetup = "":    strSetUp1 = ""
            Call GetPrivateProfileString("DOCTOR", "DOCTOR" & CStr(i), "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
            strSetUp1 = Trim(strSetup)
            strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
            'gMACHS(i) = Trim(strSetUp1)
            If strDoctor = strSetUp1 Then
                k = i
            End If
            lstDoctor.AddItem mGetP(Trim(strSetUp1), 1, "|") & Space(20 - LenB(mGetP(Trim(strSetUp1), 1, "|"))) & "|" & mGetP(Trim(strSetUp1), 2, "|")
        Next
    End If
    
    
End Sub

Private Sub lstDoctor_Click()
    
    txtDoctorName.Text = Trim(mGetP(lstDoctor.Text, 1, "|"))
    txtDoctorCode.Text = Trim(mGetP(lstDoctor.Text, 2, "|"))
    
    If Trim(mGetP(strDoctor, 1, "|")) = Trim(mGetP(lstDoctor.Text, 1, "|")) And Trim(mGetP(strDoctor, 2, "|")) = Trim(mGetP(lstDoctor.Text, 2, "|")) Then
        chkUse.Value = "1"
    Else
        chkUse.Value = "0"
    End If
End Sub
