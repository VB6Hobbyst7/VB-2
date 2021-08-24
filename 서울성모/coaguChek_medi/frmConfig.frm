VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3870
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSPanel SSPanel3 
      Height          =   1515
      Left            =   60
      TabIndex        =   4
      Top             =   2490
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   2672
      _Version        =   131072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.TextBox txtPort2 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Caption         =   "종 료"
         Height          =   345
         Left            =   2640
         TabIndex        =   8
         Top             =   990
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "확 인"
         Height          =   345
         Left            =   1650
         TabIndex        =   7
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Socket Port"
         Height          =   195
         Index           =   15
         Left            =   270
         TabIndex        =   6
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "Win2"
         Height          =   225
         Left            =   3060
         TabIndex        =   5
         Top             =   150
         Width           =   855
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   2672
      _Version        =   131072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.TextBox txtPort1 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Width           =   1965
      End
      Begin VB.CommandButton Command1 
         Caption         =   "확 인"
         Height          =   345
         Left            =   1650
         TabIndex        =   11
         Top             =   990
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "종 료"
         Height          =   345
         Left            =   2640
         TabIndex        =   10
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Win1"
         Height          =   225
         Left            =   3090
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Socket Port"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   1335
      End
   End
   Begin Threed.SSPanel SSPanel_machine 
      Height          =   675
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   1191
      _Version        =   131072
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "coaguCheck"
      BevelInner      =   1
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Serial Setting

Private Sub Command1_Click()
    Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
    
        gSetup.gPort = txtPort1.Text
        Call WritePrivateProfileString("config1", "gPort", gSetup.gPort, App.Path & "\Interface.ini")

        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
'    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
'        Exit Sub
'    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        gSetup2.gPort = txtPort2.Text
        Call WritePrivateProfileString("config2", "gPort", gSetup2.gPort, App.Path & "\Interface.ini")
        Unload Me
    End If
    Exit Sub
 
ErrorHandler:
    Resume Next

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Ret As Integer

'    Combo_Port.ListIndex = 0
    
    txtPort1.Text = gSetup.gPort
    txtPort2.Text = gSetup2.gPort
    
End Sub

