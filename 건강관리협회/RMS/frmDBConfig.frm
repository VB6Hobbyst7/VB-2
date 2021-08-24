VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDBConfig 
   BorderStyle     =   1  '단일 고정
   Caption         =   "DataBase Setting"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6165
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdLIS 
      Caption         =   "LIS업체"
      Height          =   525
      Left            =   600
      TabIndex        =   18
      Top             =   2700
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4530
      TabIndex        =   1
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      TabIndex        =   0
      Top             =   2700
      Width           =   1095
   End
   Begin Threed.SSPanel sspSQL 
      Height          =   1695
      Left            =   600
      TabIndex        =   9
      Top             =   570
      Width           =   5025
      _Version        =   65536
      _ExtentX        =   8864
      _ExtentY        =   2990
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtMSPasswd 
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Passwd"
         Top             =   1185
         Width           =   2205
      End
      Begin VB.TextBox txtMSUser 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1710
         TabIndex        =   12
         Text            =   "User"
         Top             =   885
         Width           =   2205
      End
      Begin VB.TextBox txtDB 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1710
         TabIndex        =   11
         Top             =   480
         Width           =   3030
      End
      Begin VB.TextBox txtMSServer 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1710
         TabIndex        =   10
         Top             =   150
         Width           =   3030
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "서버(&S):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   855
         TabIndex        =   17
         Top             =   195
         Width           =   720
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "암호(&B):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   915
         TabIndex        =   16
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "사용자명(&U):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   555
         TabIndex        =   15
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "데이터베이스(&B):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   14
         Top             =   540
         Width           =   1440
      End
   End
   Begin Threed.SSPanel sspOracle 
      Height          =   1695
      Left            =   600
      TabIndex        =   2
      Top             =   570
      Width           =   5025
      _Version        =   65536
      _ExtentX        =   8864
      _ExtentY        =   2990
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtServer 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1305
         TabIndex        =   5
         Top             =   150
         Width           =   3030
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1305
         TabIndex        =   4
         Text            =   "User"
         Top             =   615
         Width           =   2205
      End
      Begin VB.TextBox txtPasswd 
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   1305
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Passwd"
         Top             =   1095
         Width           =   2205
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "사용자명(&U):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "암호(&B):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   7
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "서버(&S):"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   450
         TabIndex        =   6
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.Label lblDB 
      BackStyle       =   0  '투명
      Caption         =   "Oracle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   630
      TabIndex        =   19
      Top             =   150
      Width           =   4815
   End
End
Attribute VB_Name = "frmDBConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click()
    Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        If gDB_Parm.DBType = "1" Then
            Call WritePrivateProfileString("DATABASE", "server", txtServer, App.Path & "\didim.ini")
            Call WritePrivateProfileString("DATABASE", "uid", txtUser, App.Path & "\didim.ini")
            Call WritePrivateProfileString("DATABASE", "pwd", txtPasswd, App.Path & "\didim.ini")
        Else
            Call WritePrivateProfileString("DATABASE", "SERVER", txtMSServer, App.Path & "\didim.ini")
            Call WritePrivateProfileString("DATABASE", "DATABASE", txtDB, App.Path & "\didim.ini")
            Call WritePrivateProfileString("DATABASE", "UID", txtMSUser, App.Path & "\didim.ini")
            Call WritePrivateProfileString("DATABASE", "PWD", txtMSPasswd, App.Path & "\didim.ini")
        End If
        
        If MsgBox("DB정보를 변경하면 프로그램을 재시작 해야합니다" & vbNewLine & "프로그램을 종료하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
            Unload Me
            Exit Sub
        Else
            End
        End If
        
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Ret As Integer
    
    If gDB_Parm.DBType = "1" Then
        lblDB.Caption = "Oracle DB Set"
        sspOracle.Visible = True
        sspSQL.Visible = False
        
        txtServer = gDB_Parm.Server
        txtUser = gDB_Parm.USER
        txtPasswd = gDB_Parm.Passwd
        
    Else
        lblDB.Caption = "MS-SQL DB Set"
        sspOracle.Visible = False
        sspSQL.Visible = True
    
        txtServer = gDB_Parm.Server
        txtDB = gDB_Parm.DB
        txtUser = gDB_Parm.USER
        txtPasswd = gDB_Parm.Passwd
    
    End If
    
    
End Sub

