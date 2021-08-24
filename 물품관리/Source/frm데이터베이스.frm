VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frm데이터베이스 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Database Information"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSFrame SSFrame2 
      Height          =   525
      Left            =   90
      TabIndex        =   12
      Top             =   120
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   926
      _Version        =   262144
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   1710
         TabIndex        =   13
         Text            =   "SQL Server"
         Top             =   120
         Width           =   2775
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   300
         Left            =   60
         TabIndex        =   14
         Top             =   120
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사용자명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1755
      Left            =   90
      TabIndex        =   0
      Top             =   750
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   3096
      _Version        =   262144
      BackColor       =   16777215
      Begin VB.TextBox txtSqlServer 
         Height          =   300
         Left            =   1710
         TabIndex        =   4
         Text            =   "SQL Server"
         Top             =   390
         Width           =   2775
      End
      Begin VB.TextBox txtSqlDb 
         Height          =   300
         Left            =   1710
         TabIndex        =   3
         Text            =   "GuestStandbyDB"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtSqlUser 
         Height          =   300
         Left            =   1710
         TabIndex        =   2
         Text            =   "sa"
         Top             =   1050
         Width           =   2775
      End
      Begin VB.TextBox txtSqlPswd 
         Height          =   300
         Left            =   1710
         TabIndex        =   1
         Text            =   "password"
         Top             =   1380
         Width           =   2775
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   300
         Left            =   60
         TabIndex        =   5
         Top             =   390
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Server Name"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   300
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   529
         _Version        =   262144
         Font3D          =   5
         BackColor       =   -2147483629
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "▒ DATABASE(ms-sql) ▒"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   300
         Left            =   60
         TabIndex        =   7
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Database Name"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   300
         Left            =   60
         TabIndex        =   8
         Top             =   1050
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "User Id"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   300
         Left            =   60
         TabIndex        =   9
         Top             =   1380
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Password"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   390
      Left            =   1320
      TabIndex        =   10
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "저장(&S)"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdClose 
      Height          =   390
      Left            =   2430
      TabIndex        =   11
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "닫기(&X)"
      ButtonStyle     =   2
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1425
      Left            =   90
      TabIndex        =   15
      Top             =   2580
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   2514
      _Version        =   262144
      BackColor       =   16777215
      Begin VB.TextBox txtOraPswd 
         Height          =   300
         Left            =   1710
         TabIndex        =   18
         Text            =   "hospital"
         Top             =   1050
         Width           =   2775
      End
      Begin VB.TextBox txtOraUser 
         Height          =   300
         Left            =   1710
         TabIndex        =   17
         Text            =   "twmed"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtOraTns 
         Height          =   300
         Left            =   1710
         TabIndex        =   16
         Text            =   "KAHP"
         Top             =   390
         Width           =   2775
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   300
         Left            =   60
         TabIndex        =   19
         Top             =   60
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   529
         _Version        =   262144
         Font3D          =   5
         BackColor       =   -2147483629
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "▒ 검진 DATABASE(oracle) ▒"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   300
         Left            =   60
         TabIndex        =   20
         Top             =   390
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TNS Name"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   300
         Left            =   60
         TabIndex        =   21
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "User Id"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   300
         Left            =   60
         TabIndex        =   22
         Top             =   1050
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Password"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frm데이터베이스"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cReg As clsRegister

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()

    With cReg
        .sqlserver = Trim(txtSqlServer.Text)
        .sqldb = Trim(txtSqlDb.Text)
        .sqluser = Trim(txtSqlUser.Text)
        .sqlpswd = Trim(txtSqlPswd.Text)
        
        .username = Trim(txtUserName.Text)
        
        .oratns = Trim(txtOraTns.Text)
        .orauser = Trim(txtOraUser.Text)
        .orapswd = Trim(txtOraPswd.Text)
        
        Call .csRegisterSave
    End With
    
    MsgBox "데이터베이스 연결정보가 저장되었습니다.!", vbInformation

End Sub

Private Sub Form_Load()

    Set cReg = New clsRegister
    
    txtSqlServer.Text = cReg.sqlserver
    txtSqlDb.Text = cReg.sqldb
    txtSqlUser.Text = cReg.sqluser
    txtSqlPswd.Text = cReg.sqlpswd

    txtUserName.Text = cReg.username

    txtOraTns.Text = cReg.oratns
    txtOraUser.Text = cReg.orauser
    txtOraPswd.Text = cReg.orapswd

End Sub

