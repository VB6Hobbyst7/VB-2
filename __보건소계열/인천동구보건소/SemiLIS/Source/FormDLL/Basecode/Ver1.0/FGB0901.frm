VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0901 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "프로그램 - 환경설정"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "FGB0901.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSFrame SSFrame3 
      Height          =   915
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   5835
      _Version        =   65536
      _ExtentX        =   10292
      _ExtentY        =   1614
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "Semi-LIS Program Config "
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   915
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   1614
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   735
         Left            =   60
         Picture         =   "FGB0901.frx":030A
         Stretch         =   -1  'True
         Top             =   130
         Width           =   1320
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7275
      _Version        =   65536
      _ExtentX        =   12832
      _ExtentY        =   2461
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtProTitle2 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Laboratory Information System"
         Top             =   270
         Width           =   2685
      End
      Begin VB.TextBox txtLogTitle 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   8
         Top             =   780
         Width           =   2655
      End
      Begin VB.TextBox txtProTitle1 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   7
         Top             =   270
         Width           =   1665
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   975
         Left            =   6390
         TabIndex        =   4
         Top             =   270
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1720
         _StockProps     =   78
         Caption         =   "확인"
         BevelWidth      =   3
      End
      Begin VB.Label Label3 
         Caption         =   "로그 인 타이틀"
         Height          =   345
         Left            =   270
         TabIndex        =   6
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "프로그램 타이틀"
         Height          =   375
         Left            =   270
         TabIndex        =   5
         Top             =   330
         Width           =   1755
      End
   End
End
Attribute VB_Name = "FGB0901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Dim bRetVal As Boolean
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Program Config\App.Title", "", txtProTitle1 & " " & txtProTitle2)
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Program Config\LogIn.Title", "", txtLogTitle)
                
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sBuf$
    Dim Pos%
    
    sBuf = fCurAppTitle
    txtLogTitle = fCurLogInTitle
    
    Pos = InStr(1, sBuf, "Laboratory")
    
    If Pos = 0 Then
    Else
        txtProTitle1 = Trim$(Left$(sBuf, Pos - 1))
    End If
    
    txtProTitle2 = "Laboratory Information System"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
End Sub

