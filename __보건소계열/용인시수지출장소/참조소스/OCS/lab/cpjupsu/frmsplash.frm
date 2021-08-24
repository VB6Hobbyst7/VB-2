VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   4245
   ClientLeft      =   2250
   ClientTop       =   2100
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   3690
         Left            =   135
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label lblCopyright 
         Caption         =   "저작권"
         Height          =   210
         Left            =   3975
         TabIndex        =   3
         Top             =   3105
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "회사: Win&&Win Information System"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   3390
         Width           =   3045
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "버전"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   6360
         TabIndex        =   4
         Top             =   2700
         Width           =   504
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "플랫폼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5880
         TabIndex        =   5
         Top             =   2340
         Width           =   972
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "검체접수실"
         BeginProperty Font 
            Name            =   "바탕체"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4185
         TabIndex        =   7
         Top             =   1305
         Width           =   1740
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "이 제품은 다음 사용자에게 사용이 허가되었습니다."
         Height          =   255
         Left            =   2700
         TabIndex        =   1
         Top             =   225
         Width           =   4155
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "(의)한림 병원"
         BeginProperty Font 
            Name            =   "바탕체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2655
         TabIndex        =   6
         Top             =   765
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
