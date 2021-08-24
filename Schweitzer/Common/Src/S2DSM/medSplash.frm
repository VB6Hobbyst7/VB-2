VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   4635
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "medSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFF0DF&
      BorderStyle     =   0  '없음
      Height          =   1530
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
      Begin VB.Image Image2 
         Height          =   960
         Left            =   75
         Picture         =   "medSplash.frx":0FEA
         Top             =   270
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   885
         Picture         =   "medSplash.frx":4234
         Top             =   450
         Width           =   1920
      End
   End
   Begin VB.Label lblProjectName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  '투명
      Caption         =   "Laboratory System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   1650
      Width           =   6885
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copyright 1999  POMIS, Ltd."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4080
      TabIndex        =   4
      Top             =   3795
      Width           =   2280
   End
   Begin VB.Label lblRegistMsg 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "This version is registered to"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3975
      TabIndex        =   3
      Top             =   3555
      Width           =   2355
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Windows 98/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3900
      TabIndex        =   2
      Top             =   2730
      Width           =   2400
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5490
      TabIndex        =   1
      Top             =   3045
      Width           =   810
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  '투명
      Caption         =   "Laboratory System을 로딩하고 있습니다......"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   4335
      Width           =   4725
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarProductName As String
Private mvarVersion As String
Private mvarCopyright As String


'Product명 - Splash화면의 Title
Public Property Let ProductName(ByVal pValue As String)
    mvarProductName = pValue
End Property

'Product의 버전정보
Public Property Let Version(ByVal pValue As String)
    mvarVersion = pValue
End Property

'Product의 저작권정보
Public Property Let Copyright(ByVal pValue As String)
    mvarCopyright = pValue
End Property



Private Sub Form_Load()
   
   lblProjectName = mvarProductName ' 프로젝트명
   lblVersion.Caption = "Version " & mvarVersion    '버전
   lblRegistMsg.Caption = "This version is registered to   " & medGetComNm
   lblCopyright.Caption = mvarCopyright
   Me.Show
   
   DoEvents
   
   Call medAlwaysOn(Me, 1)
   
End Sub

Public Sub ShowMessage(ByVal strMsg As String)

    lblMessage.Caption = strMsg
    DoEvents
    
End Sub

