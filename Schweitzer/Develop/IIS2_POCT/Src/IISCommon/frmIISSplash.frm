VERSION 5.00
Begin VB.Form frmIISSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   4635
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmIISSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFF0DF&
      BorderStyle     =   0  '없음
      Height          =   1185
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
      Begin VB.Image Image2 
         Height          =   960
         Left            =   75
         Picture         =   "frmIISSplash.frx":0FEA
         Top             =   135
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   960
         Picture         =   "frmIISSplash.frx":4234
         Top             =   390
         Width           =   1920
      End
   End
   Begin VB.Label lblProjectNm 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  '투명
      Caption         =   "Interface System"
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
      Height          =   675
      Left            =   1035
      TabIndex        =   5
      Top             =   1470
      Width           =   4815
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "Copyright 2002  POMIS  Co., Ltd."
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
      Left            =   3705
      TabIndex        =   4
      Top             =   3615
      Width           =   2655
   End
   Begin VB.Label lblRegister 
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
      Top             =   3375
      Width           =   2355
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Windows 98/2000/XP"
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
      Left            =   3435
      TabIndex        =   2
      Top             =   2550
      Width           =   2865
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Version"
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
      Top             =   2865
      Width           =   810
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  '투명
      Caption         =   "Laboratory System을 로딩하고 있읍니다......"
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
      Left            =   195
      TabIndex        =   0
      Top             =   4335
      Width           =   4725
   End
End
Attribute VB_Name = "frmIISSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISSplash.frm
'   작성자  : 이상대
'   내  용  : Splash Form
'   작성일  : 2003-12-04
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISSplash = Nothing
End Sub

