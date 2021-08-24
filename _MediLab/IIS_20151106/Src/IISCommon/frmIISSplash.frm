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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIISSplash.frx":0FEA
   ScaleHeight     =   4635
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Label lblCopyright 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "Copyright 2015 MediLAB  Co., Ltd."
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
      Left            =   3585
      TabIndex        =   4
      Top             =   3615
      Width           =   2775
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
      Caption         =   "Windows"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   2550
      Width           =   1260
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
      Caption         =   "Interface System을 로딩하고 있읍니다......"
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
'   작성자  : 오세원
'   내  용  : Splash Form
'   작성일  : 2005-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISSplash = Nothing
End Sub

