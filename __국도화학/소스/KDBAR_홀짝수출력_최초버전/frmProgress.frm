VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin MSComctlLib.ProgressBar Xprog 
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   820
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   8415
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmLogin.frm
'   작성자  : 오세원
'   내  용  : 프로그레스바 폼
'   작성일  : 2015-04-29
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Load()

    Screen.MousePointer = 11
    Xprog.Min = 1
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Screen.MousePointer = 0
    DoEvents

End Sub
