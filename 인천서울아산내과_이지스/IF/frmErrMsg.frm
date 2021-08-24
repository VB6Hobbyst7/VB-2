VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmErrMsg 
   BackColor       =   &H00FFFFFF&
   Caption         =   "오류내용"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   Icon            =   "frmErrMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7050
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer tmrClose 
      Left            =   2220
      Top             =   2520
   End
   Begin VB.TextBox txtErr 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   90
      Width           =   6735
   End
   Begin HSCotrol.CButton cmdSave 
      Height          =   495
      Left            =   4020
      TabIndex        =   2
      Top             =   2490
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " 저    장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmErrMsg.frx":030A
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.CButton cmdClose 
      Height          =   495
      Left            =   5490
      TabIndex        =   0
      Top             =   2490
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " 닫    기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmErrMsg.frx":0464
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
End
Attribute VB_Name = "frmErrMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Call SetErrData("ErrMsg", txtErr.Text)
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    tmrClose.Interval = 10000
    tmrClose.Enabled = True

End Sub

Private Sub tmrClose_Timer()
    
    tmrClose.Enabled = False
    Unload Me
    
End Sub
