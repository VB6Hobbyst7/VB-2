VERSION 5.00
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   435
      Left            =   5490
      TabIndex        =   1
      Top             =   2550
      Width           =   1335
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
      Height          =   2355
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   90
      Width           =   6735
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

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

