VERSION 5.00
Begin VB.Form frmBloodComment 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Cross-Matching 결과등록 Comment"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmBloodComment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D8DEDA&
      Caption         =   "확인(&O)"
      Height          =   390
      Left            =   3720
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   1560
      Width           =   930
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   1380
      Left            =   60
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4605
   End
End
Attribute VB_Name = "frmBloodComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    With frmBBS201.tblResult
        .Row = .ActiveRow
        .Col = 18
        .value = txtRemark
        If txtRemark <> "" Then
            '
        End If
    End With
    Unload Me
End Sub
