VERSION 5.00
Begin VB.Form frmInputBox 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "대화 상자 캡션"
   ClientHeight    =   1905
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtInputData 
      Height          =   300
      IMEMode         =   3  '사용 못함
      Left            =   135
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1470
      Width           =   6015
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00DBE6E6&
      Caption         =   "취소"
      Height          =   510
      Left            =   4830
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   795
      Width           =   1320
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00DBE6E6&
      Caption         =   "확인"
      Height          =   510
      Left            =   4830
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   270
      Width           =   1320
   End
   Begin VB.Label lblPrompt 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Label1"
      Height          =   1200
      Left            =   165
      TabIndex        =   3
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkClick(ByVal pInputData As String)

Private mvarFormCaption As String
Private mvarPrompt As String

Public Property Let FormCaption(ByVal vData As String)
    mvarFormCaption = vData
End Property

Public Property Get FormCaption() As String
    FormCaption = mvarFormCaption
End Property

Public Property Let Prompt(ByVal vData As String)
    mvarPrompt = vData
End Property

Public Property Get Prompt() As String
    Prompt = mvarPrompt
End Property

Private Sub CancelButton_Click()
    RaiseEvent OkClick("")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtInputData.SetFocus
End Sub

Private Sub Form_Load()
    frmInputBox.Caption = mvarFormCaption
    lblPrompt.Caption = mvarPrompt
End Sub

Private Sub OKButton_Click()
    RaiseEvent OkClick(Trim(txtInputData.Text))
End Sub

Private Sub txtInputData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtInputData.Text) = "" Then
            MsgBox mvarPrompt, vbExclamation
            Exit Sub
        End If
    
        Call OKButton_Click
    End If
End Sub
