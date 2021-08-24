VERSION 5.00
Begin VB.Form frmMicOption 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "항생제 종류"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1140
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   2565
      Begin VB.OptionButton optMicFg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "MIC 감수성"
         ForeColor       =   &H00B14563&
         Height          =   285
         Index           =   1
         Left            =   510
         TabIndex        =   1
         Tag             =   "C"
         Top             =   285
         Width           =   1380
      End
      Begin VB.OptionButton optMicFg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "일반 감수성"
         ForeColor       =   &H00B14563&
         Height          =   270
         Index           =   0
         Left            =   510
         TabIndex        =   2
         Tag             =   "S"
         Top             =   630
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmMicOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event MicSELECT(ByVal strMicFg As String)

Private Sub Form_Activate()
'    If  = "04" Then
'        optMicFg(0).Value = True
'        optMicFg(0).SetFocus
'    End If
    
End Sub

Private Sub optMicFg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        Unload Me
        Set frmMicOption = Nothing
        RaiseEvent MicSELECT(optMicFg(Index).Tag)
    End If

End Sub

Private Sub optMicFg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        Unload Me
        Set frmMicOption = Nothing
        RaiseEvent MicSELECT(optMicFg(Index).Tag)
    End If
    
End Sub
