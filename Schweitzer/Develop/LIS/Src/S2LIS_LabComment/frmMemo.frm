VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMemo 
   Appearance      =   0  '∆Ú∏È
   BackColor       =   &H00FFFBF7&
   BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin RichTextLib.RichTextBox txtMemo 
      Height          =   2760
      Left            =   135
      TabIndex        =   2
      Top             =   390
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   4868
      _Version        =   393217
      BackColor       =   16776183
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMemo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFE8D0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4350
      Style           =   1  '±◊∑°«»
      TabIndex        =   1
      Top             =   45
      Width           =   270
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   450
      BackColor       =   16771280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileNm As String
    
Private Sub cmdClose_Click()
    Unload Me
    Set frmMemo = Nothing
End Sub

Private Sub Form_Load()
    
    Me.Left = frmMain.Width - Me.Width - 220
    Me.Top = 0
    lblTitle.Caption = Now
    strFileNm = App.Path & "\" & objDoctor.DoctId & ".txt"
    If Dir(strFileNm) <> "" Then
        txtMemo.FileName = App.Path & "\" & objDoctor.DoctId & ".txt"
        txtMemo.LoadFile strFileNm
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtMemo.SaveFile strFileNm
End Sub
