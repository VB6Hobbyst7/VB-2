VERSION 5.00
Begin VB.Form frmODBCErrorMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Error "
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   60
      Top             =   60
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "»Æ¿Œ"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   510
      TabIndex        =   0
      Top             =   690
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmODBCErrMsg.frx":0000
      Top             =   150
      Width           =   495
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODBC Error Message"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   780
      TabIndex        =   1
      Top             =   180
      Width           =   1830
   End
End
Attribute VB_Name = "frmODBCErrorMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdConfirm_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    With Timer1
        .Interval = 5000
        .Enabled = True
    End With
    
End Sub

Private Sub Timer1_Timer()

    Unload Me
    
End Sub


