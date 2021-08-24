VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frmCodeList 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "List"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox lstCodeList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6810
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   4650
   End
   Begin FPSpread.vaSpread tblItemList 
      Height          =   6870
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   4650
      _Version        =   196608
      _ExtentX        =   8202
      _ExtentY        =   12118
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmCodeList.frx":0000
      TextTip         =   4
   End
End
Attribute VB_Name = "frmCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    
    If Me.BorderStyle = 1 Then
        lstCodeList.Height = Me.Height - 350
        Me.Height = lstCodeList.Height + 450
        lstCodeList.Width = Me.Width - 130
        'Me.ScaleHeight = 100
        tblItemList.Height = Me.Height - 370
        tblItemList.Width = Me.Width - 100
    Else
        'Me.ScaleHeight = 100
        lstCodeList.Height = Me.Height - 30
        Me.Height = lstCodeList.Height + 80
        lstCodeList.Width = Me.Width - 80
        'Me.ScaleHeight = 100
        tblItemList.Height = Me.Height - 50
        tblItemList.Width = Me.Width - 100
    End If
    
End Sub
