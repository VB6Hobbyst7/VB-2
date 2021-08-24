VERSION 5.00
Begin VB.Form frmEQ공용_Info 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Hi Interface EQ 정보"
   ClientHeight    =   4095
   ClientLeft      =   4860
   ClientTop       =   2820
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdQuit 
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lbl설명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Interface For Medical Machine"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   900
      Width           =   4065
   End
   Begin VB.Label lbl등록상표 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Hi Interface EQ"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  '단색
      Height          =   1275
      Index           =   0
      Left            =   120
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   6240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      FillColor       =   &H00808080&
      FillStyle       =   0  '단색
      Height          =   1275
      Index           =   1
      Left            =   165
      Shape           =   4  '둥근 사각형
      Top             =   120
      Width           =   6240
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Version 1.0"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   3780
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   $"frmEQ공용_Info.frx":263A
      Height          =   555
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "이 제품은 다음 사용자에게 사용이 허가되었습니다."
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   1980
      Width           =   4320
   End
   Begin VB.Label lbl회사이름 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "메디메이트(주)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   2340
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Copyright ⓒ 2010 Medimate Co., Ltd."
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   1620
      Width           =   3780
   End
End
Attribute VB_Name = "frmEQ공용_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdQuit_Click
    End Select
End Sub

Private Sub Form_Load()
    Me.Height = 4575
    Me.Width = 6660
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Me.Caption = App.LegalTrademarks & " 정보"
    lbl등록상표 = App.ProductName
    lbl설명 = "Interface For " & App.FileDescription
    lbl회사이름 = App.CompanyName
    
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
