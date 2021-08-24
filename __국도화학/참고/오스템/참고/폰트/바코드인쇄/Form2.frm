VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form2"
   ScaleHeight     =   1605
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      DrawStyle       =   2  '점
      Height          =   585
      Left            =   120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   480
      Width           =   3225
   End
   Begin VB.Shape Shape1 
      Height          =   1545
      Left            =   30
      Top             =   30
      Width           =   3435
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Barcode Print"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Label2.Caption = Form1.lbBarcod.Caption
    
    Picture1.Height = Form1.Barcode1.Height
    Picture1.Width = Form1.Barcode1.Width
    Picture1.Refresh
    
End Sub



