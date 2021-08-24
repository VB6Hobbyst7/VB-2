VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_DiagName_Input 
   BorderStyle     =   0  '없음
   Caption         =   "진단명 등록"
   ClientHeight    =   2985
   ClientLeft      =   1725
   ClientTop       =   2130
   ClientWidth     =   6945
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtSummary 
      Height          =   2100
      Left            =   24
      TabIndex        =   4
      Top             =   792
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   3704
      _Version        =   393217
      BackColor       =   15463915
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MaxLength       =   1000
      TextRTF         =   $"anato203.frx":0000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808000&
      Height          =   708
      Left            =   4176
      ScaleHeight     =   645
      ScaleWidth      =   2670
      TabIndex        =   0
      Top             =   48
      Width           =   2736
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Left            =   1368
         TabIndex        =   3
         Top             =   15
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "&Close              "
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "anato203.frx":0265
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   615
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "&Ok                   "
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "anato203.frx":057F
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  '투명
      Caption         =   "진 단 명 등 록"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   336
      TabIndex        =   2
      Top             =   96
      Width           =   1800
   End
End
Attribute VB_Name = "Anato_DiagName_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    txtSummary.Enabled = True
    Unload Me

End Sub


Private Sub cmdSave_Click()
    GsDiagNo = Trim(txtSummary.Text)
    Unload Me

End Sub

