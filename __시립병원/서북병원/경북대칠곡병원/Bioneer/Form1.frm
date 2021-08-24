VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton cmdUnvisible 
      Caption         =   "´Ý±â"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   750
      TabIndex        =   18
      Top             =   5940
      Width           =   1755
   End
   Begin FPSpread.vaSpread vasRes1 
      Height          =   6015
      Left            =   3420
      TabIndex        =   16
      Top             =   450
      Width           =   3255
      _Version        =   196608
      _ExtentX        =   5741
      _ExtentY        =   10610
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   20
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "Form1.frx":0000
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   7
      Text            =   "123456789012"
      Top             =   4485
      Width           =   1725
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   6
      Text            =   "123456789012"
      Top             =   3990
      Width           =   1725
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Text            =   "123456789012"
      Top             =   3510
      Width           =   1725
   End
   Begin VB.TextBox txtReceNo 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   4
      Text            =   "123456789012"
      Top             =   3030
      Width           =   1725
   End
   Begin VB.TextBox txtPos 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   3
      Text            =   "123456789012"
      Top             =   1905
      Width           =   1725
   End
   Begin VB.TextBox txtRack 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   2
      Text            =   "123456789012"
      Top             =   1425
      Width           =   1725
   End
   Begin VB.TextBox txtSeq 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   1
      Text            =   "123456789012"
      Top             =   930
      Width           =   1725
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   0
      Text            =   "123456789012"
      Top             =   450
      Width           =   1725
   End
   Begin FPSpread.vaSpread vasRes2 
      Height          =   6015
      Left            =   6840
      TabIndex        =   17
      Top             =   450
      Width           =   3255
      _Version        =   196608
      _ExtentX        =   5741
      _ExtentY        =   10610
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   20
      RowHeaderDisplay=   0
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "Form1.frx":0590
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "³ªÀÌ"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   4530
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "¼ºº°"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4035
      Width           =   450
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "È¯ÀÚÀÌ¸§"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   3555
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Á¢¼ö¹øÈ£"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   3075
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Pos"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1950
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Rack"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1470
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Seq #"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   975
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "°ËÃ¼¹øÈ£"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   495
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
