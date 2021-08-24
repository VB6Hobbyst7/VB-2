VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmPatReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   8190
   Begin HSCotrol.HSLabel HSLabel1 
      Height          =   330
      Left            =   90
      Top             =   150
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LABORATORIES Report"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel6 
      Height          =   330
      Left            =   210
      Top             =   660
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "White Blood Cell"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel5 
      Height          =   330
      Left            =   2190
      Top             =   660
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6500 (/mm)"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel7 
      Height          =   330
      Left            =   5460
      Top             =   660
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4,500~11,000"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel2 
      Height          =   330
      Left            =   210
      Top             =   1020
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Hemoglobin"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel3 
      Height          =   330
      Left            =   2190
      Top             =   1020
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "125.3 (g/dl)"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel8 
      Height          =   330
      Left            =   3330
      Top             =   1020
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "12.3~15.3"
      BevelOut        =   0
   End
   Begin HSCotrol.HSLabel HSLabel4 
      Height          =   330
      Left            =   3360
      Top             =   660
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Baskerville Old Face"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4,500~11,000"
      BevelOut        =   0
   End
   Begin HSCotrol.CButton cmdClose 
      Height          =   375
      Left            =   6390
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   16777215
      Caption         =   "¥›±‚"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤√º"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   0
      HoverColor      =   4210752
      HoverPicture    =   "frmPatReport.frx":0000
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7575
      Left            =   30
      Top             =   30
      Width           =   7845
   End
End
Attribute VB_Name = "frmPatReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
