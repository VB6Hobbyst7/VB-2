VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "CFX4032.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMQC 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Query Microorganism Quality Control Result"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14385
   Tag             =   "25800"
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSelDT1 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   360
      Left            =   3960
      Style           =   1  '그래픽
      TabIndex        =   22
      Top             =   225
      Width           =   330
   End
   Begin VB.CommandButton cmdSelDT2 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      Height          =   360
      Left            =   6465
      Style           =   1  '그래픽
      TabIndex        =   21
      Top             =   225
      Width           =   330
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Text            =   "Feb 11 1999"
      Top             =   255
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4605
      TabIndex        =   2
      Text            =   "Feb 21 1999"
      Top             =   255
      Width           =   1830
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12660
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8595
      Width           =   1245
   End
   Begin TabDlg.SSTab sstMQC 
      Height          =   7575
      Left            =   450
      TabIndex        =   0
      Tag             =   "25815"
      Top             =   840
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14411494
      TabCaption(0)   =   "QC Result for Media"
      TabPicture(0)   =   "Lis258.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMediaName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMediaList"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ssMedia"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "List2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "QC Result for biochemical test"
      TabPicture(1)   =   "Lis258.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMicrobe"
      Tab(1).Control(1)=   "lblSpecies1"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "ssBio"
      Tab(1).Control(4)=   "List1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "QC Result for susceptibility test"
      TabPicture(2)   =   "Lis258.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSpeciesName"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "lblSpeciesList"
      Tab(2).Control(3)=   "ssSusc"
      Tab(2).Control(4)=   "List3"
      Tab(2).Control(5)=   "ChartFX1"
      Tab(2).ControlCount=   6
      Begin ChartfxLibCtl.ChartFX ChartFX1 
         Height          =   2820
         Left            =   -71790
         TabIndex        =   23
         Top             =   4440
         Width           =   9795
         _cx             =   2146845565
         _cy             =   2146833262
         Build           =   7
         TypeMask        =   42467585
         View3D          =   1
         Axis(0).MinorStep=   -20
         Axis(0).TickMark=   -32767
         Axis(2).MinorStep=   -1
         RGBBk           =   14737632
         nColors         =   2
         Colors          =   "Lis258.frx":0054
         _Data_          =   "Lis258.frx":0084
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5820
         ItemData        =   "Lis258.frx":0189
         Left            =   -74640
         List            =   "Lis258.frx":019F
         TabIndex        =   18
         Top             =   1170
         Width           =   2670
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6060
         ItemData        =   "Lis258.frx":01EB
         Left            =   -74655
         List            =   "Lis258.frx":0201
         TabIndex        =   14
         Top             =   1185
         Width           =   2670
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6060
         ItemData        =   "Lis258.frx":024D
         Left            =   330
         List            =   "Lis258.frx":026C
         TabIndex        =   9
         Top             =   1095
         Width           =   2295
      End
      Begin FPSpread.vaSpread ssBio 
         Height          =   5940
         Left            =   -71835
         TabIndex        =   6
         Top             =   1155
         Width           =   9915
         _Version        =   196608
         _ExtentX        =   17489
         _ExtentY        =   10478
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         Protect         =   0   'False
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis258.frx":02B8
         VisibleCols     =   3
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread ssMedia 
         Height          =   6060
         Left            =   2760
         TabIndex        =   8
         Tag             =   "25809"
         Top             =   1110
         Width           =   10350
         _Version        =   196608
         _ExtentX        =   18256
         _ExtentY        =   10689
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         Protect         =   0   'False
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis258.frx":1E04
         VisibleCols     =   3
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread ssSusc 
         Height          =   3105
         Left            =   -71805
         TabIndex        =   13
         Top             =   1185
         Width           =   9840
         _Version        =   196608
         _ExtentX        =   17357
         _ExtentY        =   5477
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         Protect         =   0   'False
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis258.frx":3801
         VisibleCols     =   3
         VisibleRows     =   500
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         Caption         =   "S.aureus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70095
         TabIndex        =   17
         Top             =   645
         Width           =   3390
      End
      Begin VB.Label lblSpecies1 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "Standard Species"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74535
         TabIndex        =   19
         Tag             =   "25804"
         Top             =   705
         Width           =   2415
      End
      Begin VB.Label lblSpeciesList 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "Standard Species"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74490
         TabIndex        =   16
         Tag             =   "25806"
         Top             =   735
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         Caption         =   "표준 균주 표준 명칭"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69885
         TabIndex        =   15
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label lblSpeciesName 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Standard Species"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -71670
         TabIndex        =   12
         Tag             =   "25807"
         Top             =   705
         Width           =   1680
      End
      Begin VB.Label lblMediaList 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "MEDIA Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   11
         Tag             =   "25802"
         Top             =   735
         Width           =   1875
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         Caption         =   "배지 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4260
         TabIndex        =   10
         Top             =   690
         Width           =   2235
      End
      Begin VB.Label lblMediaName 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Media Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2895
         TabIndex        =   7
         Tag             =   "25803"
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblMicrobe 
         BackColor       =   &H00DBE6E6&
         Caption         =   "microorganism"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71685
         TabIndex        =   5
         Tag             =   "25805"
         Top             =   690
         Width           =   1335
      End
   End
   Begin VB.Label lblRcvDate 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Received Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   525
      TabIndex        =   20
      Tag             =   "25801"
      Top             =   285
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4410
      TabIndex        =   4
      Top             =   270
      Width           =   165
   End
End
Attribute VB_Name = "frmMQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub
