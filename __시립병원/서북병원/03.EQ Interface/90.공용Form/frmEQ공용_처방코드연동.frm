VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEQ공용_처방코드연동 
   Caption         =   "처방코드연동"
   ClientHeight    =   4605
   ClientLeft      =   2820
   ClientTop       =   570
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_처방코드연동.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   11745
   Begin FPSpread.vaSpread sprVIEW 
      Height          =   3555
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   11235
      _Version        =   393216
      _ExtentX        =   19817
      _ExtentY        =   6271
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      ColsFrozen      =   2
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   10
      SpreadDesigner  =   "frmEQ공용_처방코드연동.frx":263A
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbl장비명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비검사코드입력"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   2880
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmEQ공용_처방코드연동"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

