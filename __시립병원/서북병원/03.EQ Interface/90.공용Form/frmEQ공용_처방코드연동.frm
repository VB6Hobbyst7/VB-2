VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEQ����_ó���ڵ忬�� 
   Caption         =   "ó���ڵ忬��"
   ClientHeight    =   4605
   ClientLeft      =   2820
   ClientTop       =   570
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ����_ó���ڵ忬��.frx":0000
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
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   10
      SpreadDesigner  =   "frmEQ����_ó���ڵ忬��.frx":263A
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
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���˻��ڵ��Է�"
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
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmEQ����_ó���ڵ忬��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

