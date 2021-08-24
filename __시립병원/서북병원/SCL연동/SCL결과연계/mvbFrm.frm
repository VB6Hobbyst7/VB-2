VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{88F75480-0574-11D0-8085-0000C0BD354B}#1.0#0"; "VISM.OCX"
Begin VB.Form mvbFrm 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5085
   WindowState     =   1  '최소화
   Begin VISMLib.VisM Mvb1 
      Left            =   2100
      Top             =   1650
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      PLIST           =   ""
      pDelim          =   $"mvbFrm.frx":0000
      Interval        =   1000
      P0              =   ""
      P1              =   ""
      P2              =   ""
      P3              =   ""
      P4              =   ""
      P5              =   ""
      P6              =   ""
      P7              =   ""
      P8              =   ""
      P9              =   ""
      VALUE           =   ""
      Code            =   ""
      NameSpace       =   ""
      TimeOut         =   0
      ExecFlag        =   0
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   1470
      Top             =   450
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   555
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   660
      Top             =   330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "mvbFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

