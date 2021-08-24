VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmMicroZone 
   Caption         =   "Micro Zone / Mic 구분 Screen"
   ClientHeight    =   1560
   ClientLeft      =   3525
   ClientTop       =   3585
   ClientWidth     =   4155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4155
   Begin VB.Frame Frame1 
      Caption         =   "Zone & Mic 를 선택하세요!..."
      Height          =   1335
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   3735
      Begin VB.OptionButton optZone 
         Caption         =   "Zone"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optMic 
         Caption         =   "Mic"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   1695
      End
      Begin VB.OptionButton optCancel 
         Caption         =   "Cancel"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   795
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   1402
         _StockProps     =   78
         Caption         =   "확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmMicroZone.frx":0000
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmMicroZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

    If optZone.Value = True Then frmMicroGrmgr.tvGroup.Tag = "Zone"
    If optMic.Value = True Then frmMicroGrmgr.tvGroup.Tag = "Mic"
    If optCancel.Value = True Then frmMicroGrmgr.tvGroup.Tag = "Cancel"
    
    Unload Me
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub
