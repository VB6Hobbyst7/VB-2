VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   Caption         =   "WorkList"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpread.vaSpread vasList 
      Height          =   6045
      Left            =   30
      TabIndex        =   1
      Top             =   720
      Width           =   9105
      _Version        =   393216
      _ExtentX        =   16060
      _ExtentY        =   10663
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmWorkList.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   9105
      Begin IF_Cobas4800국립암센터.MDButton btnSear 
         Height          =   405
         Left            =   5430
         TabIndex        =   4
         Top             =   180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpStrDate 
         Height          =   345
         Left            =   1710
         TabIndex        =   2
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   40190
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   345
         Left            =   3570
         TabIndex        =   3
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   40190
      End
      Begin IF_Cobas4800국립암센터.MDButton btnSave 
         Height          =   405
         Left            =   6600
         TabIndex        =   5
         Top             =   180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "저장"
      End
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSear_Click()
    Dim sRes As String
    
    ClearSpread vasList
    
    
End Sub

Private Sub Form_Load()
    dtpStrDate.Value = frmInterface.Text_Today.Text
    dtpEndDate.Value = frmInterface.Text_Today.Text
    
    ClearSpread vasList
End Sub




