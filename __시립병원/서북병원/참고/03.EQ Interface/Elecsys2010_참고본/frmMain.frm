VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Interface 장비선택"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4710
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtTemp 
      Height          =   405
      Left            =   360
      TabIndex        =   5
      Top             =   2010
      Visible         =   0   'False
      Width           =   1965
   End
   Begin Threed.SSCommand sscOK 
      Height          =   435
      Left            =   1740
      TabIndex        =   2
      Top             =   1260
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "확인"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboEquip 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1620
      TabIndex        =   1
      Top             =   600
      Width           =   2595
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "장비선택"
      ForeColor       =   12582912
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.74
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
   End
   Begin Threed.SSCommand sscCancel 
      Height          =   435
      Left            =   3030
      TabIndex        =   3
      Top             =   1260
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "취소"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1635
      Left            =   360
      TabIndex        =   6
      Top             =   2580
      Width           =   4305
      _Version        =   196613
      _ExtentX        =   7594
      _ExtentY        =   2884
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmMain.frx":0742
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "장비를 선택하세요!"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   4
      Top             =   210
      Width           =   4065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'장비
    cboEquip.AddItem "", 0
'    cboEquip.AddItem "ABL50Meta", 1
'    cboEquip.AddItem "Axsym", 2
'    cboEquip.AddItem "CD3000", 3
'    cboEquip.AddItem "K1000", 4
'    cboEquip.AddItem "TBA30FR", 5
'    cboEquip.AddItem "TBA120FR", 6
'    cboEquip.AddItem "NOVA", 7

    cboEquip.AddItem "AXSYM", 1
    cboEquip.AddItem "CELL-DYN 3000", 2
    cboEquip.AddItem "NOVA", 3
    cboEquip.AddItem "SYSMEX(K-1000)", 4
    cboEquip.AddItem "TBA-30FR", 5
    cboEquip.AddItem "TBA-120FR", 6
    cboEquip.AddItem "URISCAN-PRO+", 7
                
End Sub

Private Sub sscCancel_Click()
    Unload Me
End Sub

Private Sub sscOK_Click()
Dim sEquipName As String

    sEquipName = Trim(cboEquip.Text)
    
    Select Case sEquipName
    Case ""
        Exit Sub
    
'    Case "ABL50Meta"
'        frmABL50.Show
        
    Case "AXSYM"
        gEquip = "00003"
'        frmAxsymBatch.Show
        frmAxsymBatch_1.Show
    
    Case "CELL-DYN 3000"
        gEquip = "00005"
        frmCD3000.Show
        
    Case "SYSMEX(K-1000)"
        gEquip = "00006"
        frmK1000.Show
        
    Case "NOVA"
        gEquip = "00004"
        frmNovaProfileM.Show
        
    Case "TBA-30FR"
        gEquip = "00002"
        frmTBA30FR_2.Show
         
    Case "TBA-120FR"
        gEquip = "00001"
        frmTBA120FR.Show
        
    Case "URISCAN-PRO+"
        gEquip = "00012"
        frmUriscan.Show
    End Select
    
    Unload Me
End Sub
