VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaster 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DBE6E6&
   Caption         =   "Master °ü¸®"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   14745
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "´Ý±â"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   30
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8640
      Width           =   3540
   End
   Begin VB.PictureBox picForm 
      Height          =   9045
      Left            =   3660
      ScaleHeight     =   8985
      ScaleWidth      =   10935
      TabIndex        =   1
      Top             =   30
      Width           =   11000
   End
   Begin MSComctlLib.ImageList imlTreeImage 
      Left            =   705
      Top             =   8490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":0000
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":0100
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":0200
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":0300
            Key             =   "Load"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   8625
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   15214
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTreeImage"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objMasterForm As clsMasterForm

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Set objMasterForm = New clsMasterForm
    Call objMasterForm.MasterTreeviewLoad(tvwMenu)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objMasterForm = Nothing
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub tvwMenu_Collapse(ByVal Node As MSComctlLib.Node)
   Node.Image = "Close"
End Sub

Private Sub tvwMenu_Expand(ByVal Node As MSComctlLib.Node)
   Node.Image = "Open"
End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    Call objMasterForm.MasterTreeviewNodeClick(Node.Key, picForm)
End Sub

