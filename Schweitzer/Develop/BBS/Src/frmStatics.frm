VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatics 
   BackColor       =   &H00DBE6E6&
   Caption         =   "통계 및 출력"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14730
   Icon            =   "frmStatics.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14730
   WindowState     =   2  '최대화
   Begin VB.PictureBox picForm 
      Height          =   9045
      Left            =   3660
      ScaleHeight     =   8985
      ScaleWidth      =   10935
      TabIndex        =   2
      Top             =   30
      Width           =   11000
   End
   Begin MSComctlLib.ImageList imlTreeImage 
      Left            =   225
      Top             =   8595
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
            Picture         =   "frmStatics.frx":076A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatics.frx":086A
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatics.frx":096A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatics.frx":0A6A
            Key             =   "Load"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8700
      Width           =   3555
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   8685
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   15319
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTreeImage"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStatics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private objStaticsForm As clsStaticsForm

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Set objStaticsForm = New clsStaticsForm
    Call objStaticsForm.StaticsTreeviewLoad(tvwMenu)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objStaticsForm = Nothing
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
    Call objStaticsForm.StaticsTreeviewNodeClick(Node.Key, picForm)
End Sub

