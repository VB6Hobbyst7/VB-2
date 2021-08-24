VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form VBControls2 
   Caption         =   "VB Controls Resize Demo"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   7305
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   60
         Top             =   150
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   18
         ScreenHeight    =   1024
         ScreenWidth     =   1280
         ScreenHeightDT  =   1024
         ScreenWidthDT   =   1280
         MinFormHeight   =   3678
         MinFormWidth    =   5035
         AutoCenterForm  =   -1  'True
         FormHeightDT    =   5490
         FormWidthDT     =   7515
         FormScaleHeightDT=   5085
         FormScaleWidthDT=   7395
         ResizePictureBoxContents=   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "VB Controls Resize Demo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   16
         Top             =   210
         Width           =   4410
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   60
      TabIndex        =   7
      Top             =   930
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "File Controls"
      TabPicture(0)   =   "VBControls2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Drive1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Dir1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "File1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Combo / List Controls"
      TabPicture(1)   =   "VBControls2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1(0)"
      Tab(1).Control(1)=   "List1(0)"
      Tab(1).Control(2)=   "List1(1)"
      Tab(1).Control(3)=   "Combo1(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "PictureBox Controls"
      TabPicture(2)   =   "VBControls2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).Control(1)=   "Picture1"
      Tab(2).ControlCount=   2
      Begin VB.PictureBox Picture2 
         Height          =   2805
         Left            =   -71460
         Picture         =   "VBControls2.frx":0054
         ScaleHeight     =   2745
         ScaleWidth      =   3465
         TabIndex        =   18
         Top             =   510
         Width           =   3525
      End
      Begin VB.PictureBox Picture1 
         Height          =   2805
         Left            =   -74790
         Picture         =   "VBControls2.frx":363C
         ScaleHeight     =   2745
         ScaleWidth      =   3225
         TabIndex        =   17
         Top             =   510
         Width           =   3285
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "VBControls2.frx":6297
         Left            =   -71250
         List            =   "VBControls2.frx":62AA
         TabIndex        =   15
         Text            =   "12345678901234567890"
         Top             =   600
         Width           =   3285
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Index           =   1
         ItemData        =   "VBControls2.frx":631C
         Left            =   -71250
         List            =   "VBControls2.frx":634A
         TabIndex        =   14
         Top             =   1050
         Width           =   3285
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Index           =   0
         ItemData        =   "VBControls2.frx":64D6
         Left            =   -74730
         List            =   "VBControls2.frx":6504
         TabIndex        =   13
         Top             =   1050
         Width           =   3285
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "VBControls2.frx":6690
         Left            =   -74730
         List            =   "VBControls2.frx":66A3
         TabIndex        =   12
         Text            =   "12345678901234567890"
         Top             =   600
         Width           =   3285
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   3720
         TabIndex        =   11
         Top             =   1020
         Width           =   3375
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   180
         TabIndex        =   10
         Top             =   1050
         Width           =   3465
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   6915
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   50
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   780
      Width           =   7425
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   50
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   4500
      Width           =   8625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   405
      Index           =   4
      Left            =   5940
      TabIndex        =   4
      Top             =   4620
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   405
      Index           =   3
      Left            =   4470
      TabIndex        =   3
      Top             =   4620
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   405
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   4620
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   405
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   4620
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   405
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   4620
      Width           =   1425
   End
End
Attribute VB_Name = "VBControls2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'-------------------------------------------------------------------------
'                   Not even a single line of code!!!
'-------------------------------------------------------------------------
    

Private Sub Command1_Click(Index As Integer)
    If Index = 4 Then Unload Me
End Sub
