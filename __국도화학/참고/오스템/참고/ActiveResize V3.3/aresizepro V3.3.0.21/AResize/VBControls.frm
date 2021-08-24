VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form VBControls 
   Caption         =   "VB Controls Resize Demo"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   8640
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   18
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   7680
      FormWidthDT     =   9315
      FormScaleHeightDT=   7275
      FormScaleWidthDT=   9195
      HideControlsOnResize=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Form State"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   30
      TabIndex        =   34
      Top             =   6810
      Width           =   2985
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore Form State"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3090
      TabIndex        =   33
      Top             =   6810
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   435
      Left            =   6180
      TabIndex        =   32
      Top             =   6810
      Width           =   2985
   End
   Begin VB.Frame Frame3 
      Caption         =   "Various Controls"
      Height          =   1545
      Left            =   30
      TabIndex        =   16
      Top             =   5220
      Width           =   9135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   7500
         TabIndex        =   28
         Text            =   "$200"
         Top             =   1170
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   7500
         TabIndex        =   26
         Text            =   "$400"
         Top             =   870
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   7500
         TabIndex        =   24
         Text            =   "$1500"
         Top             =   570
         Width           =   1485
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "VBControls.frx":0000
         Left            =   3060
         List            =   "VBControls.frx":0028
         MultiSelect     =   1  'Simple
         TabIndex        =   0
         Top             =   570
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "I have enjoyed my visits so much"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1140
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         Caption         =   "I have visited many of them"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   870
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         Caption         =   "The Greek islands are wonderful"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   570
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Greece has 350 islands"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "For food:"
         Height          =   195
         Index           =   4
         Left            =   6030
         TabIndex        =   27
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "For transportation:"
         Height          =   195
         Index           =   3
         Left            =   6030
         TabIndex        =   25
         Top             =   900
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "For accomodation:"
         Height          =   195
         Index           =   2
         Left            =   6030
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "I have spent these amounts:"
         Height          =   195
         Index           =   1
         Left            =   6030
         TabIndex        =   22
         Top             =   270
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "I have visited the following islands:"
         Height          =   195
         Index           =   0
         Left            =   3030
         TabIndex        =   21
         Top             =   270
         Width           =   2520
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buttons"
      Height          =   2025
      Left            =   30
      TabIndex        =   3
      Top             =   3120
      Width           =   9135
      Begin VB.CommandButton Command1 
         Caption         =   "PATMOS"
         Height          =   465
         Index           =   14
         Left            =   6090
         TabIndex        =   31
         Top             =   1470
         Width           =   2985
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ZAKYNTHOS"
         Height          =   465
         Index           =   13
         Left            =   3090
         TabIndex        =   30
         Top             =   1470
         Width           =   2985
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CHIOS"
         Height          =   465
         Index           =   12
         Left            =   90
         TabIndex        =   29
         Top             =   1470
         Width           =   2985
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TINOS"
         Height          =   465
         Index           =   11
         Left            =   7590
         TabIndex        =   15
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "KITHIRA"
         Height          =   465
         Index           =   10
         Left            =   6090
         TabIndex        =   14
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "MILOS"
         Height          =   465
         Index           =   9
         Left            =   4590
         TabIndex        =   13
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CORFU"
         Height          =   465
         Index           =   8
         Left            =   3090
         TabIndex        =   12
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CRETA"
         Height          =   465
         Index           =   7
         Left            =   1590
         TabIndex        =   11
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "POROS"
         Height          =   465
         Index           =   6
         Left            =   90
         TabIndex        =   10
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SIFNOS"
         Height          =   765
         Index           =   5
         Left            =   7590
         TabIndex        =   9
         Top             =   210
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LEFKADA"
         Height          =   765
         Index           =   4
         Left            =   6090
         TabIndex        =   8
         Top             =   210
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "KOS"
         Height          =   765
         Index           =   3
         Left            =   4590
         TabIndex        =   7
         Top             =   210
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SANTORINI"
         Height          =   765
         Index           =   2
         Left            =   3090
         TabIndex        =   6
         Top             =   210
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RHODES"
         Height          =   765
         Index           =   1
         Left            =   1590
         TabIndex        =   5
         Top             =   210
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "MYKONOS"
         Height          =   765
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PictureBox Controls"
      Height          =   2505
      Left            =   30
      TabIndex        =   2
      Top             =   540
      Width           =   9135
      Begin VB.PictureBox Picture3 
         Height          =   2145
         Left            =   6000
         Picture         =   "VBControls.frx":008A
         ScaleHeight     =   2085
         ScaleWidth      =   3015
         TabIndex        =   37
         Top             =   300
         Width           =   3075
      End
      Begin VB.PictureBox Picture2 
         Height          =   2145
         Left            =   3030
         Picture         =   "VBControls.frx":438C
         ScaleHeight     =   2085
         ScaleWidth      =   2895
         TabIndex        =   36
         Top             =   300
         Width           =   2955
      End
      Begin VB.PictureBox Picture1 
         Height          =   2145
         Left            =   60
         Picture         =   "VBControls.frx":6A0C
         ScaleHeight     =   2085
         ScaleWidth      =   2895
         TabIndex        =   35
         Top             =   300
         Width           =   2955
      End
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
      Left            =   2310
      TabIndex        =   1
      Top             =   120
      Width           =   4560
   End
End
Attribute VB_Name = "VBControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'-------------------------------------------------------------------------
'                   Not even a single line of code!!!
'-------------------------------------------------------------------------

'This is just to show how you can save the form state (form size and
'position and size / position of all its controls) and restore the form
'to this exact state at any later time
Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case Is = 0
            ActiveResize1.SaveForm Me, True, App.ProductName
        Case Is = 1
            ActiveResize1.RestoreForm Me, True, App.ProductName
    End Select
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

