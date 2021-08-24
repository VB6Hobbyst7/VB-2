VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form RunTimeControls 
   Caption         =   "Run-Time Controls Demo"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
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
   ScaleHeight     =   5265
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   60
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   18
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   5670
      FormWidthDT     =   6255
      FormScaleHeightDT=   5265
      FormScaleWidthDT=   6135
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   2
      Top             =   4830
      Width           =   1065
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
      Height          =   2355
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   6045
      Begin VB.PictureBox Picture2 
         Height          =   2115
         Index           =   0
         Left            =   3000
         Picture         =   "RunTimeControls.frx":0000
         ScaleHeight     =   2055
         ScaleWidth      =   2925
         TabIndex        =   4
         Top             =   180
         Width           =   2985
      End
      Begin VB.PictureBox Picture1 
         Height          =   2145
         Index           =   0
         Left            =   60
         Picture         =   "RunTimeControls.frx":2680
         ScaleHeight     =   2085
         ScaleWidth      =   2865
         TabIndex        =   3
         Top             =   150
         Width           =   2925
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the form to create new controls at run-time and then resize the form..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   630
      TabIndex        =   1
      Top             =   2640
      Width           =   4875
   End
End
Attribute VB_Name = "RunTimeControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'-------------------------------------------------------------------------
'           Only one line of code to handle run-time controls!!!
'-------------------------------------------------------------------------


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    
    'Create new control array elements at run time
    Load Frame1(1)
    Load Picture1(1)
    Load Picture2(1)
    Frame1(1).Top = Frame1(0).Height
    Set Picture1(1).Container = Frame1(1)
    Set Picture2(1).Container = Frame1(1)
    Frame1(1).Visible = True
    Picture1(1).Visible = True
    Picture2(1).Visible = True
    
    'Now reset ActiveResize so it can refresh its internal memory and
    'detect and store information on the new controls
    ActiveResize1.Reset

End Sub

