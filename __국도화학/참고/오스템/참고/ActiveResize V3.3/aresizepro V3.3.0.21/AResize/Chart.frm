VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Chart 
   Caption         =   "Chart Resize Demo"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
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
   ScaleHeight     =   4980
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   150
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   18
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   5385
      FormWidthDT     =   7995
      FormScaleHeightDT=   4980
      FormScaleWidthDT=   7875
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   6750
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3915
      Index           =   1
      Left            =   3990
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3495
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "Chart.frx":0000
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3915
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3495
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "Chart.frx":2516
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Note: Flickering that might appear in the charts is not caused by ActiveResize!"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   90
      TabIndex        =   6
      Top             =   4650
      Width           =   6435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sample Chart Demo"
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
      Left            =   2130
      TabIndex        =   2
      Top             =   150
      Width           =   3660
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'-------------------------------------------------------------------------
'                   Not even a single line of code!!!
'-------------------------------------------------------------------------

Private Sub cmdExit_Click()
    Unload Me
End Sub
