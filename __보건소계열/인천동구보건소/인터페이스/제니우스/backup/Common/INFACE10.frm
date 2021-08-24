VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form INTcomm10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Åë½Å±¸¼º ¼³Á¤ ¹× ¼öÁ¤"
   ClientHeight    =   5340
   ClientLeft      =   630
   ClientTop       =   1440
   ClientWidth     =   7545
   ClipControls    =   0   'False
   Icon            =   "INFACE10.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5340
   ScaleWidth      =   7545
   Begin VB.Frame Frame1 
      Caption         =   "Æ÷Æ®"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   2145
      Begin VB.OptionButton OptCom 
         Caption         =   "COM1"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   300
         Width           =   1000
      End
      Begin VB.OptionButton OptCom 
         Caption         =   "COM2"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   270
         TabIndex        =   9
         Top             =   630
         Width           =   1000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Bit"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   2145
      Begin VB.OptionButton OptData 
         Caption         =   "7bit"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   390
         Width           =   1000
      End
      Begin VB.OptionButton OptData 
         Caption         =   "8bit"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1000
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Stop Bit"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   5160
      TabIndex        =   2
      Top             =   2940
      Width           =   2145
      Begin VB.OptionButton OptStop 
         Caption         =   "1bit"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1000
      End
      Begin VB.OptionButton OptStop 
         Caption         =   "2bit"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   810
         Width           =   1000
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Baud Rate"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   360
      TabIndex        =   1
      Top             =   2940
      Width           =   2145
      Begin VB.OptionButton OptBaud 
         Caption         =   "9600"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   3
         Left            =   270
         TabIndex        =   14
         Top             =   1440
         Width           =   1000
      End
      Begin VB.OptionButton OptBaud 
         Caption         =   "4800"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   270
         TabIndex        =   13
         Top             =   1110
         Width           =   1000
      End
      Begin VB.OptionButton OptBaud 
         Caption         =   "2400"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   270
         TabIndex        =   12
         Top             =   630
         Width           =   1000
      End
      Begin VB.OptionButton OptBaud 
         Caption         =   "1200"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   11
         Top             =   330
         Width           =   1000
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Parity"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   2760
      TabIndex        =   0
      Top             =   2940
      Width           =   2145
      Begin VB.OptionButton OptParity 
         Caption         =   "Even"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   270
         TabIndex        =   17
         Top             =   1140
         Width           =   1000
      End
      Begin VB.OptionButton OptParity 
         Caption         =   "Odd"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   270
         TabIndex        =   16
         Top             =   660
         Width           =   1000
      End
      Begin VB.OptionButton OptParity 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   15
         Top             =   330
         Width           =   1000
      End
   End
   Begin Threed.SSCommand CmdCancle 
      Height          =   870
      Left            =   6510
      TabIndex        =   19
      Top             =   240
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Á¾   ·á"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE10.frx":0442
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   870
      Left            =   5685
      TabIndex        =   18
      Top             =   240
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Àú   Àå"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE10.frx":1DE4
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   2760
      Picture         =   "INFACE10.frx":3786
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2760
      Picture         =   "INFACE10.frx":3BC8
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5100
      Picture         =   "INFACE10.frx":400A
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "INFACE10.frx":444C
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   480
      Picture         =   "INFACE10.frx":488E
      Top             =   2400
      Width           =   480
   End
End
Attribute VB_Name = "INTcomm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim commset As commset
Private Sub CmdCancle_Click()

    Unload Me
    FrmFlag = 0
End Sub

Private Sub cmdSave_Click()

    Screen.MousePointer = 11
    
    With tbcomm
        .Edit
        !Port = commset.Port
        !baud_rate = commset.baud_rate
        !data_bit = commset.data_bit
        !stop_bit = commset.stop_bit
        !parity = commset.parity
        .Update
    End With

    'Screen.MousePointer = 0
    'MsgBox "ÀúÀåÀÌ µÇ¾ú½À´Ï´Ù.  È®ÀÎÀ» ´©¸£½Å ÈÄ ´Ù¸¥ °÷À¸·Î ÀÌµ¿ÇÏ·Á¸é ´Ý±â¸¦ ´©¸£¼¼¿ä!!"
    Unload Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
            
    'formÀ» °¡¿îµ¥¿¡ À§Ä¡
    Me.Top = (INTmain00.Height - INTmain00.pnlMain.Height - Me.Height) / 3
    Me.Left = (INTmain00.Width - Me.Width) / 2
    
    Set dbcomm = OpenDatabase(FileName & commstr)
    Set tbcomm = dbcomm.OpenRecordset("cfgcomm")
    Call display_commcfg
    
    FrmFlag = 10
    
End Sub



Private Sub display_commcfg()

    OptCom(Val(tbcomm!Port) - 1).Value = True
    OptBaud(((Log(Val(tbcomm!baud_rate) / 600)) / Log(2#)) - 1).Value = True
    OptData(Val(tbcomm!data_bit) Mod 7).Value = True
    OptStop(Val(tbcomm!stop_bit) - 1).Value = True
    Select Case UCase(tbcomm!parity)
    Case "N"
        OptParity(0).Value = True
    Case "O"
        OptParity(1).Value = True
    Case "E"
        OptParity(2).Value = True
    End Select
    
    With commset
        .Port = tbcomm!Port
        .baud_rate = tbcomm!baud_rate
        .data_bit = tbcomm!data_bit
        .stop_bit = tbcomm!stop_bit
        .parity = tbcomm!parity
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    tbcomm.Close
    dbcomm.Close
    
End Sub


Private Sub OptBaud_Click(Index As Integer)

    commset.baud_rate = 2 ^ (Index + 1) * 600
    
End Sub

Private Sub OptCom_Click(Index As Integer)
    
    commset.Port = Index + 1
    
End Sub


Private Sub OptData_Click(Index As Integer)

    commset.data_bit = Index + 7
    
End Sub


Private Sub OptParity_Click(Index As Integer)

    Select Case Index
    Case 0
        commset.parity = "N"
    Case 1
        commset.parity = "O"
    Case 2
        commset.parity = "E"
    End Select
    
End Sub

Private Sub OptStop_Click(Index As Integer)

    commset.stop_bit = Index + 1
    
End Sub


