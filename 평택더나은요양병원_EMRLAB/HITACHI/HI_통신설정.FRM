VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm통신설정 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "통신구성 설정 및 수정"
   ClientHeight    =   5040
   ClientLeft      =   630
   ClientTop       =   1440
   ClientWidth     =   7515
   ClipControls    =   0   'False
   Icon            =   "Hi_통신설정.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   7515
   Begin Threed.SSFrame SSFrame1 
      Height          =   1545
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   2725
      _StockProps     =   14
      Caption         =   "통신 포트"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin Threed.SSOption optPort 
         Height          =   345
         Index           =   0
         Left            =   320
         TabIndex        =   0
         Top             =   360
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "COM 1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optPort 
         Height          =   345
         Index           =   1
         Left            =   320
         TabIndex        =   1
         Top             =   750
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "COM 2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   840
      Left            =   5280
      TabIndex        =   14
      Top             =   3840
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "취     소"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   2
      RoundedCorners  =   0   'False
      Picture         =   "Hi_통신설정.frx":0442
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   840
      Left            =   5280
      TabIndex        =   13
      Top             =   2940
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "저     장"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   2
      RoundedCorners  =   0   'False
      Picture         =   "Hi_통신설정.frx":12A0
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1545
      Left            =   2760
      TabIndex        =   16
      Top             =   600
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   2725
      _StockProps     =   14
      Caption         =   "Data Bit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin Threed.SSOption optDataBit 
         Height          =   345
         Index           =   0
         Left            =   320
         TabIndex        =   2
         Top             =   360
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "7 Bit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optDataBit 
         Height          =   345
         Index           =   1
         Left            =   320
         TabIndex        =   3
         Top             =   750
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "8 Bit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1545
      Left            =   5130
      TabIndex        =   17
      Top             =   600
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   2725
      _StockProps     =   14
      Caption         =   "Stop Bit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin Threed.SSOption optStopBit 
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   360
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "1 Bit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optStopBit 
         Height          =   345
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   750
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "2 Bit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1965
      Left            =   330
      TabIndex        =   18
      Top             =   2760
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   3466
      _StockProps     =   14
      Caption         =   "Baud Rate"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin Threed.SSOption optBaudRate 
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   350
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optBaudRate 
         Height          =   345
         Index           =   1
         Left            =   300
         TabIndex        =   7
         Top             =   710
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "2400"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optBaudRate 
         Height          =   345
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   1080
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "4800"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optBaudRate 
         Height          =   345
         Index           =   3
         Left            =   300
         TabIndex        =   9
         Top             =   1460
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "9600"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   1965
      Left            =   2760
      TabIndex        =   19
      Top             =   2760
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   3466
      _StockProps     =   14
      Caption         =   "Parity"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin Threed.SSOption optParity 
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   360
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "None"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optParity 
         Height          =   345
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   750
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Odd"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optParity 
         Height          =   345
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   1140
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Even"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   2790
      Picture         =   "Hi_통신설정.frx":2342
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2790
      Picture         =   "Hi_통신설정.frx":2C0C
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5190
      Picture         =   "Hi_통신설정.frx":3AD6
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "Hi_통신설정.frx":3F18
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   450
      Picture         =   "Hi_통신설정.frx":47E2
      Top             =   2280
      Width           =   480
   End
End
Attribute VB_Name = "frm통신설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    Dim FileName    As String
    Dim i   As Integer
    
    MousePointer = 11
    
    '----- DataBase Open
    FileName = strLAB_PATH + "System\Config.mdb"
    
    Set DbComm = OpenDatabase(FileName)
    Set TbComm = DbComm.OpenRecordset("Config")
    '===================
   
    ' 자료의 존재여부 검사
    If TbComm.EOF And TbComm.BOF Then
        TbComm.AddNew
    Else
        TbComm.Edit
    End If
    
    ' 포트 설정
    If optPort(0) = True Then
        TbComm!ComPort = Right(optPort(0).Caption, 1)
        CommCfg.ComPort = Right(optPort(0).Caption, 1)
    Else
        TbComm!ComPort = Right(optPort(1).Caption, 1)
        CommCfg.ComPort = Right(optPort(1).Caption, 1)
    End If
    ' 전송속도 설정
    For i = 0 To 3
        With optBaudRate(i)
            If .Value = True Then
                TbComm!BaudRate = .Caption
                CommCfg.BaudRate = .Caption
                Exit For
            End If
        End With
    Next i
    ' 데이타 비트 설정
    If optDataBit(0) = True Then
        TbComm!DataBit = Left(optDataBit(0).Caption, 1)
        CommCfg.DataBit = Left(optDataBit(0).Caption, 1)
    Else
        TbComm!DataBit = Left(optDataBit(1).Caption, 1)
        CommCfg.DataBit = Left(optDataBit(1).Caption, 1)
    End If
    ' 스톱비트 설정
    If optStopBit(0) = True Then
        TbComm!StopBit = Left(optStopBit(0).Caption, 1)
        CommCfg.StopBit = Left(optStopBit(0).Caption, 1)
    Else
        TbComm!StopBit = Left(optStopBit(1).Caption, 1)
        CommCfg.StopBit = Left(optStopBit(1).Caption, 1)
    End If
    ' 패리티 비트 설정
    For i = 0 To 2
        With optParity(i)
            If .Value = True Then
                TbComm!Parity = Left$(.Caption, 1)
                CommCfg.Parity = Left$(.Caption, 1)
                Exit For
            End If
        End With
    Next i
    
    TbComm.Update
    '--- DataBase Close
    TbComm.Close
    DbComm.Close
    '==================
    cmdExit.Caption = "닫     기"
    
    MousePointer = 0

End Sub

Private Sub Form_Load()
            
    Me.Top = (Hi_Main.Height - Hi_Main.pnlBack.Height - Me.Height) / 3
    Me.Left = (Hi_Main.Width - Me.Width) / 2
    
    Call Display_CommCfg    '현재 통신구성 상태표시

End Sub



Private Sub Display_CommCfg()
        
    With CommCfg
        ' 포트
        optPort(Val(.ComPort) - 1) = True
        ' 전송속도
        Select Case .BaudRate
            Case 1200
                optBaudRate(0) = True
            Case 2400
                optBaudRate(1) = True
            Case 4800
                optBaudRate(2) = True
            Case 9600
                optBaudRate(3) = True
        End Select
        ' 데이타 비트
        optDataBit(Val(.DataBit) - 7) = True
        ' 스톱비트
        optStopBit(Val(.StopBit) - 1) = True
        ' 패리티 비트
        Select Case .Parity
            Case "N"
                optParity(0) = True
            Case "O"
                optParity(1) = True
            Case "E"
                optParity(2) = True
        End Select
    End With
    
End Sub
