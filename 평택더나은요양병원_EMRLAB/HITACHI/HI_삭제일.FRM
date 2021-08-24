VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm삭제일설정 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "파일정리 기간 설정"
   ClientHeight    =   5325
   ClientLeft      =   2010
   ClientTop       =   2145
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5325
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "<주의!!>"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   690
      Width           =   6015
      Begin VB.TextBox txtDelDate 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "##"
         Top             =   1830
         Width           =   405
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "데이터파일들을 지울 기간을 설정하겠습니다."
         ForeColor       =   16711680
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel2 
         DataSource      =   "f"
         Height          =   405
         Left            =   1110
         TabIndex        =   3
         Top             =   1830
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   " 일  이전의 파일들을 지우도록 설정되었습니다."
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   840
      Left            =   1335
      TabIndex        =   4
      Top             =   4020
      Width           =   1890
      _Version        =   65536
      _ExtentX        =   3334
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "저    장"
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
      RoundedCorners  =   0   'False
      Picture         =   "Hi_삭제일.frx":0000
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   840
      Left            =   3210
      TabIndex        =   5
      Top             =   4020
      Width           =   1890
      _Version        =   65536
      _ExtentX        =   3334
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "취    소"
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
      RoundedCorners  =   0   'False
      Picture         =   "Hi_삭제일.frx":10A2
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Hi_삭제일.frx":1F00
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frm삭제일설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    Dim FileName    As String
    
    If txtDelDate.Text = "" Then
        MsgBox "보관할 기간을 입력하세요!!"
        txtDelDate.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtDelDate.Text) Then
        MsgBox "날자 입력이 잘못되었습니다."
        txtDelDate.SetFocus
        Exit Sub
    Else
        Screen.MousePointer = 11
        dDate = Trim$(txtDelDate.Text)
    End If
    
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
    '----- Data 저장일자 입력
    TbComm!SaveDate = Format$(dDate, "00")
            
    TbComm.Update
    
    '--- DataBase Close
    TbComm.Close
    DbComm.Close
    '==================
    
    cmdExit.Caption = "닫    기"
    
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call cmdExit.DoClick
    End Select
    
End Sub

Private Sub Form_Load()
    
    Me.Top = (Hi_Main.Height - Hi_Main.pnlBack.Height - Me.Height) / 3
    Me.Left = (Hi_Main.Width - Me.Width) / 2
        
    txtDelDate = Format$(dDate, "00")
    
End Sub


Private Sub txtDelDate_GotFocus()

    With txtDelDate
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    
End Sub

Private Sub txtDelDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Call cmdOk.DoClick
    End If
    
End Sub

Private Sub txtDelDate_LostFocus()

    txtDelDate = Format$(txtDelDate.Text, "00")
    
End Sub

