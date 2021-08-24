VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form INTclear20 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "파일정리 기간 설정"
   ClientHeight    =   4860
   ClientLeft      =   2010
   ClientTop       =   2145
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4860
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "<주의!!>"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
      Begin VB.TextBox txdeldate 
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
         Left            =   660
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "##"
         Top             =   2280
         Width           =   495
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "데이터파일들을 지울 기간을 설정하겠습니다."
         ForeColor       =   12582912
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         BevelInner      =   1
      End
      Begin Threed.SSPanel SSPanel2 
         DataSource      =   "f"
         Height          =   375
         Left            =   1140
         TabIndex        =   5
         Top             =   2280
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "일  이전의 파일들을 지우도록 설정되었습니다."
         BackColor       =   12632256
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
   Begin Threed.SSCommand cmdensure 
      Height          =   870
      Left            =   4755
      TabIndex        =   2
      Top             =   165
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "저   장"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE20.frx":0000
   End
   Begin Threed.SSCommand cmdcancel 
      Height          =   870
      Left            =   5580
      TabIndex        =   1
      Top             =   165
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "종   료"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE20.frx":19A2
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   480
      Picture         =   "INFACE20.frx":3344
      Stretch         =   -1  'True
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "INTclear20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()

    Unload Me
    FrmFlag = 0
End Sub


Private Sub cmdensure_Click()
    
    Dim Tmp$

    Tmp = Trim$(txdeldate.Text)
    Screen.MousePointer = 11
    
    If Len(Tmp) = 0 Then
        MsgBox "보관할 기간을 입력하세요!!"
        txdeldate.SetFocus
        Exit Sub
    End If
    If IsNumeric(Tmp) = True Then
        If Val(Tmp) > 0 Then
            Call WritePrivateProfileString("slip setting", ByVal machstr & "delete_date", ByVal Tmp, "SLIP.INI")
            ddate = Val(Tmp)
            Unload Me
            Screen.MousePointer = 0
        Else
            MsgBox "입력된 값이 0보다 작습니다!!"
            txdeldate.SetFocus
        End If
    Else
        MsgBox "숫자형을 입력하세요!!"
        txdeldate.SetFocus
    End If
    Unload Me
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()

    Me.Top = (INTmain00.Height - INTmain00.pnlMain.Height - Me.Height) / 3
    Me.Left = (INTmain00.Width - Me.Width) / 2
    
    txdeldate.Text = Format$(ddate, "00")
    
    FrmFlag = 20
End Sub


Private Sub txdeldate_GotFocus()

    Call txbox_highlight(txdeldate)
    
End Sub


Private Sub txdeldate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdensure.SetFocus
    End If
    
End Sub


