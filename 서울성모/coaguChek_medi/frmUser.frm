VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmUser 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows 기본값
   Begin Threed.SSPanel spUser 
      Height          =   6075
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   10716
      _StockProps     =   15
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdExit1 
         Caption         =   "종료"
         Height          =   345
         Left            =   3930
         TabIndex        =   18
         Top             =   540
         Width           =   705
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "확인"
         Height          =   345
         Left            =   3120
         TabIndex        =   15
         Top             =   540
         Width           =   705
      End
      Begin VB.TextBox txtLID 
         Height          =   300
         Left            =   870
         TabIndex        =   13
         Top             =   240
         Width           =   1965
      End
      Begin VB.TextBox txtLPW 
         Height          =   300
         IMEMode         =   3  '사용 못함
         Left            =   870
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label lblGrade 
         BackStyle       =   0  '투명
         Height          =   225
         Left            =   3090
         TabIndex        =   17
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "ID"
         Height          =   315
         Left            =   420
         TabIndex        =   16
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "PW"
         Height          =   315
         Left            =   420
         TabIndex        =   12
         Top             =   660
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[사용자관리]"
      Height          =   2325
      Left            =   4620
      TabIndex        =   1
      Top             =   330
      Width           =   2805
      Begin VB.CommandButton cmdExit 
         Caption         =   "종료"
         Height          =   405
         Left            =   1830
         TabIndex        =   10
         Top             =   1710
         Width           =   825
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "삭제"
         Height          =   405
         Left            =   990
         TabIndex        =   9
         Top             =   1710
         Width           =   825
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         Height          =   405
         Left            =   150
         TabIndex        =   8
         Top             =   1710
         Width           =   825
      End
      Begin VB.ComboBox cmbGrade 
         Height          =   300
         Left            =   930
         TabIndex        =   7
         Top             =   990
         Width           =   1695
      End
      Begin VB.TextBox txtPW 
         Height          =   300
         IMEMode         =   3  '사용 못함
         Left            =   690
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   630
         Width           =   1965
      End
      Begin VB.TextBox txtUser 
         Height          =   300
         Left            =   690
         TabIndex        =   2
         Top             =   270
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "Grade"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1050
         Width           =   585
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "PW"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "ID"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Width           =   405
      End
   End
   Begin FPSpread.vaSpread vasUser 
      Height          =   5925
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4395
      _Version        =   393216
      _ExtentX        =   7752
      _ExtentY        =   10451
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   50
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmUser.frx":0000
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDel_Click()
    SQL = "delete from user_data where userid = '" & Trim(txtUser.Text) & "' and userpw = '" & Trim(txtPW.Text) & "'"
    res = SendQuery(gLocal, SQL)
    
    uClear
    uDisplay
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub cmdExit1_Click()
    Unload Me
    
End Sub

Private Sub cmdLogin_Click()
    SQL = "select grade from user_data where userid = '" & Trim(txtLID.Text) & "' and userpw = '" & Trim(txtLPW.Text) & "'"
    res = db_select_Col(gLocal, SQL)
    
    If Trim(gReadBuf(0)) = "0" Then
        spUser.Visible = False
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdSave_Click()
    SQL = "delete from user_data where userid = '" & Trim(txtUser.Text) & "' and userpw = '" & Trim(txtPW.Text) & "'"
    res = SendQuery(gLocal, SQL)
    
    SQL = "insert into user_data(userid, userpw, grade) " & vbCrLf & _
          "values('" & Trim(txtUser.Text) & "', '" & Trim(txtPW.Text) & "', " & cmbGrade.ListIndex & ") "
    res = SendQuery(gLocal, SQL)
    
    uClear
    uDisplay
End Sub

Private Sub Form_Load()
    cmbGrade.AddItem "Supervisor", 0
    cmbGrade.AddItem "User", 1
    cmbGrade.AddItem "Viewer", 2
    
    uClear
    uDisplay
End Sub

Private Sub uDisplay()
    ClearSpread vasUser
    
    SQL = "select userid, userpw, grade from user_data order by grade, userid "
    res = db_select_Vas(gLocal, SQL, vasUser)
    
End Sub

Private Sub uClear()
    
    txtUser = ""
    txtPW = ""
    cmbGrade.ListIndex = 0
    
End Sub

Private Sub vasUser_Click(ByVal Col As Long, ByVal Row As Long)
    txtUser = Trim(GetText(vasUser, Row, 1))
    txtPW = Trim(GetText(vasUser, Row, 2))
    If IsNumeric(Trim(GetText(vasUser, Row, 3))) = True Then
        cmbGrade.ListIndex = Trim(GetText(vasUser, Row, 3))
    Else
        cmbGrade.ListIndex = 0
    End If
    
End Sub
