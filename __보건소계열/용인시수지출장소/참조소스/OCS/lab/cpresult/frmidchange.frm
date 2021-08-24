VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmIDChange 
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 Change Screen"
   ClientHeight    =   1800
   ClientLeft      =   2895
   ClientTop       =   2610
   ClientWidth     =   4695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4695
   Begin Threed.SSPanel SSPanel2 
      Height          =   510
      Left            =   90
      TabIndex        =   7
      Top             =   45
      Width           =   4470
      _Version        =   65536
      _ExtentX        =   7885
      _ExtentY        =   900
      _StockProps     =   15
      Caption         =   " 현재Logon User"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox txtCurrent 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   90
         Width           =   2850
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1140
      Left            =   90
      TabIndex        =   3
      Top             =   585
      Width           =   4470
      _Version        =   65536
      _ExtentX        =   7885
      _ExtentY        =   2011
      _StockProps     =   15
      Caption         =   "Change User"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin Threed.SSCommand cmdCancel 
         Height          =   735
         Left            =   3645
         TabIndex        =   6
         Top             =   270
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "취소"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmIDChange.frx":0000
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   735
         Left            =   2925
         TabIndex        =   2
         Top             =   270
         Width           =   690
         _Version        =   65536
         _ExtentX        =   1217
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmIDChange.frx":08DA
      End
      Begin VB.TextBox txtPass 
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   1530
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   1320
      End
      Begin VB.TextBox txtUser 
         Height          =   330
         Left            =   1530
         MaxLength       =   6
         TabIndex        =   0
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Password :"
         Height          =   240
         Left            =   315
         TabIndex        =   5
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "UserID :"
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   315
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmIDChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_click()
    Unload Me
    
End Sub

Private Sub CmdOK_Click()
    Dim strSql      As String
    
    If Not IsNumeric(txtUser.Text) Then
        MsgBox "UserID 는 숫자로만 입력이 가능합니다!..", vbCritical
        Exit Sub
    End If
    
    
    strSql = ""
'O  strSql = strSql & " SELECT Name, PassWord, Class, SubClass, Grade, Part, SubPart, DeptCode, Rank "
    strSql = strSql & " SELECT Name, PassWord, Class, Grade, Part,Buse"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PASS "
    strSql = strSql & " WHERE (ProgramID IS NULL Or  ProgramID = ' ')"
    strSql = strSql & " AND    IDnumber  = '" & Trim(txtUser.Text) & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "UserID 와 PassWord 를 Check 하세요!.."
        Exit Sub
    End If
    'GstrPassRank = adoSet.Fields("Rank").Value & ""
    'GstrSubClass = adoSet.Fields("SubCLass").Value & ""
    'GstrSubPart = adoSet.Fields("SubPart").Value & ""
    GstrPassWord = adoSet.Fields("PassWord").Value & ""
    GstrPassName = adoSet.Fields("Name").Value & ""
    GstrPassClass = adoSet.Fields("Class").Value & ""
    GstrPassGrade = adoSet.Fields("Grade").Value & ""
    GstrPassPart = adoSet.Fields("Part").Value & ""
    GstrPassDept = adoSet.Fields("Buse").Value & ""
    GstrIdnumber = Format(txtUser.Text, "000000")
    GstrPassIDnumber = Format(txtUser.Text, "000000")
    
    Call adoSetClose(adoSet)
    
    MsgBox "사용자가 " & txtCurrent.Text & " 에서 " & GstrPassName & " 으로 변경되었습니다.!"
    mdiMain.stbMain.Panels(2).Text = GstrPassName
    
    Dim X       As Form
    
    On Error Resume Next
    
    For Each X In Forms
        If X.MDIChild = True Then
            If X.Name <> "mdiMain" Then
                Unload X
            End If
        End If
    Next
    
    Unload Me
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub Form_Load()
    
    txtCurrent.Text = GstrPassName
    
End Sub
