VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPassword 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Password Check Box"
   ClientHeight    =   1485
   ClientLeft      =   2580
   ClientTop       =   3420
   ClientWidth     =   5340
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   1035
      Left            =   1140
      TabIndex        =   5
      Top             =   60
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1826
      _StockProps     =   15
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
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtUser 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  '사용 못함
         Left            =   1200
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   540
         Width           =   1215
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   675
         Left            =   3300
         TabIndex        =   3
         Top             =   180
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   1191
         _StockProps     =   78
         Caption         =   "Cancel"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmPassword.frx":0000
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   675
         Left            =   2520
         TabIndex        =   2
         Top             =   180
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1191
         _StockProps     =   78
         Caption         =   "확인"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmPassword.frx":08DA
      End
      Begin VB.Label Label1 
         Caption         =   "UserID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
   End
   Begin MSComctlLib.StatusBar stbPass 
      Align           =   2  '아래 맞춤
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   1170
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6429
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin MSForms.Image Image1 
      Height          =   1035
      Left            =   60
      Top             =   60
      Width           =   1035
      BorderStyle     =   0
      SizeMode        =   3
      SpecialEffect   =   6
      Size            =   "1826;1826"
      Picture         =   "frmPassword.frx":14B4
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
    
End Sub

Private Sub cmdOk_Click()
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Password"
    strSql = strSql & " WHERE  UserID   =  '" & Trim(txtUser.Text) & "'"
    strSql = strSql & " AND    Password = '" & Trim(txtPass.Text) & "'"
    
    If adoSetOpen(strSql, adoSet) Then
        Call adoSetClose(adoSet)
        Unload Me
    Else
        MsgBox "Password 가 틀립니다!. 확인바랍니다", vbOKOnly + vbInformation, "PassCheck_Error"
        gIntPassCnt = gIntPassCnt + 1
        If gIntPassCnt > 2 Then End
        txtPass.SetFocus
        Exit Sub
    End If
    
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
    
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then
    '    Call cmdOk_Click
    'End If
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)
    
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        GoSub Variable_Clear
        strSql = ""
        strSql = strSql & " SELECT *"
        strSql = strSql & " FROM   TWEXAM_Password"
        strSql = strSql & " WHERE  UserID  =  '" & Trim(txtUser.Text) & "'"
        
        If False = adoSetOpen(strSql, adoSet) Then
            stbPass.Panels(1).Text = "☞. 해당 User 등록되지 않았습니다!.."
            cmdOk.Enabled = False
            txtUser.Text = ""
            txtUser.SetFocus
            Exit Sub
        End If
        
        stbPass.SimpleText = ""
        stbPass.Panels(1).Text = "Username = " & Trim(adoSet.Fields("Username").Value & "")
        
        gStrUserID = Trim(adoSet.Fields("UserID").Value & "")
        gStrUsername = Trim(adoSet.Fields("Username").Value & "")
        gStrPass = Trim(adoSet.Fields("Password").Value & "")
        gStrDept = Trim(adoSet.Fields("DeptCode").Value & "")
        gStrRank = Trim(adoSet.Fields("Rank").Value & "")
        gStrToisa = Trim(adoSet.Fields("ToisaGb").Value & "")
        gStrSlip = "11"
        cmdOk.Enabled = True
        Call adoSetClose(adoSet)
    End If
    Exit Sub
    

Variable_Clear:
    gStrUserID = ""
    gStrPass = ""
    gStrDept = ""
    gStrRank = ""
    gStrToisa = ""
    gStrSlip = ""

    Return
    
End Sub

Private Sub txtUser_LostFocus()
    
    If Trim(txtUser.Text) <> "" Then
        Call txtUser_KeyDown(vbKeyReturn, 0)
    End If
    
End Sub
