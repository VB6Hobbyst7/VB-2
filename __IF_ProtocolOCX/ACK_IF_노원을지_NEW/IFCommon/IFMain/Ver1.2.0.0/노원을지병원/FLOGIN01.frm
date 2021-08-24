VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FLOGIN01 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "사용자 ID 확인"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "FLOGIN01.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin Threed.SSCommand cmdOK 
      Height          =   765
      Left            =   390
      TabIndex        =   5
      Top             =   2160
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3360
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "확 인   F2"
      ForeColor       =   128
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
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   570
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   2355
      _StockProps     =   14
      Caption         =   "사용자 확인"
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel2 
         Height          =   330
         Left            =   330
         TabIndex        =   3
         Top             =   330
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "사용자 ID"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  '사용 못함
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   750
         Width           =   1635
      End
      Begin VB.TextBox txtLogID 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   330
         Width           =   1635
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   330
         Left            =   330
         TabIndex        =   4
         Top             =   750
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "PASSWORD"
         ForeColor       =   12648447
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   765
      Left            =   2310
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "취 소   Esc"
      ForeColor       =   4210752
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
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   330
      Picture         =   "FLOGIN01.frx":000C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "FLOGIN01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'PassWord 입력Count(4회까지)
Dim miPassCnt       As Integer
Dim msUserNm        As String
Dim msPWD           As String
Dim msUserOtherInfo As String

Private Function GetUserInfo(ByVal sUID$, ByVal sPWD$) As String
    On Error GoTo ErrHandler
    
    Dim objUSER As Object
    
    Set objUSER = CreateObject("BGETUSER" & Left(FMAIN01.msVerUserInfo, 2) & ".BCGETUSER" & FMAIN01.msVerUserInfo)
    
    GetUserInfo = objUSER.GetUserInfo(sUID, sPWD)
    
    Set objUSER = Nothing
    
    Exit Function
    
ErrHandler:
End Function

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim sBuf$, sTmp$
    
    sBuf = GetUserInfo(Trim(txtLogID), Trim(txtPass))
    
    If sBuf = "" Then
        MsgBox "사용자를 확인할 수 없습니다!!"
        
        Exit Sub
    End If
    
    sTmp = GetByOne(sBuf, sBuf)
    
    If sTmp = "" Then
        FMAIN01.txtUID = Trim(txtLogID)
        FMAIN01.txtPWD = Trim(txtPass)
        FMAIN01.txtUserNm = GetByOne(sBuf, sBuf)
        FMAIN01.txtUserOther = GetByOne(sBuf, sBuf)
        Unload Me
    Else
        MsgBox sTmp
        
        If Left(sTmp, 8) = "PASSWORD" Then
            txtPass.SetFocus
        Else
            txtLogID.SetFocus
        End If
        
        '입력오류회수 증가
        miPassCnt = miPassCnt + 1
        
        If miPassCnt > 3 Then
            End
        End If
        
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            Call cmdOK.DoClick
        Case vbKeyEscape
            Call cmdExit.DoClick
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    miPassCnt = 0
    msUserNm = ""
    msPWD = ""
End Sub

Private Sub txtLogID_GotFocus()
    Call Txt_Highlight(txtLogID)
End Sub

Private Sub txtLogID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'Validation Check
        If txtLogID = "" Then
            MsgBox "사용자 ID를 입력하십시오"
            txtLogID.SetFocus
            
            Exit Sub
        End If
        
        txtPass.SetFocus
    End If
End Sub

Private Sub txtPass_GotFocus()
    Call Txt_Highlight(txtPass)
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtPass = "" Then
            MsgBox "PASSWORD를 입력하십시오."
            
            txtPass = ""
            txtPass.SetFocus
            
            Exit Sub
        End If
        
        Call cmdOK.DoClick
    End If
End Sub
