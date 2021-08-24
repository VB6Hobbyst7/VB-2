VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmSetFTP 
   Caption         =   "TCP/IP Setting"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5460
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtFtpTime 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   4695
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1950
      Width           =   690
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   5460
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   5460
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "FTP Server Settting."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   210
         Width           =   1980
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4785
         Picture         =   "frmSetFTP.frx":0000
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   5
      Top             =   2340
      Width           =   5850
   End
   Begin VB.TextBox txtHostIP 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1095
      TabIndex        =   0
      Top             =   780
      Width           =   2490
   End
   Begin VB.TextBox txtFtpUser 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1095
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1170
      Width           =   2490
   End
   Begin VB.TextBox txtFtpPasswd 
      Appearance      =   0  '평면
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1095
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2490
   End
   Begin VB.TextBox txtFtpPort 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   4695
      MaxLength       =   5
      TabIndex        =   4
      Top             =   780
      Width           =   690
   End
   Begin VB.TextBox txtDir 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1095
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1950
      Width           =   2490
   End
   Begin BHButton.BHImageButton cmdOk 
      Height          =   375
      Left            =   2820
      TabIndex        =   14
      Top             =   2520
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Ok"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdCancel 
      Height          =   375
      Left            =   4125
      TabIndex        =   13
      Top             =   2520
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Cancle"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "대기시간 :"
      Height          =   180
      Left            =   3720
      TabIndex        =   16
      Top             =   1995
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      Caption         =   "Server IP :"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   825
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ftp Port :"
      Height          =   180
      Left            =   3720
      TabIndex        =   11
      Top             =   825
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Ftp ID :"
      Height          =   180
      Left            =   465
      TabIndex        =   10
      Top             =   1215
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Passwd :"
      Height          =   180
      Left            =   255
      TabIndex        =   9
      Top             =   1605
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Directory :"
      Height          =   180
      Left            =   195
      TabIndex        =   8
      Top             =   1995
      Width           =   870
   End
End
Attribute VB_Name = "frmSetFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
  If Trim(txtHostIP) = "" Then
      MsgBox " 사용할 IP를 입력 하시오.", vbExclamation
      Exit Sub
  ElseIf Trim(txtFtpPort) = "" Then
      MsgBox " 사용할 FTP Prot를 입력 하시오."
      Exit Sub
  ElseIf Trim(txtFtpUser) = "" Then
      MsgBox " 사용할 FTP ID를 입력 하시오.", vbExclamation
      Exit Sub
  ElseIf Trim(txtFtpPasswd) = "" Then
      MsgBox " 사용할 FTP 암호를 입력 하시오.", vbExclamation
      Exit Sub
  End If


    Call SaveString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_HOSTIP, txtHostIP)
    Call SaveString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FPORT, txtFtpPort)
    Call SaveString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FID, txtFtpUser)
    Call SaveString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FPW, txtFtpPasswd)
    Call SaveString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FDIR, txtDir)
    Call SaveString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FDIR, txtFtpTime)
    
  Call Unload(Me)

End Sub

Private Sub Form_Load()
  Dim fHostIP     As String
  Dim fPort       As String
  Dim fID         As String
  Dim fPW         As String
  Dim fDir        As String
  Dim ftime       As String

    fHostIP = GetString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_HOSTIP)
    fPort = GetString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FPORT)
    fID = GetString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FID)
    fPW = GetString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FPW)
    fDir = GetString(HKEY_CURRENT_USER, REG_SRV_FTP, KEY_FDIR)
    
  txtFtpPasswd = fPW
  txtFtpPort = fPort
  txtFtpUser = fID
  txtDir = fDir

  txtHostIP = Trim(fHostIP)
    
  If Trim(fPort) = "" Then
      txtFtpPort = 21
  End If
  
 
End Sub

