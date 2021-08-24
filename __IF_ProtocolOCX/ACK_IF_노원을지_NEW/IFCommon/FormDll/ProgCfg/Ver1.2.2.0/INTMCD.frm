VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmIntMcd 
   BorderStyle     =   1  '단일 고정
   Caption         =   "   장비코드 설정"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2610
   Icon            =   "INTMCD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2610
   Begin VB.TextBox txtINITMode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtAPMode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1470
      Width           =   255
   End
   Begin VB.TextBox txtTXMode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1140
      Width           =   255
   End
   Begin VB.TextBox txtIFMode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   6
      Top             =   810
      Width           =   255
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   495
      Left            =   390
      TabIndex        =   2
      Top             =   2430
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "OK !!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMCd 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   390
      TabIndex        =   1
      Top             =   390
      Width           =   1875
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   390
      TabIndex        =   0
      Top             =   60
      Width           =   1875
      _Version        =   65536
      _ExtentX        =   3307
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "장비 코드"
      ForeColor       =   8454143
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   495
      Left            =   1380
      TabIndex        =   3
      Top             =   2430
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   315
      Left            =   390
      TabIndex        =   4
      Top             =   810
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "InterfaceMode"
      ForeColor       =   8454143
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   315
      Left            =   390
      TabIndex        =   5
      Top             =   1140
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "TransmitMode"
      ForeColor       =   8454143
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   315
      Left            =   390
      TabIndex        =   9
      Top             =   1470
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "Auto.P.Mode"
      ForeColor       =   8454143
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   315
      Left            =   390
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "Initial.Mode"
      ForeColor       =   8454143
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmIntMcd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisplayMachineCd()
    Call GetInterfaceCd
    
    txtMCd = gsMachineCd
    txtIFMode = gsIFMode
    txtTXMode = gsTXMode
    txtAPMode = gsAPMode
    txtINITMode = gsINITMode
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim retval As Long
    Dim bRetVal As Boolean
    
    retval = WritePrivateProfileString("InterfaceMachineCode", "InterfaceMachineCd", "" & txtMCd & "", App.Path & "\장비코드.ini")
    
    If retval = 1 Then
        gsMachineCd = txtMCd
    End If

'Interface Mode
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Interface.Mode", txtIFMode)

    If bRetVal = True Then
        gsIFMode = txtIFMode
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If

'Transmit Mode
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Transmit.Mode", txtTXMode)

    If bRetVal = True Then
        gsTXMode = txtTXMode
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
'AutoPrint Mode
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Auto.P.Mode", txtAPMode)

    If bRetVal = True Then
        gsAPMode = txtAPMode
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
'Initialize Mode
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Initialize.Mode", txtINITMode)

    If bRetVal = True Then
        gsAPMode = txtINITMode
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call DisplayMachineCd
End Sub
