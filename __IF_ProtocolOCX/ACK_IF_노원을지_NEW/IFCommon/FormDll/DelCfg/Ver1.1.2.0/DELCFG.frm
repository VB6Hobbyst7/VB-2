VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDelCfg 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "로컬 데이터 삭제기간 설정"
   ClientHeight    =   4920
   ClientLeft      =   2010
   ClientTop       =   2145
   ClientWidth     =   6195
   Icon            =   "DELCFG.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4920
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Height          =   4845
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   6135
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
         Height          =   360
         Left            =   630
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "###"
         Top             =   2190
         Width           =   585
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   735
         Left            =   420
         TabIndex        =   2
         Top             =   1110
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "로컬 데이터를 삭제할 기간을 설정하겠습니다."
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
         Left            =   1230
         TabIndex        =   3
         Top             =   2190
         Width           =   4305
         _Version        =   65536
         _ExtentX        =   7594
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "일  이전의 로컬 데이터를 지우도록 설정되었습니다."
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
      Begin Threed.SSCommand cmdEnsure 
         Height          =   990
         Left            =   1992
         TabIndex        =   4
         Top             =   3300
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   1746
         _StockProps     =   78
         Caption         =   "확 인"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "DELCFG.frx":27A2
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   990
         Left            =   3092
         TabIndex        =   5
         Top             =   3300
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   1746
         _StockProps     =   78
         Caption         =   "취 소"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "DELCFG.frx":3844
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "DELCFG.frx":46A2
         Top             =   300
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmDelCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEnsure_Click()
    Dim sTmp$
    Dim bRetVal As Boolean
    
    sTmp = Trim$(txtDelDate.Text)
    
    If Len(sTmp) = 0 Then
        MsgBox "로컬 데이터를 삭제할 기간을 입력하세요!!"
        txtDelDate.SetFocus
        Exit Sub
    End If
        
    If Val(sTmp) > 0 Then
        '<----- Delete Interval 등록 ------------------>
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Delete.Interval", txtDelDate)
    Else
        MsgBox "입력된 값이 0보다 작습니다!!"
        txtDelDate.SetFocus
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim sBuf$
    
    '<----- Delete Interval 구하기 ------------------->
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Delete.Interval")
        
    txtDelDate = sBuf
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegEditCurFrmTitle "DelCfg", ""
    ViewMsg ""
End Sub

Private Sub txtDelDate_GotFocus()
    Call Txt_Highlight(txtDelDate)
End Sub

Private Sub txtDelDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13     'ENTER
            KeyAscii = 0
        Case 8      'BS
        
        Case Else
            TxtTypeOnlyNumeric txtDelDate, KeyAscii
    End Select
End Sub
