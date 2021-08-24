VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm화일명지정 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "결과 화일 지정"
   ClientHeight    =   1140
   ClientLeft      =   3405
   ClientTop       =   3615
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1140
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel3 
      Height          =   1155
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7845
      _Version        =   65536
      _ExtentX        =   13838
      _ExtentY        =   2037
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   2430
         TabIndex        =   0
         Top             =   375
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   661
         _Version        =   327680
         HideSelection   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "MMDD"
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   660
         TabIndex        =   4
         Top             =   330
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "월일입력 : HI_"
         ForeColor       =   12582912
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
         BevelWidth      =   2
         RoundedCorners  =   0   'False
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   645
         Left            =   4170
         TabIndex        =   1
         Top             =   240
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "확      인"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         RoundedCorners  =   0   'False
         Picture         =   "Hi_화일명.frx":0000
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   645
         Left            =   5820
         TabIndex        =   2
         Top             =   240
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1138
         _StockProps     =   78
         Caption         =   "취      소"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         RoundedCorners  =   0   'False
         Picture         =   "Hi_화일명.frx":10A2
      End
   End
End
Attribute VB_Name = "frm화일명지정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

    Unload Me
    
End Sub


Private Sub cmdOK_Click()

    Dim tmpname As String
    
    Screen.MousePointer = 11
        
    tmpname = "HI_" + mskDate.Text
    
    tlFileName = mskDate.Text   '일자 저장
        
    Call Create_HIDB(tmpname)     '당일날자의 DB생성
        
    Unload Me
    
    Screen.MousePointer = 0

    Load frmHi_결과받기
    frmHi_결과받기.Show

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
   
    mskDate = Format$(Now, "MMDD")
       
End Sub


Private Sub mskDate_GotFocus()

    With mskDate
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    
End Sub

Private Sub mskDate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdOK.DoClick
    End If
    
End Sub


