VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FSS0401 
   BorderStyle     =   3  'Å©±â °íÁ¤ ´ëÈ­ »óÀÚ
   ClientHeight    =   3915
   ClientLeft      =   3045
   ClientTop       =   1050
   ClientWidth     =   4170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3915
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlbottom 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   3510
      Width           =   4170
      _Version        =   65536
      _ExtentX        =   7355
      _ExtentY        =   714
      _StockProps     =   15
      ForeColor       =   16576
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Font3D          =   3
      Alignment       =   1
      Begin Threed.SSPanel pnlMsg 
         Height          =   330
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   8388608
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
      End
   End
   Begin Threed.SSPanel pnlmain 
      Align           =   1  'À§ ¸ÂÃã
      Height          =   3525
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4170
      _Version        =   65536
      _ExtentX        =   7355
      _ExtentY        =   6218
      _StockProps     =   15
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      Begin FPSpread.vaSpread SpdCode 
         Height          =   3000
         Left            =   90
         OleObjectBlob   =   "FSS0401.frx":0000
         TabIndex        =   5
         Top             =   90
         Width           =   3990
      End
      Begin VB.TextBox txtCd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   690
         TabIndex        =   0
         Top             =   3150
         Width           =   2235
      End
      Begin Threed.SSCommand CmdClk 
         Height          =   285
         Left            =   2970
         TabIndex        =   3
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "View"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand CmdEsc 
         Cancel          =   -1  'True
         Height          =   285
         Left            =   3510
         TabIndex        =   2
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Esc"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "µ¸¿òÃ¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "ÄÚµå¸í"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   3195
         Width           =   540
      End
   End
End
Attribute VB_Name = "FSS0401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChangeKeyFlag As Integer
Dim vCd
Dim vCdNm

Private Sub DisplayInit()
    txtCd = ""
    
    'SpreadBackColor Option
    iSpdBackColorOption = 3
    
    With SpdCode
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = SpdBackcolor(iSpdBackColorOption)   'GBR
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Col2 = 2
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        .MaxRows = 0
        .MaxRows = 10
    End With
End Sub
Private Sub CmdClk_Click()

    MousePointer = 11

    If ChangeKeyFlag = True Then
        
        Call SetWindowText(hWndCd, CStr(vCd))
        
        ChangeKeyFlag = False
        
        SpdCode.Col = 1
        SpdCode.Col2 = SpdCode.MaxCols
        SpdCode.Row = 1
        SpdCode.Row2 = SpdCode.MaxRows
        SpdCode.BlockMode = True
        SpdCode.Action = SS_ACTION_CLEAR_TEXT
        SpdCode.BlockMode = False
    End If
    
    MousePointer = 0
    
    Unload Me
End Sub

Private Sub CmdEsc_Click()
  
    Unload Me
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        CmdEsc_Click
    End If

End Sub

Private Sub Form_Load()

    Dim ret%
    Dim i%
    Dim j%
    
    Call DisplayInit
    
    Me.KeyPreview = True
    
'    For i = 1 To giCodeHlpCnt
        For j = 1 To giCodeHlpCnt
'            If Format$(i, "00000") = gCodeHlpTable(j).sSeq Then
                If j > 10 Then
                    SpdCode.MaxRows = j
                End If
                
                Call SpdCode.SetText(1, j, gCodeHlpTable(j).sCodeNm & "")
                Call SpdCode.SetText(2, j, gCodeHlpTable(j).sCode & "")
 '           End If
        Next
 '   Next
End Sub

Private Sub SpdCode_Click(ByVal Col As Long, ByVal Row As Long)
    
    
    If Row = 0 Then
        Exit Sub
    End If
    
    Call spdReverse(SpdCode, -1, -1, Row, Row, RGB(255, 230, 230), iSpdBackColorOption)
    
    Call SpdCode.GetText(1, Row, vCdNm)
    Call SpdCode.GetText(2, Row, vCd)
    
    txtCd = CStr(vCdNm)
    
'    Call SetWindowText(hWndCd, CStr(vCd))
'    Call SetActiveWindow(hWndCd)

End Sub

Private Sub SpdCode_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim vCdNm, vCd

    If Row <> 0 Then
'        Call SpdCode.GetText(1, Row, vCdNm)
        Call SpdCode.GetText(2, Row, vCd)

        Call SetWindowText(hWndCd, CStr(vCd))
'        Call SetWindowText(lblCdNm, CStr(vCdNm))

        Unload Me
    End If

End Sub

Private Sub txtCd_Change()

    ChangeKeyFlag = True

End Sub

Private Sub txtCd_GotFocus()

    txtCd.SelStart = 0

End Sub

Private Sub txtCD_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then CmdClk_Click

End Sub




