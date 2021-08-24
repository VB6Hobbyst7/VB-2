VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form comFRMCdHlp 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   3945
   ClientLeft      =   1800
   ClientTop       =   2145
   ClientWidth     =   4065
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   4065
   Begin Threed.SSPanel pnlbottom 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   3540
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   714
      _StockProps     =   15
      ForeColor       =   16576
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
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
         TabIndex        =   6
         Top             =   30
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   8388608
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
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
   Begin Threed.SSPanel pnlMain 
      Align           =   1  '위 맞춤
      Height          =   3525
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
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
      Begin FPSpread.vaSpread spdDisp 
         Height          =   2985
         Left            =   90
         OleObjectBlob   =   "comFRMCdHlp.frx":0000
         TabIndex        =   7
         Top             =   90
         Width           =   3150
      End
      Begin VB.TextBox txtCd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   690
         TabIndex        =   0
         Top             =   3150
         Width           =   2145
      End
      Begin Threed.SSCommand CmdView 
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "View"
         ForeColor       =   255
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
         Left            =   3420
         TabIndex        =   3
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Esc"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "코드명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   3195
         Width           =   540
      End
   End
End
Attribute VB_Name = "comFRMCdHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' 99/2/3 ojm 수정
'
Option Explicit

Private Const BarColor = &HFFE8DD


Private Sub CmdEsc_Click()
    
    D0COM_CODEHELP.clsLCol = 1:    D0COM_CODEHELP.clsLDBValue = ""
    Unload Me

End Sub


Private Sub CmdView_Click()

    With D0COM_CODEHELP
        .clsLCol = 1
        .clsLQryWhere2 = txtCd.Text
        .subDisplayCodeHLP
    End With
    
End Sub


Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    Dim setX    As Integer, setY    As Integer
    
    Dim WidthFrm    As Long, HeightFrm  As Long
    Dim sqlConn     As Long, ret        As Integer
    
    Me.Width = Screen.Width
    
    With spdDisp
        .VirtualMode = True
        .Col = 1:   .Col2 = .MaxCols
        .Row = 2:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
        
        D0COM_CODEHELP.subDefindCodeHLP
        
        .AutoSize = False
    
    End With
    
    ret = DoEvents
    Me.Width = Screen.Width
    
    WidthFrm = D0COM_CODEHELP.clsGFormWidth
    HeightFrm = D0COM_CODEHELP.clsGFormHeight
    
    Me.Width = WidthFrm:    Me.Height = HeightFrm
    
    If D0COM_CODEHELP.clsGSqlConn = 0 And D0COM_CODEHELP.clsGGbnmenberCnt = 0 Then
         '본 화면에서 사용할 Index Open
        If Not QSqlOpen(D0COM_SERVER01, Me.hWnd, sqlConn) _
             = QSQL_SUCCESS Then Exit Sub
    
        D0COM_CODEHELP.clsLSqlConn = sqlConn
        D0COM_CODEHELP.clsLDBOpen = True
    End If

    D0COM_CODEHELP.subDisplayCodeHLP
    
    spdDisp.VirtualMode = False
    
    setX = D0COM_CODEHELP.clsGPosX: setY = D0COM_CODEHELP.clsGPosY
    
    If Me.Width + setX > D0COM_CODEHELP.clsGCallForm.Width Then
        setX = D0COM_CODEHELP.clsGCallForm.Width - Me.Width
    End If
    
    If Me.Height + setY > D0COM_CODEHELP.clsGCallForm.Height Then
        setY = setY - Me.Height - D0COM_CODEHELP.clsGObjectHeight
    End If
    
    Me.Move setX, setY
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim DbFlag  As Integer
    
    DbFlag = D0COM_CODEHELP.clsGDBOpen
    
    If DbFlag = True Then Call QSqlClose(D0COM_CODEHELP.clsGSqlConn)
    
End Sub

Private Sub spdDisp_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If Row = 1 Then D0COM_CODEHELP.subSpreadDataSort (CInt(Col))
    
    '--- 99/3/24
    If Row = 1 Then
        With spdDisp
            Dim ix9     As Integer
            Dim iPosX1  As Integer
            Dim iPosX2  As Integer
            
            For ix9 = 1 To .MaxCols
                .Col = ix9: .Row = 1
                                
                iPosX1 = InStr(.TypeButtonText, "▼")
                iPosX2 = InStr(.TypeButtonText, "▲")
                If iPosX1 <> 0 Then
                    .TypeButtonText = Mid(.TypeButtonText, 1, iPosX1 - 2)
                    If Col = ix9 Then
                        .TypeButtonText = .TypeButtonText & " ▲"
                        Exit Sub
                    End If
                    iPosX1 = 0
                ElseIf iPosX2 <> 0 Then
                    .TypeButtonText = Mid(.TypeButtonText, 1, iPosX2 - 2)
                    If Col = ix9 Then
                        .TypeButtonText = .TypeButtonText & " ▼"
                        Exit Sub
                    End If
                    iPosX2 = 0
                End If
            Next ix9
            
            .Col = Col: .Row = 1
            If iPosX1 = 0 And iPosX2 = 0 Then .TypeButtonText = .TypeButtonText & " ▼"
        End With
    End If

End Sub

Private Sub spdDisp_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row <= 1 Then Exit Sub
    
    Dim idx As Integer
    
    With spdDisp
    
    .Row = Row
    
    For idx = 1 To .MaxCols
        .Col = idx
        D0COM_CODEHELP.clsLCol = idx:   D0COM_CODEHELP.clsLDBValue = .Text
    Next
    
    End With
    
    Unload Me
    
End Sub


Private Sub spdDisp_KeyDown(KeyCode As Integer, Shift As Integer)

    With spdDisp
    If .ActiveRow = 1 Then
        If KeyCode = vbKeyLeft Then
            If .ActiveCol = 1 Then
                .Col = 3
            Else
                .Col = .ActiveCol - 1
            End If
        ElseIf KeyCode = vbKeyRight Then
            If .ActiveCol < 3 Then
                .Col = .ActiveCol + 1
            Else
                .Col = 1
            End If
        End If
        .Action = SS_ACTION_ACTIVE_CELL
        Exit Sub
    End If
    
    If Not .ActiveRow = 2 Or Not KeyCode = vbKeyUp Then Exit Sub
        
    .Col = 1: .Row = 3
    .Action = SS_ACTION_ACTIVE_CELL
    
    End With
    
End Sub

Private Sub spdDisp_KeyPress(KeyAscii As Integer)

    If spdDisp.ActiveRow < 1 Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If spdDisp.ActiveRow > 1 Then
           Call spdDisp_DblClick(1, spdDisp.ActiveRow)
        End If
    End If
    
End Sub


Private Sub spdDisp_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    If NewRow < 1 Then Exit Sub
    
    Dim BackColor   As Variant
    
    With spdDisp
    If NewRow > 2 Then
        .Col = 1:   .Row = NewRow
        BackColor = .BackColor
    ElseIf Row = .MaxRows Then
        .Col = 1:   .Row = .MaxRows - 1
        BackColor = .BackColor
    Else
        .Col = 1:   .Row = .MaxRows
        BackColor = .BackColor
    End If
    
    .Col = 1:       .Col2 = .MaxCols
    .Row = Row:     .Row2 = Row
    .BlockMode = True
    .BackColor = BackColor
    .BlockMode = False

    .Col = 1:       .Col2 = .MaxCols
    .Row = NewRow:  .Row2 = NewRow
    .BlockMode = True
    .BackColor = BarColor
    .BlockMode = False
    
    End With

End Sub


Private Sub txtCD_GotFocus()

    txtCd.SelStart = 0
    txtCd.SelLength = Len(txtCd.Text)

End Sub

Private Sub txtCD_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call CmdView_Click: KeyAscii = 0
    End If

End Sub




