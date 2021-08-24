VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FSR0201 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4290
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin Threed.SSPanel pnlbottom 
      Align           =   2  '아래 맞춤
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   4590
      Width           =   4290
      _Version        =   65536
      _ExtentX        =   7567
      _ExtentY        =   926
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
         Height          =   390
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   4155
         _Version        =   65536
         _ExtentX        =   7329
         _ExtentY        =   688
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
   Begin Threed.SSPanel pnlmain 
      Align           =   1  '위 맞춤
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4290
      _Version        =   65536
      _ExtentX        =   7567
      _ExtentY        =   8070
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
         Height          =   3810
         Left            =   60
         OleObjectBlob   =   "FSR0201.frx":0000
         TabIndex        =   3
         Top             =   90
         Width           =   4170
      End
      Begin VB.TextBox txtCd 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   660
         TabIndex        =   4
         Top             =   4050
         Width           =   2175
      End
      Begin Threed.SSCommand CmdClk 
         Height          =   555
         Left            =   2880
         TabIndex        =   5
         Top             =   3930
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   979
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
         Height          =   555
         Left            =   3540
         TabIndex        =   6
         Top             =   3930
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "Esc"
         ForeColor       =   0
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
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   90
         TabIndex        =   7
         Top             =   4125
         Width           =   540
      End
   End
End
Attribute VB_Name = "FSR0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisplayInit()
    txtCd = ""
    
    With SpdCode
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        
        If giCodeHlpMode = 2 Then
            .BackColor = 연빨강
        ElseIf giCodeHlpMode = 3 Then
            .BackColor = 연초록
        End If
        
        .EditModePermanent = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 2
        .Col2 = 4
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        If giCodeHlpMode = 2 Or giCodeHlpMode = 3 Then
            .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = -1
            .Row2 = -1
            .ColHidden = True
            .BlockMode = False
        End If
        
        .MaxRows = 0
        .MaxRows = 15
    End With
End Sub

Private Sub CmdClk_Click()
    Dim i%
    Dim vChk
    Dim vTestNm
    Dim vTestCd
    Dim vTestGbn
    
    MousePointer = 11
    
    Erase gCodeHlpTable
    giCodeHlpCnt = 0

    With SpdCode
        For i = 1 To .MaxRows
            If giCodeHlpMode = 1 Then
                Call .GetText(1, i, vChk)
                
                If vChk = "1" Then
            'FSB0201에서 새로운 항목추가
                    giCodeHlpCnt = giCodeHlpCnt + 1
                    
                    ReDim Preserve gCodeHlpTable(giCodeHlpCnt)
                    
                    Call .GetText(2, i, vTestNm)
                    Call .GetText(3, i, vTestCd)
                    Call .GetText(4, i, vTestGbn)
                    
                    gCodeHlpTable(giCodeHlpCnt).sCodeNm = CStr(vTestNm)
                    gCodeHlpTable(giCodeHlpCnt).sCode = CStr(vTestCd)
                    gCodeHlpTable(giCodeHlpCnt).sGbn = CStr(vTestGbn)
                End If
            End If
                        
            If giCodeHlpMode = 2 Then
                .Row = i
                .Col = 2
                 If .BackColor = 연초록 Then
                    giCodeHlpCnt = giCodeHlpCnt + 1
                    
                    ReDim Preserve gCodeHlpTable(giCodeHlpCnt)
                    
                    Call .GetText(2, i, vTestNm)
                    Call .GetText(3, i, vTestCd)
                    'Call .GetText(4, i, vTestGbn)
                    
                    gCodeHlpTable(giCodeHlpCnt).sCodeNm = CStr(vTestNm)
                    gCodeHlpTable(giCodeHlpCnt).sCode = CStr(vTestCd)
                    'gCodeHlpTable(giCodeHlpCnt).sGbn = CStr(vTestGbn)
                    
                    Exit For
                 End If
            End If
            
            If giCodeHlpMode = 3 Then
                .Row = i
                .Col = 2
                 If .BackColor = RGB(255, 230, 230) Then
                    giCodeHlpCnt = giCodeHlpCnt + 1
                    
                    ReDim Preserve gCodeHlpTable(giCodeHlpCnt)
                    
                    Call .GetText(2, i, vTestNm)
                    Call .GetText(3, i, vTestCd)
                    'Call .GetText(4, i, vTestGbn)
                    
                    gCodeHlpTable(giCodeHlpCnt).sCodeNm = CStr(vTestNm)
                    gCodeHlpTable(giCodeHlpCnt).sCode = CStr(vTestCd)
                    'gCodeHlpTable(giCodeHlpCnt).sGbn = CStr(vTestGbn)
                    
                    Exit For
                 End If
            End If
        Next
        
        .Col = 1
        .Col2 = SpdCode.MaxCols
        .Row = 1
        .Row2 = SpdCode.MaxRows
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
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
    
    For i = 1 To giCodeHlpCnt
        For j = 1 To giCodeHlpCnt
            If Format$(i, "00000") = gCodeHlpTable(j).sSeq Then
                If i > 10 Then
                    SpdCode.MaxRows = i
                End If
                
                Call SpdCode.SetText(2, i, gCodeHlpTable(j).sCodeNm & "")
                Call SpdCode.SetText(3, i, gCodeHlpTable(j).sCode & "")
                Call SpdCode.SetText(4, i, gCodeHlpTable(j).sGbn & "")
            End If
        Next
    Next
    
    Erase gCodeHlpTable
    giCodeHlpCnt = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i%
    
    For i = 1 To giCodeHlpCnt
        With gCallObject
            If giCodeHlpMode = 1 Then
                .MaxRows = .MaxRows + 1
                .BlockMode = True
                .Col = -1
                .Col2 = -1
                .Row = .MaxRows
                .Row2 = .MaxRows
                .BackColor = 연노랑
                .BlockMode = False
                
                Call .SetText(1, .MaxRows, "1")
                Call .SetText(2, .MaxRows, gCodeHlpTable(i).sCode & "")
                Call .SetText(3, .MaxRows, gCodeHlpTable(i).sGbn & "")
                Call .SetText(4, .MaxRows, gCodeHlpTable(i).sCodeNm & "")
            End If
            
            If giCodeHlpMode = 2 Then
                'Call .SetText(2, giCallSpdRow, gCodeHlpTable(i).sCode & "")
                'Call .SetText(3, giCallSpdRow, gCodeHlpTable(i).sGbn & "")
                Call .SetText(2, giCallSpdRow, gCodeHlpTable(i).sCodeNm & "")
            End If
            
            If giCodeHlpMode = 3 Then
                Call .SetText(1, giCallSpdRow, Right$(gCodeHlpTable(i).sCode, 4) & "")
                'Call .SetText(3, giCallSpdRow, gCodeHlpTable(i).sGbn & "")
                Call .SetText(3, giCallSpdRow, gCodeHlpTable(i).sCodeNm & "")
            End If
        End With
    Next
    
End Sub

Private Sub SpdCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vCd
    Dim vCdNm
    Dim vGbn
    
    If Row = 0 Then
        Exit Sub
    End If
    
    If giCodeHlpMode = 2 Then
        Call spdReverse(SpdCode, -1, -1, Row, Row, 연초록, 연빨강)
    ElseIf giCodeHlpMode = 3 Then
        Call spdReverse(SpdCode, -1, -1, Row, Row, RGB(255, 230, 230), 연초록)
    End If
    
    If Col <> 1 Then
        SpdCode.Col = 1
        SpdCode.Row = Row
        
        If SpdCode.Text = "" Or SpdCode.Text = "0" Then
            SpdCode.Text = "1"
        Else
            SpdCode.Text = ""
        End If
    End If
    
    Call SpdCode.GetText(2, Row, vCdNm)
        
    txtCd = CStr(vCdNm)
    
End Sub

Private Sub SpdCode_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row <> 0 Then
        Call CmdClk_Click
        Unload Me
    End If
End Sub
