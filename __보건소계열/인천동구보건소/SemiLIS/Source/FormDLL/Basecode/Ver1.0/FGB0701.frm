VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0701 
   Caption         =   "기초자료 - COMMENT"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "FGB0701.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11280
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread spdBaseCode 
      Height          =   4605
      Left            =   420
      OleObjectBlob   =   "FGB0701.frx":030A
      TabIndex        =   10
      Top             =   2370
      Width           =   8295
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   420
      TabIndex        =   11
      Top             =   0
      Width           =   8280
      Begin VB.OptionButton optComGbn 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Interface입력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   3
         Top             =   690
         Width           =   1545
      End
      Begin VB.OptionButton optComGbn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "소견입력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3570
         TabIndex        =   2
         Top             =   690
         Width           =   1275
      End
      Begin VB.OptionButton optComGbn 
         BackColor       =   &H00C0E0FF&
         Caption         =   "결과입력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2100
         TabIndex        =   1
         Top             =   690
         Width           =   1275
      End
      Begin VB.TextBox txtPartCd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '영문
         Left            =   2070
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "H"
         Top             =   285
         Width           =   300
      End
      Begin VB.TextBox txtComNm 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2070
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   5
         Text            =   "FGB0701.frx":065F
         Top             =   1425
         Width           =   5685
      End
      Begin VB.TextBox txtComCd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '영문
         Left            =   2070
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "001"
         Top             =   1020
         Width           =   510
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   345
         Left            =   270
         TabIndex        =   12
         Top             =   1040
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Comment 코드"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel Panel3D1 
         Height          =   345
         Left            =   270
         TabIndex        =   13
         Top             =   1425
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Comment 내용"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   65535
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Left            =   270
         TabIndex        =   14
         Top             =   270
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "PART 코드"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdButtonPart 
         Height          =   330
         Left            =   2370
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   285
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGB0701.frx":067A
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   270
         TabIndex        =   17
         Top             =   655
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "코드 종류 구분"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin VB.Label lblPartNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   16
         Top             =   285
         Width           =   3885
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   9480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "삭제 F4"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0701.frx":079C
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   9480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "조회 F3"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0701.frx":1076
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   9480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "종료 ESC"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0701.frx":1950
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   1005
      Left            =   9480
      TabIndex        =   6
      Top             =   660
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "등록 F2"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0701.frx":222A
   End
End
Attribute VB_Name = "FGB0701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCurSelRow%
Dim iSpdClick1%
Dim iSpdClick2%

Private Sub InsertOrUpdate()
    Dim i%
    Dim vComCd, vComNm
    Dim sGbn$
    
    If optComGbn(0).Value = True Then   '결과입력
        sGbn = "E"
    ElseIf optComGbn(1).Value = True Then   '소견입력
        sGbn = "S"
    ElseIf optComGbn(2).Value = True Then   'Interface입력
        sGbn = "I"
    End If
    
    With spdBaseCode
        For i = 1 To .MaxRows
            Call .GetText(1, i, vComCd)
            
            If CStr(vComCd) = txtPartCd & sGbn & txtComCd Then
                Call .GetText(2, i, vComNm)
                
                If CStr(vComNm) = txtComNm Then
                    MsgBox "이미 존재하는 항목입니다"
                    Exit Sub
                Else
                    Call UpdateComCd(i)
                    Exit Sub
                End If
            End If
        Next
    End With
    
    Call InsertComCd
    
End Sub

Private Sub InsertComCd()
    Dim CCom As DCB0101
    Dim sGbn$
    
    Set CCom = New DCB0101
    
    ViewMsg ""
    
    If optComGbn(0).Value = True Then   '결과입력
        sGbn = "E"
    ElseIf optComGbn(1).Value = True Then   '소견입력
        sGbn = "S"
    ElseIf optComGbn(2).Value = True Then   'Interface입력
        sGbn = "I"
    End If
    
    CCom.Add_COMCD txtPartCd, sGbn & txtComCd, txtComNm
    
    If CCom.AdoErrNum = 0 Then
        ViewMsg "등록작업이 성공적으로 수행되었습니다..."
                
        spdBaseCode.MaxRows = spdBaseCode.MaxRows + 1
        
        Call spdBaseCode.SetText(1, spdBaseCode.MaxRows, txtPartCd & sGbn & txtComCd & "")
        Call spdBaseCode.SetText(2, spdBaseCode.MaxRows, txtComNm & "")
    End If
    
    Set CCom = Nothing
End Sub

Private Sub UpdateComCd(ByVal iRow As Integer)
    Dim CCom As DCB0101
    Dim sGbn$
    
    Set CCom = New DCB0101
    
    ViewMsg ""
    
    If optComGbn(0).Value = True Then   '결과입력
        sGbn = "E"
    ElseIf optComGbn(1).Value = True Then   '소견입력
        sGbn = "S"
    ElseIf optComGbn(2).Value = True Then   'Interface입력
        sGbn = "I"
    End If
    
    CCom.Edit_COMCD txtPartCd, sGbn & txtComCd, txtComNm
    
    If CCom.AdoErrNum = 0 Then
        ViewMsg "변경작업이 성공적으로 수행되었습니다..."
                
        Call spdBaseCode.SetText(1, iRow, txtPartCd & sGbn & txtComCd & "")
        Call spdBaseCode.SetText(2, iRow, txtComNm & "")
        
    End If
    
    Set CCom = Nothing
End Sub

Private Sub SelectComCd(ByVal iMode As Integer)
    Dim CCom As DCB0101
    Dim sField01$, sField02$, sField03$
    Dim i%, j%
    
    Set CCom = New DCB0101
    
    If iMode = 1 Then   'All Search
        CCom.Get_COMCD
        
    ElseIf iMode = 2 Then   '결과입력 Comment
        CCom.Get_COMCD txtPartCd, "E"
            
    ElseIf iMode = 3 Then   '소견입력 Comment
        CCom.Get_COMCD txtPartCd, "S"
        
    ElseIf iMode = 4 Then   'Interface입력 Comment
        CCom.Get_COMCD txtPartCd, "I"
        
    ElseIf iMode = 5 Then
        CCom.Get_COMCD txtPartCd, "E" & txtComCd
    
    ElseIf iMode = 6 Then
        CCom.Get_COMCD txtPartCd, "S" & txtComCd
    
    ElseIf iMode = 7 Then
        CCom.Get_COMCD txtPartCd, "I" & txtComCd
    
    End If
    
    i = CCom.CurItemCnt
    
    If i = 0 Then
        MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
        Set CCom = Nothing
        Exit Sub
    End If
                
    sField01 = CCom.TotField01
    sField02 = CCom.TotField02
    sField03 = CCom.TotField03
    
    spdBaseCode.MaxRows = 0
    
    For j = 1 To i
        With spdBaseCode
            .MaxRows = j
            If iMode = 1 Then
                Call .SetText(1, j, GetByOne(sField01, sField01) & GetByOne(sField02, sField02) & "")
                Call .SetText(2, j, GetByOne(sField03, sField03) & "")
                optComGbn(0).Value = False
                optComGbn(1).Value = False
                optComGbn(2).Value = False
            Else
                Call .SetText(1, j, txtPartCd & GetByOne(sField01, sField01) & "")
                Call .SetText(2, j, GetByOne(sField02, sField02) & "")
                txtComCd = ""
                txtComNm = ""
            End If
        End With
    Next
    
    Set CCom = Nothing
    
End Sub

Private Sub DisplayInit()
    
    txtPartCd = fCurUserPartCd
    lblPartNm = fCurUserPartNm
    txtComCd = ""
    optComGbn(0).Value = True
    optComGbn(1).Value = False
    optComGbn(2).Value = False
    txtComNm = ""
    
    Me.KeyPreview = True
            
    'SpreadBackColor Option
    iSpdBackColorOption = 2
    
    With spdBaseCode
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = SpdBackcolor(iSpdBackColorOption)   'GBR
        .EditModePermanent = True
        .Protect = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Col2 = 2
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        .MaxRows = 0
    End With
End Sub

Private Sub DisplayComCd()
    Dim i%
    Dim sGbn$
    
    Dim CCom As DCB0101
    
    Set CCom = New DCB0101
    
    If optComGbn(0).Value = True Then
        sGbn = "E"
    ElseIf optComGbn(1).Value = True Then
        sGbn = "S"
    ElseIf optComGbn(2).Value = True Then
        sGbn = "I"
    End If
    
    CCom.Get_COMCD txtPartCd, sGbn & txtComCd
    
    i = CCom.CurItemCnt
    
    If i = 0 Then
        txtComNm = ""
        Set CCom = Nothing
        Exit Sub
    End If
    
    txtComNm = GetByOne(CCom.TotField02, CCom.TotField02)
        
    Call FindCurSpreadRow(txtPartCd, sGbn & txtComCd, txtComNm)
    
    Set CCom = Nothing
    
End Sub

Private Sub FindCurSpreadRow(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String)
    Dim vField01, vField02
    Dim iCurRow%
    Dim i%
    
    iCurRow = 0
    
    With spdBaseCode
        For i = 1 To .MaxRows
            Call .GetText(1, i, vField01)
            Call .GetText(2, i, vField02)
                        
            If Left$(CStr(vField01), 1) = sField01 And Right$(CStr(vField01), 4) = sField02 And CStr(vField02) = sField03 Then
                iCurRow = i
                Exit For
            End If
        Next
    
        If iCurRow = 0 Then
            .MaxRows = 1
            Call .SetText(1, .MaxRows, sField01 & sField02 & "")
            Call .SetText(2, .MaxRows, sField03 & "")
        Else
            Call spdReverse(spdBaseCode, -1, -1, iCurRow, iCurRow, RGB(255, 230, 230), 2)
        End If
    End With
End Sub

Private Sub cmdButtonPart_Click()
    Dim i%
    Dim iDefaultSeq%
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(giPartCnt) As CodeTBL
    
    For i = 1 To giPartCnt
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = gPartTable(i).sPartInit
        gCodeHlpTable(i).sCodeNm = gPartTable(i).sPartName
    Next
        
    giCodeHlpCnt = giPartCnt
    
    hWndCd = txtPartCd.hwnd
    
    FSB0101.Left = 3500
    FSB0101.Top = 1750
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdDelete_Click()
    Dim iRetVal%
    Dim CCom As DCB0101
    Dim sGbn$
    
    If txtPartCd = "" Then
        Exit Sub
    Else
        If lblPartNm = "" Or txtComNm = "" Then
            Exit Sub
        End If
        
        If optComGbn(0).Value = True Then
            sGbn = "E"
        ElseIf optComGbn(1).Value = True Then
            sGbn = "S"
        ElseIf optComGbn(2).Value = True Then
            sGbn = "I"
        End If
        
        iRetVal = MsgBox("COMMENT 코드 : " & txtPartCd & sGbn & txtComCd & vbCrLf & _
                "COMMENT 내용 : " & txtComNm & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "COMMENT 코드 삭제 확인")
    
        If iRetVal = 1 Then
            Set CCom = New DCB0101
            
            CCom.Delete_COMCD txtPartCd, sGbn & txtComCd
            
            If CCom.AdoErrNum = 0 Then
                ViewMsg "삭제작업이 성공적으로 이루어졌습니다..."
                
                With spdBaseCode
                    .Row = iCurSelRow
                    .Action = SS_ACTION_DELETE_ROW
                    .MaxRows = .MaxRows - 1
                End With
                
            End If
            
            Set CCom = Nothing
        Else
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    If txtPartCd = "" Then
        MsgBox "PART 코드, 코드종류 구분, COMMENT 코드, COMMENT 내용을 모두 입력한 후 등록해 주십시요!!"
    Else
        If optComGbn(0).Value = False And optComGbn(1).Value = False And optComGbn(2).Value = False Then
            MsgBox "PART 코드, 코드종류 구분, COMMENT 코드, COMMENT 내용을 모두 입력한 후 등록해 주십시요!!"
        Else
            If txtComCd = "" Then
                MsgBox "PART 코드, 코드종류 구분, COMMENT 코드, COMMENT 내용을 모두 입력한 후 등록해 주십시요!!"
            Else
                If txtComNm = "" Then
                    MsgBox "PART 코드, 코드종류 구분, COMMENT 코드, COMMENT 내용을 모두 입력한 후 등록해 주십시요!!"
                Else
                    Call InsertOrUpdate
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim iRetVal%
    
    If txtPartCd = "" Then
        iRetVal = MsgBox("모든 PART의 결과입력, 소견입력, Interface입력 코드를 " & vbCrLf & _
                "조회하므로 시간이 걸립니다. 계속 진행하시겠습니까?", _
                 vbOKCancel, "COMMENT 코드 조회 조건 확인")
        
        If iRetVal = 1 Then
            
            Call SelectComCd(1)
            
        End If
    Else
        If txtComCd = "" Then
            If optComGbn(0).Value = True Then
                Call SelectComCd(2)
            ElseIf optComGbn(1).Value = True Then
                Call SelectComCd(3)
            ElseIf optComGbn(2).Value = True Then
                Call SelectComCd(4)
            Else
                MsgBox "PART 코드를 비워두고 전체를 조회하거나 코드종류 구분 중 하나를 선택하고 조회하세요!!"
            End If
            
        Else
            If optComGbn(0).Value = True Then
                Call SelectComCd(5)
            ElseIf optComGbn(1).Value = True Then
                Call SelectComCd(6)
            ElseIf optComGbn(2).Value = True Then
                Call SelectComCd(7)
            End If
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    If txtPartCd = "" Then
        txtPartCd.SetFocus
    ElseIf txtComCd = "" Then
        txtComCd.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonPart_Click
        Case vbKeyF2:        Call cmdReg_Click
        Case vbKeyF3:        Call cmdSearch_Click
        Case vbKeyF4:        Call cmdDelete_Click
        Case vbKeyEscape:    Call cmdExit_Click
    End Select
End Sub

Private Sub Form_Load()
    If Me.StartUpPosition = 2 Then
    Else
        Me.Left = 250
        Me.Top = 10
        Me.Width = 11400
        Me.Height = 7500
    End If
    
    iSpdClick1 = 0
    iSpdClick2 = 0
    
    Call DisplayInit
    Call cmdSearch_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
    ViewMsg ""
End Sub

Private Sub optComGbn_Click(Index As Integer)
    txtComCd = ""
    txtComNm = ""
End Sub

Private Sub optComGbn_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtComCd.SetFocus
    End If
End Sub

Private Sub spdBaseCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vField01
    Dim vField02
    
    If Row = 0 Then
        Exit Sub
    End If
    
    iCurSelRow = CInt(Row)
    
    iSpdClick1 = 1
    iSpdClick2 = 1
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
    
    txtPartCd = Left$(CStr(vField01), 1)
    
    If Mid$(CStr(vField01), 2, 1) = "E" Then
        optComGbn(0).Value = True
    ElseIf Mid$(CStr(vField01), 2, 1) = "S" Then
        optComGbn(1).Value = True
    ElseIf Mid$(CStr(vField01), 2, 1) = "I" Then
        optComGbn(2).Value = True
    End If
    
    txtComCd = Right$(CStr(vField01), 3)
    
    txtComNm = CStr(vField02)
    
End Sub

Private Sub txtComCd_Change()
    
    
    If Len(txtComCd) = txtComCd.MaxLength Then
        If txtPartCd = "" Or _
           (optComGbn(0).Value = False And optComGbn(1).Value = False And optComGbn(2).Value = False) Then
            
            MsgBox "PART 코드를 입력한 후 COMMENT 코드를 입력하여 주십시요!!"
            Exit Sub
        Else
            If iSpdClick2 = 1 Then
            Else
                Call DisplayComCd
            End If
        End If
        
        iSpdClick2 = 0
        
        txtComNm.SetFocus
    ElseIf Len(txtComCd) = 0 Then
        txtComNm = ""
    End If
End Sub

Private Sub txtComCd_Click()
    Call Txt_Highlight(txtComCd)
End Sub

Private Sub txtComCd_GotFocus()
    Call Txt_Highlight(txtComCd)
End Sub

Private Sub txtComCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtComNm.SetFocus
    End If
End Sub

Private Sub txtComCd_LostFocus()
    If Len(txtComCd) < txtComCd.MaxLength Then
        txtComCd = Format$(txtComCd, "000")
    End If
End Sub

Private Sub txtComNm_Click()
    Call Txt_Highlight(txtComNm)
End Sub

Private Sub txtComNm_GotFocus()
    Call Txt_Highlight(txtComNm)
End Sub

Private Sub txtPartCd_Change()
    On Error GoTo ErrHandler
    
    Dim i%
    
    If Len(txtPartCd) = txtPartCd.MaxLength Then
        
        If iSpdClick1 = 1 Then
        Else
            For i = 1 To giPartCnt
                lblPartNm = ""
                If gPartTable(i).sPartInit = txtPartCd Then
                    lblPartNm = gPartTable(i).sPartName
                    optComGbn(0).SetFocus
                    Exit For
                End If
            Next
        End If
        
    ElseIf Len(txtPartCd) = 0 Then
        lblPartNm = ""
        optComGbn(0).Value = False
        optComGbn(1).Value = False
        optComGbn(2).Value = False
        txtComCd = ""
        txtComNm = ""
    End If
    
    iSpdClick1 = 0
    
ErrHandler:
End Sub

Private Sub txtPartCd_Click()
    Call Txt_Highlight(txtPartCd)
End Sub

Private Sub txtPartCd_GotFocus()
    Call Txt_Highlight(txtPartCd)
End Sub

Private Sub txtPartCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If optComGbn(0).Value = True Then
            optComGbn(0).SetFocus
        ElseIf optComGbn(1).Value = True Then
            optComGbn(1).SetFocus
        ElseIf optComGbn(2).Value = True Then
            optComGbn(2).SetFocus
        End If
    End If
End Sub
