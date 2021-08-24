VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0101 
   Caption         =   "기초자료 - SLIP"
   ClientHeight    =   7095
   ClientLeft      =   495
   ClientTop       =   1110
   ClientWidth     =   11280
   Icon            =   "FGB0101.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11280
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread spdBaseCode 
      Height          =   4605
      Left            =   480
      OleObjectBlob   =   "FGB0101.frx":030A
      TabIndex        =   13
      Top             =   2160
      Width           =   8115
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   480
      TabIndex        =   7
      Top             =   150
      Width           =   8100
      Begin VB.TextBox txtSlip 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Top             =   1305
         Width           =   5355
      End
      Begin VB.TextBox txtPartGbn 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '영문
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "H2"
         Top             =   810
         Width           =   390
      End
      Begin VB.TextBox txtPart 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "H"
         Top             =   360
         Width           =   300
      End
      Begin Threed.SSPanel Panel3D7 
         Height          =   375
         Left            =   270
         TabIndex        =   8
         Top             =   810
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "PART 구분"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   375
         Left            =   270
         TabIndex        =   9
         Top             =   345
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "PART 코드"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel Panel3D1 
         Height          =   375
         Left            =   270
         TabIndex        =   10
         Top             =   1305
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SLIP 명"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   65535
      End
      Begin Threed.SSCommand cmdButtonPart 
         Height          =   330
         Left            =   2100
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
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
         Picture         =   "FGB0101.frx":13E5
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   2490
         TabIndex        =   14
         Top             =   810
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SLIP 코드"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin VB.Label lblSlip 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "H01"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4020
         TabIndex        =   15
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label lblPart 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
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
         Left            =   2490
         TabIndex        =   12
         Top             =   360
         Width           =   3885
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   9390
      TabIndex        =   5
      Top             =   2790
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
      Picture         =   "FGB0101.frx":1507
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   9390
      TabIndex        =   4
      Top             =   1770
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
      Picture         =   "FGB0101.frx":1DE1
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   9390
      TabIndex        =   6
      Top             =   3810
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
      Picture         =   "FGB0101.frx":26BB
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   1005
      Left            =   9390
      TabIndex        =   3
      Top             =   750
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
      Picture         =   "FGB0101.frx":2F95
   End
End
Attribute VB_Name = "FGB0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSpdClick%

Private Function CompareSpread() As Integer
    Dim sComp1$
    Dim sComp2$
    Dim i%
    Dim vPart, vPartGbn, vSlip
    
    CompareSpread = 0
    sComp1 = Left$(lblSlip.Caption, 1)
    sComp2 = Right$(lblSlip.Caption, 2)
    
    If spdBaseCode.MaxRows > 0 Then
        For i = 1 To spdBaseCode.MaxRows
            Call spdBaseCode.GetText(1, i, vPart)
            Call spdBaseCode.GetText(2, i, vPartGbn)
            
            If sComp1 = vPart And sComp2 = vPartGbn Then
                Call spdBaseCode.GetText(3, i, vSlip)
                txtSlip = CStr(vSlip)
                
                Call spdReverse(spdBaseCode, -1, -1, i, i, RGB(255, 230, 230), 2)
                CompareSpread = i
                Exit For
            End If
        Next
    End If
    
    If CompareSpread = 0 Then
        txtSlip = ""
    End If
End Function

Private Sub ShortKeyOrTabOrderInit()
    Me.KeyPreview = True
    
    txtPart.TabIndex = 0
    txtPartGbn.TabIndex = 1
    txtSlip.TabIndex = 2
    cmdReg.TabIndex = 3
    cmdSearch.TabIndex = 4
    cmdDelete.TabIndex = 5
    cmdExit.TabIndex = 6
    
End Sub

Private Sub DisplayInit()
    txtPart = ""
    lblPart.Caption = ""
    txtPartGbn = ""
    lblSlip.Caption = ""
    txtSlip = ""
    
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
        .Col2 = 3
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        .MaxRows = 0
    End With
End Sub

Private Sub DisplaySlip()
    Dim i%
    Dim sPart$, sPartGbn$, sSlipNm$
    Dim CPart As DCB0101
    
    Set CPart = New DCB0101
           
    CPart.Get_PART txtPart, txtPartGbn
    
    i = CPart.CurItemCnt
    
    If i = 0 Then
        txtSlip = ""
        lblSlip.Caption = ""
        
        Set CPart = Nothing
        Exit Sub
    End If
    
    sPart = CPart.TotField01
    sPartGbn = CPart.TotField02
    sSlipNm = CPart.TotField03
    
    lblSlip.Caption = GetByOne(sPart, sPart) & GetByOne(sPartGbn, sPartGbn)
    txtSlip = GetByOne(sSlipNm, sSlipNm)
    
    If spdBaseCode.MaxRows = 0 Then
        spdBaseCode.MaxRows = 1
        Call spdBaseCode.SetText(1, spdBaseCode.MaxRows, txtPart & "")
        Call spdBaseCode.SetText(2, spdBaseCode.MaxRows, txtPartGbn & "")
        Call spdBaseCode.SetText(3, spdBaseCode.MaxRows, txtSlip & "")
    Else
        Call FindCurSpreadRow(txtPart, txtPartGbn, txtSlip)
    End If
    
    Set CPart = Nothing
    
End Sub

Private Sub FindCurSpreadRow(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String)
        Dim vField01, vField02, vField03
        Dim iCurRow%
        Dim i%
        
        iCurRow = 0
        
        With spdBaseCode
            For i = 1 To .MaxRows
                Call .GetText(1, i, vField01)
                Call .GetText(2, i, vField02)
                Call .GetText(3, i, vField03)
                
                If CStr(vField01) = sField01 And CStr(vField02) = sField02 And CStr(vField03) = sField03 Then
                    iCurRow = i
                    Exit For
                End If
            Next
        
            If iCurRow = 0 Then
                .MaxRows = 1
                Call .SetText(1, .MaxRows, sField01 & "")
                Call .SetText(2, .MaxRows, sField02 & "")
                Call .SetText(3, .MaxRows, sField03 & "")
            Else
                Call spdReverse(spdBaseCode, -1, -1, iCurRow, iCurRow, RGB(255, 230, 230), 2)
            End If
        End With
        
        
End Sub

Private Sub BaseCodeInit()
    Dim CPart As DCB0101
    Dim i%
    Dim j%
    Dim sPartCd$
    Dim sPartGbn$
    Dim sSlipName$
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    i = CPart.CurItemCnt
    
    If i = 0 Then
        MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
        Set CPart = Nothing
        Exit Sub
    End If
    
    sPartCd = CPart.TotField01
    sPartGbn = CPart.TotField02
    sSlipName = CPart.TotField03
    
    For j = 1 To i
        spdBaseCode.MaxRows = j
        Call spdBaseCode.SetText(1, j, GetByOne(sPartCd, sPartCd) & "")
        Call spdBaseCode.SetText(2, j, GetByOne(sPartGbn, sPartGbn) & "")
        Call spdBaseCode.SetText(3, j, GetByOne(sSlipName, sSlipName) & "")
    Next
    
    Set CPart = Nothing
    
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
    
    hWndCd = txtPart.hwnd
    
    FSB0101.Left = 5300 '3300
    FSB0101.Top = 1950
    
    Load FSB0101
    FSB0101.Show vbModal
    
    '-- 2001.07.06추가(osw)
    txtPartGbn = ""
    lblSlip.Caption = ""
    txtSlip = ""
    '----------------------

End Sub

Private Sub cmdDelete_Click()
    On Err GoTo ErrHandler
    
    Dim vDefault
    Dim CPart As DCB0101
    Dim iRetVal%
    
    If CompareSpread > 0 Then
        
        iRetVal = MsgBox("SLIP 코드 : " & lblSlip & vbCrLf & _
                "SLIP 명 : " & txtSlip & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "SLIP 코드 삭제 확인")
        
        If iRetVal = 1 Then
            Set CPart = New DCB0101
            
            CPart.Delete_PART txtPart, txtPartGbn
            
            Set CPart = Nothing
            
            With spdBaseCode
                .Row = CompareSpread
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            End With
            
            '-- 2001.07.06(osw)
            txtPart = ""
            lblPart.Caption = ""
            txtPartGbn = ""
            lblSlip.Caption = ""
            txtSlip = ""
            '-------------------
        
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case Err.Number
        Case 13
            MsgBox Err.Description, vbInformation, "확인"
        Case Else
            MsgBox Err.Description, vbCritical, "오류"
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    On Error GoTo ErrHandler
    
    Dim CPart As DCB0101
    Dim vPart, vPartGbn, vSlip
    Dim i%
    Dim bMatch As Boolean
    
    bMatch = False
    
    If lblPart = "" Then
        MsgBox "PART가 존재하지 않습니다!!"
        Exit Sub
    End If
    
    If txtPart = "" Then
        MsgBox "PART 1자리와 PART 구분 2자리의 코드가 필요합니다!!"
        Exit Sub
    Else
        If txtPartGbn = "" Then
            MsgBox "PART 1자리와 PART 구분 2자리의 코드가 필요합니다!!"
            Exit Sub
        End If
    End If
    
    If spdBaseCode.MaxRows > 0 Then
        For i = 1 To spdBaseCode.MaxRows
            Call spdBaseCode.GetText(1, i, vPart)
            Call spdBaseCode.GetText(2, i, vPartGbn)
            
            If txtPart = vPart And txtPartGbn = vPartGbn Then
                Call spdBaseCode.GetText(3, i, vSlip)
                If vSlip = txtSlip Then
                    MsgBox "기존의 존재하는 데이터와 일치합니다", vbInformation, "확인"
                    Exit Sub
                Else
                    'EditItem - Slip명이 틀려진 경우
                    Set CPart = New DCB0101
                    
                    CPart.Edit_PART txtPart, txtPartGbn, Left(txtSlip, 40)
                    
                    If CPart.AdoErrNum = 0 Then
                        '화면에 반영
                        With spdBaseCode
                            Call .SetText(1, i, txtPart & "")
                            Call .SetText(2, i, txtPartGbn & "")
                            Call .SetText(3, i, txtSlip & "")
                        End With
                        
                        txtPartGbn.SetFocus
                    End If
                    
                    Set CPart = Nothing
                End If
                
                bMatch = True
                
                Exit For
            Else
                bMatch = False
            End If
        Next
    End If
    
    If bMatch = False Then
        Set CPart = New DCB0101
        
        CPart.Add_PART txtPart, txtPartGbn, Left(txtSlip, 40)
        
        If CPart.AdoErrNum = 0 Then
            '화면에 반영
            With spdBaseCode
                 .MaxRows = .MaxRows + 1
                Call .SetText(1, spdBaseCode.MaxRows, txtPart & "")
                Call .SetText(2, spdBaseCode.MaxRows, txtPartGbn & "")
                Call .SetText(3, spdBaseCode.MaxRows, txtSlip & "")
            End With
            
            txtPartGbn.SetFocus
        End If
        
        Set CPart = Nothing
        
    End If
                
    Exit Sub
    
ErrHandler:
    MsgBox Err.Number
    MsgBox Err.Description
    
End Sub

Private Sub cmdSearch_Click()
    Call DisplayInit
    Call BaseCodeInit
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
    Me.Left = 250
    Me.Top = 10
    Me.Width = 11400
    Me.Height = 7500
    
    'SpreadClick구별 키 초기화
    iSpdClick = 0
    
    Call DisplayInit
    Call ShortKeyOrTabOrderInit
    'Call BaseCodeInit
    
    Call cmdSearch_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
End Sub

Private Sub spdBaseCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vPartCd
    Dim vPartGbn
    Dim vSlipNm
    Dim vTmp
    
    If Row = 0 Then
        Exit Sub
    End If
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    
    Call spdBaseCode.GetText(1, Row, vPartCd)
    Call spdBaseCode.GetText(2, Row, vPartGbn)
    Call spdBaseCode.GetText(3, Row, vSlipNm)
    
    iSpdClick = 1
    
    txtPart = CStr(vPartCd)
    txtPartGbn = CStr(vPartGbn)
    lblSlip = CStr(vPartCd) & CStr(vPartGbn)
    txtSlip = CStr(vSlipNm)
    
End Sub

Private Sub txtPart_Change()
    On Error GoTo ErrHandler
    
    Dim i%
    Dim bExist As Boolean
    
    'txtPart = UCase(txtPart)

    If Len(txtPart) = txtPart.MaxLength Then
        For i = 1 To giPartCnt
            lblPart.Caption = ""
            If gPartTable(i).sPartInit = txtPart Then
                lblPart.Caption = gPartTable(i).sPartName
                bExist = True
                Exit For
            End If
        Next
        
        txtPartGbn.SetFocus
        
        If bExist = False Then
            
            '-- 2001.07.06(osw)
            MsgBox "미등록 PART코드 입니다", vbOKOnly, Me.Caption
            '------------------
            
            lblPart.Caption = ""
            lblSlip.Caption = ""
            txtSlip = ""
            txtPartGbn = ""
                    
            '-- 2001.07.06(osw)
            txtPart.Text = ""
            txtPart.SetFocus
            '------------------
        End If
    ElseIf Len(txtPart) = 0 Then
        lblPart = ""
        lblSlip = ""
    End If
ErrHandler:
End Sub

Private Sub txtPart_Click()
    Call Txt_Highlight(txtPart)
End Sub

Private Sub txtPart_GotFocus()
    Call Txt_Highlight(txtPart)
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPartGbn.SetFocus
    End If
End Sub

Private Sub txtPart_Validate(Cancel As Boolean)
    Dim i%
    
    For i = 1 To giPartCnt
        If gPartTable(i).sPartInit = txtPart Then
            lblPart.Caption = gPartTable(i).sPartName
            Exit For
        End If
    Next
End Sub

Private Sub txtPartGbn_Change()
    'txtPartGbn = UCase(txtPartGbn)
    
    If Len(txtPartGbn) = txtPartGbn.MaxLength Then
        If txtPart = "" Then
            MsgBox "PART 코드를 입력한 후 PART 구분을 입력하여 주십시요!!"
            Exit Sub
        Else
            If iSpdClick = 1 Then
            Else
                'MsgBox "SpdClick 0"
                Call DisplaySlip
            End If
        End If
        
        iSpdClick = 0
        
        txtSlip.SetFocus
        
        lblSlip.Caption = txtPart & txtPartGbn
    ElseIf Len(txtPartGbn) = 0 Then
        lblSlip = ""
        txtSlip = ""
    End If
End Sub

Private Sub txtPartGbn_Click()
    Call Txt_Highlight(txtPartGbn)
End Sub

Private Sub txtPartGbn_GotFocus()
    Call Txt_Highlight(txtPartGbn)
End Sub

Private Sub txtPartGbn_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = 13 Then
'''        KeyCode = 0
'''        If Len(txtPartGbn) = txtPartGbn.MaxLength Then
'''            txtPartGbn = txtPartGbn
'''            lblSlip.Caption = txtPart & txtPartGbn
'''            txtSlip.SetFocus
'''        End If
'''
'''        Call CompareSpread
'''    End If
End Sub

Private Sub txtPartGbn_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSlip.SetFocus
    End If
End Sub

Private Sub txtPartGbn_LostFocus()
    If Len(txtPartGbn) < txtPartGbn.MaxLength Then
        txtPartGbn = Format$(txtPartGbn, "00")
    End If
End Sub

Private Sub txtPartGbn_Validate(Cancel As Boolean)
        
    If lblPart.Caption = "" Then
    Else
        If txtPartGbn = "" Then
        Else
            lblSlip.Caption = txtPart & txtPartGbn
            Call CompareSpread
        End If
    End If
End Sub

Private Sub txtSlip_Click()
    Call Txt_Highlight(txtSlip)
End Sub

Private Sub txtSlip_GotFocus()
    Call Txt_Highlight(txtSlip)
End Sub
