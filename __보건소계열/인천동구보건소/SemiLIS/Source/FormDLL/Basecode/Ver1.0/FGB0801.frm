VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0801 
   Caption         =   "기초자료 - MACHINE"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "FGB0801.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11280
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1170
      TabIndex        =   6
      Top             =   300
      Width           =   6795
      Begin VB.TextBox txtMachNm 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   795
         Width           =   3915
      End
      Begin VB.TextBox txtMachCd 
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
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "001"
         Top             =   360
         Width           =   480
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   345
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Machine 코드"
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
         Left            =   600
         TabIndex        =   8
         Top             =   795
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Machine 명"
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
      Begin VB.Label lblMachNm 
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
         Left            =   2700
         TabIndex        =   9
         Top             =   360
         Width           =   3375
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   9060
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2730
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
      Picture         =   "FGB0801.frx":030A
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   9060
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1710
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
      Picture         =   "FGB0801.frx":0BE4
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   9060
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3750
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
      Picture         =   "FGB0801.frx":14BE
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   1005
      Left            =   9060
      TabIndex        =   2
      Top             =   690
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
      Picture         =   "FGB0801.frx":1D98
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5010
      Left            =   1170
      TabIndex        =   10
      Top             =   1710
      Width           =   6780
      _Version        =   65536
      _ExtentX        =   11959
      _ExtentY        =   8837
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdBaseCode 
         Height          =   4605
         Left            =   630
         OleObjectBlob   =   "FGB0801.frx":2672
         TabIndex        =   11
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "FGB0801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCurSelRow%
Dim iSpdClick%

Private Sub DisplayMach()
    Dim i%
    
    Dim CMach As DCB0101
    
    Set CMach = New DCB0101
    
    CMach.Get_MACH Chr(CInt(txtMachCd))
    
    i = CMach.CurItemCnt
    
    If i = 0 Then
        txtMachNm = ""
        lblMachNm = ""
                
        Set CMach = Nothing
        Exit Sub
    End If
    
    lblMachNm = CMach.TotField01
    txtMachNm = lblMachNm
    
    Call FindCurSpreadRow(txtMachCd, txtMachNm)
    
    Set CMach = Nothing
End Sub

Private Sub DisplayInit()
    txtMachCd = ""
    txtMachNm = ""
    lblMachNm = ""
    
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

Private Sub FindCurSpreadRow(ByVal sField01 As String, ByVal sField02 As String)
    Dim vField01, vField02
    Dim iCurRow%
    Dim i%
    
    iCurRow = 0
    
    With spdBaseCode
        For i = 1 To .MaxRows
            Call .GetText(1, i, vField01)
            Call .GetText(2, i, vField02)
                        
            If CStr(vField01) = sField01 And CStr(vField02) = sField02 Then
                iCurRow = i
                Exit For
            End If
        Next
    
        If iCurRow = 0 Then
        Else
            Call spdReverse(spdBaseCode, -1, -1, iCurRow, iCurRow, RGB(255, 230, 230), 2)
        End If
    End With
End Sub

Private Sub cmdDelete_Click()
    Dim iRetVal%
    Dim CMach As DCB0101
    
    If txtMachCd = "" Then
        Exit Sub
    Else
        If lblMachNm.Caption = "" Then
            Exit Sub
        End If
        
        iRetVal = MsgBox("MACHINE 코드 : " & txtMachCd & vbCrLf & _
                "MACHINE 명 : " & lblMachNm.Caption & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "MACHINE 코드 삭제 확인")
    
        If iRetVal = 1 Then
            Set CMach = New DCB0101
            
            CMach.Delete_MACH Chr(CInt(txtMachCd))
                
            With spdBaseCode
                .Row = iCurSelRow
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            End With
            
            Set CMach = Nothing
        Else
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Dim CMach As DCB0101
    Dim i%, j%
    Dim vField01, vField02
    Dim iRegState%
    
    If txtMachCd = "" Then
        MsgBox "MACHINE 코드(숫자 3자리), MACHINE 명을 입력한 후, 등록하여 주십시요!!"
        Exit Sub
    End If
    
    If txtMachNm = "" Then
        MsgBox "MACHINE 코드(숫자 3자리), MACHINE 명을 입력한 후, 등록하여 주십시요!!"
        Exit Sub
    End If
    
    If LenH(txtMachNm) > 20 Then
        MsgBox "MACHINE 명은 한글 2 BYTE, 영문 1 BYTE 로 하여 20 BYTE 이하이어야 합니다!!"
        Exit Sub
    End If
    
    With spdBaseCode
        If .MaxRows > 0 Then
            For i = 1 To .MaxRows
                Call .GetText(1, i, vField01)
                Call .GetText(2, i, vField02)
                
                If CStr(vField01) = txtMachCd Then
                    If CStr(vField02) = txtMachNm Then
                        MsgBox "이미 존재하는 MACHINE 코드입니다!!"
                        Exit Sub
                    Else
                        iRegState = 1
                        iCurSelRow = i
                        Exit For
                    End If
                End If
            Next
        End If
    End With
    
    Set CMach = New DCB0101
        
    If iRegState = 0 Then
        
        If CInt(txtMachCd) > 250 Then
            MsgBox "001부터 250까지의 세자리의 숫자로 설정해 주십시요!!"
            Exit Sub
        End If
        
        CMach.Add_MACH Chr(CInt(txtMachCd)), txtMachNm
        
        If CMach.AdoErrNum = 0 Then
            '스프레드에 로를 추가
            spdBaseCode.MaxRows = spdBaseCode.MaxRows + 1
            Call spdBaseCode.SetText(1, spdBaseCode.MaxRows, txtMachCd & "")
            Call spdBaseCode.SetText(2, spdBaseCode.MaxRows, txtMachNm & "")
            lblMachNm.Caption = txtMachNm
        End If
    ElseIf iRegState = 1 Then
        CMach.Edit_MACH Chr(CInt(txtMachCd)), txtMachNm
        
        If CMach.AdoErrNum = 0 Then
            Call spdBaseCode.SetText(2, iCurSelRow, txtMachNm & "")
            lblMachNm.Caption = txtMachNm
        End If
    End If
    
    Set CMach = Nothing
End Sub

Private Sub cmdSearch_Click()
    Dim CMach As DCB0101
    Dim j%
    Dim i%
    Dim sField01$, sField02$
    
    Set CMach = New DCB0101
    
    CMach.Get_MACH
    
    i = CMach.CurItemCnt
    
    If i = 0 Then
        MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
        Set CMach = Nothing
        Exit Sub
    End If
    
    sField01 = CMach.TotField01
    sField02 = CMach.TotField02
    
    For j = 1 To i
        spdBaseCode.MaxRows = j
        Call spdBaseCode.SetText(1, j, Format(Asc(GetByOne(sField01, sField01)), "000") & "")
        Call spdBaseCode.SetText(2, j, GetByOne(sField02, sField02) & "")
    Next
    
    Set CMach = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
'        Case vbKeyF1:        Call cmdButtonPart_Click
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
    
    iSpdClick = 0
    
    Call DisplayInit
    Call cmdSearch_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
End Sub

Private Sub spdBaseCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vField01
    Dim vField02
    
    If Row = 0 Then
        Exit Sub
    End If
    
    iSpdClick = 1
    
    iCurSelRow = CInt(Row)
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
    
    txtMachCd = CStr(vField01)
    txtMachNm = CStr(vField02)
    lblMachNm = CStr(vField02)
    
End Sub

Private Sub txtMachCd_Change()
    If Len(txtMachCd) = txtMachCd.MaxLength Then
        If iSpdClick = 1 Then
        Else
            Call DisplayMach
        End If
        txtMachNm.SetFocus
    ElseIf Len(txtMachCd) = 0 Then
        txtMachNm = ""
        lblMachNm = ""
    End If
    
    iSpdClick = 0
End Sub

Private Sub txtMachCd_Click()
    Call Txt_Highlight(txtMachCd)
End Sub

Private Sub txtMachCd_GotFocus()
    Call Txt_Highlight(txtMachCd)
End Sub

Private Sub txtMachCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtMachNm.SetFocus
    End If
End Sub

Private Sub txtMachCd_LostFocus()
    If Len(txtMachCd) < txtMachCd.MaxLength Then
        txtMachCd = Format$(txtMachCd, "000")
    End If
End Sub

Private Sub txtMachNm_Click()
    Call Txt_Highlight(txtMachNm)
End Sub

Private Sub txtMachNm_GotFocus()
    Call Txt_Highlight(txtMachNm)
End Sub
