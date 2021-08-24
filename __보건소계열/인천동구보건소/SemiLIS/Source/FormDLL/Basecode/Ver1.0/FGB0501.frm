VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0501 
   Caption         =   "기초자료 - DEPT"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   Icon            =   "FGB0501.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11160
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread spdBaseCode 
      Height          =   4605
      Left            =   1110
      OleObjectBlob   =   "FGB0501.frx":030A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2010
      Width           =   6675
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1140
      TabIndex        =   7
      Top             =   270
      Width           =   6600
      Begin VB.TextBox txtDept 
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
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "01"
         Top             =   360
         Width           =   450
      End
      Begin VB.TextBox txtDeptNm 
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
         TabIndex        =   1
         Top             =   795
         Width           =   4425
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   375
         Left            =   270
         TabIndex        =   8
         Top             =   345
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "DEPT 코드"
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
         TabIndex        =   9
         Top             =   795
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "DEPT 명"
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
      Begin VB.Label lblDept 
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
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   3945
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   8910
      TabIndex        =   4
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
      Picture         =   "FGB0501.frx":13AA
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   8910
      TabIndex        =   3
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
      Picture         =   "FGB0501.frx":1C84
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   8910
      TabIndex        =   5
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
      Picture         =   "FGB0501.frx":255E
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   1005
      Left            =   8910
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
      Picture         =   "FGB0501.frx":2E38
   End
End
Attribute VB_Name = "FGB0501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCurSelRow%
Dim iSpdClick%

Private Sub DisplayDept()
    Dim i%
    
    Dim CDept As DCB0101
    
    Set CDept = New DCB0101
    
    CDept.Get_DEPT txtDept
    
    i = CDept.CurItemCnt
    
    If i = 0 Then
        txtDeptNm = ""
        lblDept.Caption = ""
                
        Set CDept = Nothing
        Exit Sub
    End If
    
    lblDept.Caption = CDept.TotField01
    txtDeptNm = lblDept
    
    Call FindCurSpreadRow(txtDept, txtDeptNm)
    
    Set CDept = Nothing
    
End Sub

Private Sub DisplayInit()
    txtDept = ""
    txtDeptNm = ""
        
    lblDept = ""
    
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

Private Sub ShortKeyOrTabOrderInit()
    Me.KeyPreview = True
    
    txtDept.TabIndex = 0
    txtDeptNm.TabIndex = 1
    
    cmdReg.TabIndex = 2
    cmdSearch.TabIndex = 3
    cmdDelete.TabIndex = 4
    cmdExit.TabIndex = 5
    
End Sub

Private Sub cmdDelete_Click()
    Dim iRetVal%
    Dim CDept As DCB0101
    
    If txtDept = "" Then
        Exit Sub
    Else
        If lblDept.Caption = "" Then
            Exit Sub
        End If
        
        iRetVal = MsgBox("DEPT 코드 : " & txtDept & vbCrLf & _
                "DEPT 명 : " & lblDept.Caption & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "DEPT 코드 삭제 확인")
    
        If iRetVal = 1 Then
            Set CDept = New DCB0101
            
            CDept.Delete_DEPT txtDept
                
            With spdBaseCode
                .Row = iCurSelRow
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            End With
            
            Set CDept = Nothing
        Else
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Dim CDept As DCB0101
    Dim i%, j%
    Dim vField01, vField02
    Dim iRegState%
    
    If txtDept = "" Then
        MsgBox "DEPT 코드(숫자 2자리), DEPT 명을 입력한 후, 등록하여 주십시요!!"
        Exit Sub
    End If
    
    If txtDeptNm = "" Then
        MsgBox "DEPT 코드(숫자 2자리), DEPT 명을 입력한 후, 등록하여 주십시요!!"
        Exit Sub
    End If
    
    With spdBaseCode
        If .MaxRows > 0 Then
            For i = 1 To .MaxRows
                Call .GetText(1, i, vField01)
                Call .GetText(2, i, vField02)
                
                If CStr(vField01) = txtDept Then
                    If CStr(vField02) = txtDeptNm Then
                        MsgBox "이미 존재하는 DEPT 코드입니다!!"
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
    
    Set CDept = New DCB0101
        
    If iRegState = 0 Then
        CDept.Add_DEPT txtDept, txtDeptNm
        
        If CDept.AdoErrNum = 0 Then
            '스프레드에 로를 추가
            spdBaseCode.MaxRows = spdBaseCode.MaxRows + 1
            Call spdBaseCode.SetText(1, spdBaseCode.MaxRows, txtDept & "")
            Call spdBaseCode.SetText(2, spdBaseCode.MaxRows, txtDeptNm & "")
            lblDept.Caption = txtDeptNm
        End If
    ElseIf iRegState = 1 Then
        CDept.Edit_DEPT txtDept, txtDeptNm
        
        If CDept.AdoErrNum = 0 Then
            Call spdBaseCode.SetText(2, iCurSelRow, txtDeptNm & "")
            lblDept.Caption = txtDeptNm
        End If
    End If
    
    Set CDept = Nothing
    
End Sub

Private Sub cmdSearch_Click()
    Dim i%, j%
    Dim sField01$, sField02$
    Dim CDept As DCB0101
    
    Set CDept = New DCB0101
    
    CDept.Get_DEPT
    
    i = CDept.CurItemCnt
    
    If i = 0 Then
        MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
        Set CDept = Nothing
        Exit Sub
    End If
    
    sField01 = CDept.TotField01
    sField02 = CDept.TotField02
    
    For j = 1 To i
        spdBaseCode.MaxRows = j
        Call spdBaseCode.SetText(1, j, GetByOne(sField01, sField01) & "")
        Call spdBaseCode.SetText(2, j, GetByOne(sField02, sField02) & "")
    Next
    
    Set CDept = Nothing
    
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
    Call ShortKeyOrTabOrderInit
    
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
    
    iCurSelRow = CInt(Row)
    
    iSpdClick = 1
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
    
    txtDept = CStr(vField01)
    txtDeptNm = CStr(vField02)
    lblDept = CStr(vField02)
    
End Sub

Private Sub txtDept_Change()
    On Error GoTo ErrHandler
    
    If Len(txtDept) = txtDept.MaxLength Then
        If txtDept = "" Then
            MsgBox "DEPT 코드를 입력한 후 DEPT 명을 입력하여 주십시요!!"
            Exit Sub
        Else
            If iSpdClick = 1 Then
            Else
                Call DisplayDept
            End If
        End If
        
        txtDeptNm.SetFocus
    ElseIf Len(txtDept) = 0 Then
        txtDeptNm = ""
        lblDept.Caption = ""
    End If
    
    iSpdClick = 0
ErrHandler:
End Sub

Private Sub txtDept_Click()
    Call Txt_Highlight(txtDept)
End Sub

Private Sub txtDept_GotFocus()
    Call Txt_Highlight(txtDept)
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDeptNm.SetFocus
    End If
End Sub

Private Sub txtDept_LostFocus()
    If Len(txtDept) < txtDept.MaxLength Then
        txtDept = Format$(txtDept, "00")
    End If
End Sub

Private Sub txtDeptNm_Click()
    Call Txt_Highlight(txtDeptNm)
End Sub

Private Sub txtDeptNm_GotFocus()
    Call Txt_Highlight(txtDeptNm)
End Sub
