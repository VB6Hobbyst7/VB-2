VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0201 
   Caption         =   "기초자료 - SPECIMEN"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11190
   Icon            =   "FGB0201.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11190
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread spdBaseCode 
      Height          =   4605
      Left            =   420
      OleObjectBlob   =   "FGB0201.frx":030A
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2190
      Width           =   10365
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   420
      TabIndex        =   8
      Top             =   180
      Width           =   8100
      Begin VB.TextBox txtCd 
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
         Height          =   360
         IMEMode         =   8  '영문
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "XXX"
         Top             =   345
         Width           =   660
      End
      Begin VB.TextBox txtBriefNm 
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
         Height          =   360
         IMEMode         =   8  '영문
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "EDTA WHOLEBLOOD"
         Top             =   810
         Width           =   5355
      End
      Begin VB.TextBox txtFullNm 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   1305
         Width           =   5355
      End
      Begin Threed.SSPanel Panel3D7 
         Height          =   375
         Left            =   270
         TabIndex        =   9
         Top             =   810
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "검체명"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
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
         TabIndex        =   10
         Top             =   345
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "검체코드"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
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
         TabIndex        =   11
         Top             =   1305
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "검체원명"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         FloodColor      =   65535
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   8550
      TabIndex        =   5
      Top             =   1170
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
      Picture         =   "FGB0201.frx":13DC
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   9660
      TabIndex        =   4
      Top             =   180
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
      Picture         =   "FGB0201.frx":1CB6
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   9660
      TabIndex        =   6
      Top             =   1170
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
      Picture         =   "FGB0201.frx":2590
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   1005
      Left            =   8550
      TabIndex        =   3
      Top             =   180
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
      Picture         =   "FGB0201.frx":2E6A
   End
End
Attribute VB_Name = "FGB0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSpdClick%

Private Sub BaseCodeInit()
    Dim CSpecimen As DCB0101
    Dim i%
    Dim j%
    Dim sCd$
    Dim sBriefNm$
    Dim sFullNm$
    
    Set CSpecimen = New DCB0101
    
    CSpecimen.Get_SPC
    
    i = CSpecimen.CurItemCnt
    
    If i = 0 Then
        MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
        Set CSpecimen = Nothing
        Exit Sub
    End If
    
    sCd = CSpecimen.TotField01
    sBriefNm = CSpecimen.TotField02
    sFullNm = CSpecimen.TotField03
    
    For j = 1 To i
        spdBaseCode.MaxRows = j
        Call spdBaseCode.SetText(1, j, GetByOne(sCd, sCd) & "")
        Call spdBaseCode.SetText(2, j, GetByOne(sBriefNm, sBriefNm) & "")
        Call spdBaseCode.SetText(3, j, GetByOne(sFullNm, sFullNm) & "")
    Next
    
    Set CSpecimen = Nothing

End Sub

Private Function CompareSpread() As Integer
    Dim sComp1$
    'Dim sComp2$
    Dim i%
    Dim vField01, vField02, vField03
    
    CompareSpread = 0
    sComp1 = Left$(txtCd, 3)
    'sComp2 = Right$(lblSlip.Caption, 2)
    
    If spdBaseCode.MaxRows > 0 Then
        For i = 1 To spdBaseCode.MaxRows
            Call spdBaseCode.GetText(1, i, vField01)
            Call spdBaseCode.GetText(2, i, vField02)
            Call spdBaseCode.GetText(3, i, vField03)
            
            If sComp1 = vField01 Then
                
                txtBriefNm = CStr(vField02)
                txtFullNm = CStr(vField03)
                
                Call spdReverse(spdBaseCode, -1, -1, i, i, RGB(255, 230, 230), iSpdBackColorOption)
                CompareSpread = i
                Exit For
            End If
        Next
    End If
    
    If CompareSpread = 0 Then
        txtBriefNm = ""
        txtFullNm = ""
    End If
End Function

Private Sub DisplayInit()
    txtCd = ""
    txtBriefNm = ""
    txtFullNm = ""
    
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

Private Sub DisplaySpecimen()
    Dim i%
    Dim sField01$, sField02$, sField03$
    Dim CSpecimen As DCB0101
    
    Set CSpecimen = New DCB0101
    
    CSpecimen.Get_SPC txtCd
    
    i = CSpecimen.CurItemCnt
    
    If i = 0 Then
        txtBriefNm = ""
        txtFullNm = ""
        
        Set CSpecimen = Nothing
        Exit Sub
    End If
    
    sField01 = CSpecimen.TotField01
    sField02 = CSpecimen.TotField02
    sField03 = CSpecimen.TotField03
    
    txtBriefNm = GetByOne(sField02, sField02)
    txtFullNm = GetByOne(sField03, sField03)
    
    If spdBaseCode.MaxRows = 0 Then
        spdBaseCode.MaxRows = 1
        Call spdBaseCode.SetText(1, spdBaseCode.MaxRows, txtCd & "")
        Call spdBaseCode.SetText(2, spdBaseCode.MaxRows, txtBriefNm & "")
        Call spdBaseCode.SetText(3, spdBaseCode.MaxRows, txtFullNm & "")
    Else
        Call FindCurSpreadRow(txtCd, txtBriefNm, txtFullNm)
    End If
    
    Set CSpecimen = Nothing
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
                
                If IsNull(vField03) = True Then
                    vField03 = ""
                End If
                
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

Private Sub ShortKeyOrTabOrderInit()
    Me.KeyPreview = True
    
    txtCd.TabIndex = 0
    txtBriefNm.TabIndex = 1
    txtFullNm.TabIndex = 2
    cmdReg.TabIndex = 3
    cmdSearch.TabIndex = 4
    cmdDelete.TabIndex = 5
    cmdExit.TabIndex = 6
    
End Sub

Private Sub cmdDelete_Click()
    On Err GoTo ErrHandler
    
    Dim CSpecimen As DCB0101
    Dim iRetVal%
    
    If CompareSpread > 0 Then
        
        iRetVal = MsgBox("검체코드 : " & txtCd & vbCrLf & _
                "검체명 : " & txtBriefNm & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "검체코드 삭제 확인")
                 
        If iRetVal = 1 Then
            Set CSpecimen = New DCB0101
            
            CSpecimen.Delete_SPC txtCd
            
            Set CSpecimen = Nothing
            
            With spdBaseCode
                .Row = CompareSpread
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            End With
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
    
    Dim CSpecimen As DCB0101
    Dim vField01, vField02, vField03
    Dim i%
    Dim bMatch As Boolean
    
    If txtCd = "" Then
        MsgBox "검체코드 3자리의 숫자가 필요합니다!!"
        Exit Sub
    End If
    
    bMatch = False
    
    If spdBaseCode.MaxRows > 0 Then
        For i = 1 To spdBaseCode.MaxRows
            Call spdBaseCode.GetText(1, i, vField01)
            
            If txtCd = vField01 Then
                Call spdBaseCode.GetText(2, i, vField02)
                Call spdBaseCode.GetText(3, i, vField03)
                
                If vField02 = txtBriefNm And vField03 = txtFullNm Then
                    MsgBox "기존의 존재하는 데이터와 일치합니다", vbInformation, "확인"
                    Exit Sub
                Else
                    'EditItem - Slip명이 틀려진 경우
                    Set CSpecimen = New DCB0101
                    
                    CSpecimen.Edit_SPC txtCd, Left(txtBriefNm, 30), Left(txtFullNm, 30), 3
                    
                    If CSpecimen.AdoErrNum = 0 Then
                        '화면에 반영
                        With spdBaseCode
                            Call .SetText(1, i, txtCd & "")
                            Call .SetText(2, i, txtBriefNm & "")
                            Call .SetText(3, i, txtFullNm & "")
                        End With
                        
                        txtCd.SetFocus
                    End If
                    
                    Set CSpecimen = Nothing
                    
                End If
                
                bMatch = True
                
                Exit For
            Else
                bMatch = False
            End If
        Next
    End If
    
    If bMatch = False Then
        Set CSpecimen = New DCB0101
        
        CSpecimen.Add_SPC txtCd, Left(txtBriefNm, 30), Left(txtFullNm, 30)
        
        If CSpecimen.AdoErrNum = 0 Then
            '화면에 반영
            With spdBaseCode
                 .MaxRows = .MaxRows + 1
                Call .SetText(1, spdBaseCode.MaxRows, txtCd & "")
                Call .SetText(2, spdBaseCode.MaxRows, txtBriefNm & "")
                Call .SetText(3, spdBaseCode.MaxRows, txtFullNm & "")
            End With
            
            txtCd.SetFocus
        End If
        
        Set CSpecimen = Nothing
        
    End If
    
'<---------- 다시 화면에 뿌리는데 시간이 소요되는 문제 ------------------>
    '바뀐내용 다시 화면에
    'spdBaseCode.MaxRows = 0
    
    'Call BaseCodeInit
'<------------------------------------------------------------------------->
    
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
        'Case vbKeyF1:        Call cmdButtonPart_Click
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
    Dim vField03
    Dim vTmp
    
    If Row = 0 Then
        Exit Sub
    End If
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), iSpdBackColorOption)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
    Call spdBaseCode.GetText(3, Row, vField03)
    
    iSpdClick = 1
    
    txtCd = CStr(vField01)
    txtBriefNm = CStr(vField02)
    txtFullNm = CStr(vField03)
    
End Sub

Private Sub txtBriefNm_Change()
    If Len(txtBriefNm) = 0 Then
        txtFullNm = ""
    End If
End Sub

Private Sub txtBriefNm_Click()
    Call Txt_Highlight(txtBriefNm)
End Sub

Private Sub txtBriefNm_GotFocus()
    Call Txt_Highlight(txtBriefNm)
End Sub

Private Sub txtBriefNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(txtCd) = txtCd.MaxLength Then
            If txtFullNm = "" Then
                txtFullNm = txtBriefNm
            End If
            txtFullNm.SetFocus
        End If
    End If
End Sub

Private Sub txtBriefNm_Validate(Cancel As Boolean)
    If Len(txtCd) = txtCd.MaxLength Then
        If txtFullNm = "" Then
            txtFullNm = txtBriefNm
        End If
        txtFullNm.SetFocus
    End If
End Sub

Private Sub txtCd_Change()
    On Error GoTo ErrHandler
    'txtCd = UCase(txtCd)
    
    If Len(txtCd) = txtCd.MaxLength Then
        
        If iSpdClick = 1 Then
        Else
            Call DisplaySpecimen
        End If
        
        iSpdClick = 0
        
        txtBriefNm.SetFocus
    ElseIf Len(txtCd) = 0 Then
        txtBriefNm = ""
        txtFullNm = ""
    End If
    
ErrHandler:
    
End Sub

Private Sub txtCd_Click()
    Call Txt_Highlight(txtCd)
End Sub

Private Sub txtCd_GotFocus()
    Call Txt_Highlight(txtCd)
End Sub

Private Sub txtCD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtBriefNm.SetFocus
    End If
End Sub

Private Sub txtCd_LostFocus()
    If Len(txtCd) < txtCd.MaxLength Then
        txtCd = Format$(txtCd, "000")
    End If
End Sub

Private Sub txtCd_Validate(Cancel As Boolean)
    Call CompareSpread
End Sub

Private Sub txtFullNm_Click()
    Call Txt_Highlight(txtFullNm)
End Sub

Private Sub txtFullNm_GotFocus()
    Call Txt_Highlight(txtFullNm)
End Sub
