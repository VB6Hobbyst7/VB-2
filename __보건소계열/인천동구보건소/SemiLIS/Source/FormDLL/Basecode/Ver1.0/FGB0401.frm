VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0401 
   Caption         =   "기초자료 - ROUTINE"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   Icon            =   "FGB0401.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11775
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   90
      TabIndex        =   14
      Top             =   3360
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   7064
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
      Begin FPSpread.vaSpread spdTestItem 
         Height          =   3735
         Left            =   210
         OleObjectBlob   =   "FGB0401.frx":030A
         TabIndex        =   15
         Top             =   180
         Width           =   7425
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   930
         Left            =   8820
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   1640
         _StockProps     =   78
         Caption         =   "검사항목선택 F9"
         ForeColor       =   8388736
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
         Picture         =   "FGB0401.frx":1477
      End
   End
   Begin VB.Frame fraRoutine2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   90
      TabIndex        =   8
      Top             =   -60
      Width           =   11610
      Begin FPSpread.vaSpread spdBaseCode 
         Height          =   1785
         Left            =   210
         OleObjectBlob   =   "FGB0401.frx":1D51
         TabIndex        =   12
         Top             =   1530
         Width           =   8745
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
         Left            =   1770
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "H"
         Top             =   255
         Width           =   300
      End
      Begin VB.TextBox txtRtnNm 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   8  '영문
         Left            =   1770
         TabIndex        =   2
         Top             =   1065
         Width           =   5955
      End
      Begin VB.TextBox txtRtn 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   8  '영문
         Left            =   1770
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "XXX"
         ToolTipText     =   "Routine 검사코드"
         Top             =   660
         Width           =   585
      End
      Begin Threed.SSPanel pnlRtnCd 
         Height          =   345
         Left            =   210
         TabIndex        =   9
         Top             =   660
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Routine 순번"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdButtonRtn 
         Height          =   330
         Left            =   2370
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Routine 검사코드 Help"
         Top             =   660
         Width           =   270
         _Version        =   65536
         _ExtentX        =   476
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FGB0401.frx":2089
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   945
         Left            =   10260
         TabIndex        =   4
         Top             =   1410
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "조회 F3"
         ForeColor       =   12582912
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
         Picture         =   "FGB0401.frx":21AB
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   945
         Left            =   9120
         TabIndex        =   5
         Top             =   2370
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   1667
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
         Picture         =   "FGB0401.frx":2A85
      End
      Begin Threed.SSPanel pnlRoutineNm 
         Height          =   345
         Left            =   210
         TabIndex        =   13
         Top             =   1080
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Routine 명"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdReg 
         Height          =   945
         Left            =   9120
         TabIndex        =   3
         Top             =   1410
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   1667
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
         Picture         =   "FGB0401.frx":335F
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   945
         Left            =   10260
         TabIndex        =   6
         Top             =   2370
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "종료Esc"
         ForeColor       =   128
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
         Picture         =   "FGB0401.frx":3C39
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   345
         Left            =   210
         TabIndex        =   16
         Top             =   240
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "PART 코드"
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
      Begin Threed.SSCommand cmdButtonPart 
         Height          =   330
         Left            =   2070
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
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
         Picture         =   "FGB0401.frx":4513
      End
      Begin VB.Label lblPart 
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
         Left            =   2400
         TabIndex        =   18
         Top             =   255
         Width           =   3885
      End
      Begin VB.Label lblRtnNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Left            =   2700
         TabIndex        =   11
         Top             =   660
         Width           =   6945
      End
   End
End
Attribute VB_Name = "FGB0401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCurSelRow%
Dim iSpdClick1%
Dim iSpdClick2%

Private Sub DisplayInit()
    txtPart = ""
    txtRtn = ""
    txtRtnNm = ""
    
    lblPart = ""
    lblRtnNm = ""
    
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
    
      With spdTestItem
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
        .Col = 2
        .Col2 = 4
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        .MaxRows = 0
    End With
End Sub

Private Sub DisplayRtn()
    Dim i%
    
    Dim CRoutine As DCB0101
    
    Set CRoutine = New DCB0101
    
    CRoutine.Get_RTN 3, txtPart, txtRtn
    
    i = CRoutine.CurItemCnt
    
    If i = 0 Then
        txtRtnNm = ""
        lblRtnNm.Caption = ""
        spdTestItem.MaxRows = 0
        
        Set CRoutine = Nothing
        Exit Sub
    End If
    
    lblRtnNm.Caption = CRoutine.TotField01
    txtRtnNm = lblRtnNm
    
    Call FindCurSpreadRow(txtPart, txtRtn, txtRtnNm)
    
    Set CRoutine = Nothing
    
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
                        
            If Left$(CStr(vField01), 1) = sField01 And Right$(CStr(vField01), 3) = sField02 And CStr(vField02) = sField03 Then
                iCurRow = i
                Call spdBaseCode_Click(1, i)
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

Private Sub AddAfterDelete()

    Dim i%, j%
    Dim vTestCd, vTestGbn, vChk, vRtnCd
    Dim sMulti$
    Dim iRepeatCnt%
    Dim iCurRow%
    Dim sField01$, sField02$, sField03$, sField04$, sField05$, sField06$, sField07$, sField08$
    
    Dim CRoutine As DCB0101
    
    Set CRoutine = New DCB0101
    
    CRoutine.Delete_RTN txtPart, txtRtn     'DELETE WITH PRIMARYKEY
    
    iRepeatCnt = 0
    iCurRow = 0
    
    If CRoutine.AdoErrNum = 0 Then
        With spdTestItem
            For i = 1 To .MaxRows
                Call .GetText(1, i, vChk)
                Call .GetText(2, i, vTestCd)
                Call .GetText(3, i, vTestGbn)
                
                If vChk = "1" Then
                    iRepeatCnt = iRepeatCnt + 1
                    sField01 = sField01 & txtPart & "|"
                    sField02 = sField02 & txtRtn & "|"
                    sField03 = sField03 & Mid$(CStr(vTestCd), 2, 2) & "|"
                    sField04 = sField04 & Mid$(CStr(vTestCd), 4, 3) & "|"
                    sField05 = sField05 & Mid$(CStr(vTestCd), 7, 3) & "|"
                    sField06 = sField06 & fJudgeSUBMCD(CStr(vTestGbn)) & "|"
                    sField07 = sField07 & txtRtnNm & "|"
                    
                    .BlockMode = True
                    .Col = -1
                    .Col2 = -1
                    .Row = i
                    .Row2 = i
                    .BackColor = 연하늘
                    .BlockMode = False
                Else
                    If vChk = "" Or vChk = "0" Then
                        If i < .MaxRows Then
                            .Row = i
                            .Action = SS_ACTION_DELETE_ROW
                            .MaxRows = .MaxRows - 1
                            i = i - 1
                        End If
                    End If
                End If
            Next
        End With
        
        If iRepeatCnt = 0 Then
        Else
            CRoutine.Add_RTN sField01, sField02, sField03, sField04, _
                        sField05, sField06, sField07, iRepeatCnt
        End If
        
        If CRoutine.AdoErrNum = 0 Then
            With spdBaseCode
                For i = 1 To .MaxRows
                    Call .GetText(1, i, vRtnCd)
                    
                    If CStr(vRtnCd) = txtPart & txtRtn Then
                        iCurRow = i
                        Exit For
                    End If
                Next
                
                If iCurRow = 0 Then
                    .MaxRows = .MaxRows + 1
                    Call .SetText(1, .MaxRows, txtPart & txtRtn & "")
                    Call .SetText(2, .MaxRows, txtRtnNm & "")
                Else
                    Call .SetText(1, i, txtPart & txtRtn & "")
                    Call .SetText(2, i, txtRtnNm & "")
                End If
            End With
        End If
    Else
    End If
    
    Set CRoutine = Nothing
End Sub

Private Sub EditOnlyRtnNm()
    Dim i%
    Dim vRtnCd
    Dim CRoutine As DCB0101
    
    Set CRoutine = New DCB0101
    
    CRoutine.Edit_RTNNM txtPart, txtRtn, txtRtnNm, 0
    
    For i = 1 To spdBaseCode.MaxRows
        Call spdBaseCode.GetText(1, i, vRtnCd)
        
        If Left$(CStr(vRtnCd), 1) = txtPart And Right$(CStr(vRtnCd), 3) = txtRtn Then
            Call spdBaseCode.SetText(2, i, txtRtnNm & "")
            lblRtnNm.Caption = txtRtnNm
        End If
    Next
    
    If CRoutine.AdoErrNum = 0 Then
        ViewMsg "등록작업이 성공적으로 수행되었습니다..."
    Else
        ViewMsg "등록작업 중 에러가 발생했습니다. 에러번호( " & CRoutine.AdoErrNum & " ) 을 참조하세요..."
    End If
    
    Set CRoutine = Nothing
End Sub

Private Sub ShortKeyOrTabOrderInit()
    Me.KeyPreview = True
    
    txtPart.TabIndex = 0
    txtRtn.TabIndex = 1
    txtRtnNm.TabIndex = 2
    cmdReg.TabIndex = 3
    cmdSearch.TabIndex = 4
    cmdDelete.TabIndex = 5
    cmdExit.TabIndex = 6
    
End Sub

Private Sub cmdButtonPart_Click()
    Dim i%
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(giPartCnt) As CodeTBL
    
    For i = 1 To giPartCnt
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = gPartTable(i).sPartInit
        gCodeHlpTable(i).sCodeNm = gPartTable(i).sPartName
    Next
    
    giCodeHlpCnt = giPartCnt
    
    hWndCd = txtPart.hwnd
    
    FSB0101.Left = 2600
    FSB0101.Top = 1500
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdButtonRtn_Click()
    Dim i%, j%
    Dim CRoutine As DCB0101
    Dim sField01$, sField02$
    
    Set CRoutine = New DCB0101
    
    If txtPart = "" Then
        MsgBox "PART 코드를 입력하여야 그 PART 아래의 ROUTINE 코드를 볼 수 있습니다!!"
        Exit Sub
    Else
        CRoutine.Get_RTN 0, txtPart, ""     'SELECT WITH PARTCD
    End If
    
    i = CRoutine.CurItemCnt
    
    If i = 0 Then
        MsgBox "PART 코드 - " & txtPart & " 에 설정된 ROUTINE 코드가 없습니다!!"
        Set CRoutine = Nothing
        Exit Sub
    End If
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(i) As CodeTBL
    
    sField01 = CRoutine.TotField01
    sField02 = CRoutine.TotField02
    
    For j = 1 To i
        gCodeHlpTable(j).sSeq = Format$(j, "00000")
        gCodeHlpTable(j).sCode = GetByOne(sField01, sField01)
        gCodeHlpTable(j).sCodeNm = GetByOne(sField02, sField02)
    Next

    giCodeHlpCnt = i
    
    hWndCd = txtRtn.hwnd
    
    FSB0101.Left = 2900
    FSB0101.Top = 2000
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdDelete_Click()
    Dim iRetVal As Integer
    Dim CRoutine As DCB0101
    
    If txtPart = "" Then
        Exit Sub
    End If
    
    If txtRtn = "" Then
        Exit Sub
    End If
    
    If lblRtnNm.Caption = "" Then
        Exit Sub
    End If
    
    iRetVal = MsgBox("ROUTINE 코드 : " & txtPart & txtRtn & vbCrLf & _
                "ROUTINE 명 : " & lblRtnNm.Caption & " 을(를) 삭제하시겠습니까?", _
                 vbOKCancel, "ROUTINE 코드 삭제 확인")
    
    If iRetVal = 1 Then
        Set CRoutine = New DCB0101
        
        CRoutine.Delete_RTN txtPart, txtRtn     'DELETE WITH PRIMARYKEY
            
        With spdBaseCode
            .Row = iCurSelRow
            .Action = SS_ACTION_DELETE_ROW
            .MaxRows = .MaxRows - 1
        End With
        
        spdTestItem.MaxRows = 0
        
        Set CRoutine = Nothing
    Else
    End If
                     
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Dim i%
    Dim vRtnCd, vRtnNm, vTestCd, vTestGbn
    Dim iNewRtn%
    Dim vChk
    Dim sMulti$
    
    If Len(txtPart) = 0 Then
        MsgBox "등록을 하려면 PART 코드(H, C, S, U, M, O ...), " & vbCrLf & _
               "ROUTINE 순번(3자리의 숫자), " & vbCrLf & _
               "ROUTINE 명이 필요합니다!!"
        Exit Sub
    Else
        If Len(txtRtn) <> 3 Or IsNumeric(txtRtn) = False Then
            MsgBox "등록을 하려면 PART 코드(H, C, S, U, M, O ...), " & vbCrLf & _
               "ROUTINE 순번(3자리의 숫자), " & vbCrLf & _
               "ROUTINE 명이 필요합니다!!"
            Exit Sub
        End If
        
        If txtRtnNm = "" Then
            MsgBox "등록을 하려면 PART 코드(H, C, S, U, M, O ...), " & vbCrLf & _
               "ROUTINE 순번(3자리의 숫자), " & vbCrLf & _
               "ROUTINE 명이 필요합니다!!"
            Exit Sub
        End If
    End If
    
    iNewRtn = 0     '바뀌거나 추가된 것이 없는 상태
    
    With spdTestItem
        If .MaxRows > 0 Then    '항목 추가, 삭제 check
            For i = 1 To .MaxRows
                Call .GetText(1, i, vChk)

                If vChk = "1" Then
                    .Row = i
                    If .BackColor = 연노랑 Then     '추가
                        iNewRtn = 1
                        Exit For
                    End If
                ElseIf vChk = "" Or vChk = "0" Then
                    .Row = i
                    If .BackColor = 연하늘 Then     '삭제
                        iNewRtn = 1
                        Exit For
                    End If
                End If
            Next
        End If
    End With
    
    If iNewRtn = 0 Then
        With spdBaseCode
            If .MaxRows > 0 Then
                For i = 1 To .MaxRows
                    Call .GetText(1, i, vRtnCd)
                    Call .GetText(2, i, vRtnNm)
                    
                    If txtPart & txtRtn = CStr(vRtnCd) Then
                        If txtRtnNm <> CStr(vRtnNm) Then
                            iNewRtn = 2     '이름만 바뀜
                            Exit For
                        End If
                    End If
                Next
            End If
        End With
    End If
    
    If iNewRtn = 0 Then
    ElseIf iNewRtn = 1 Then
        '추가(DELETE 후) 필요
        Call AddAfterDelete
    ElseIf iNewRtn = 2 Then
        '루틴명만 수정
        Call EditOnlyRtnNm
    End If
    
End Sub

Private Sub cmdSearch_Click()
    Dim i%
    Dim j%
    Dim sField01$, sField02$, sField03$
    Dim vRtnCd
    Dim iDisplayRow%
    
    Dim CRoutine As DCB0101
    
    Set CRoutine = New DCB0101
    
    If txtPart = "" Then
        CRoutine.Get_RTN 1, "", ""
        
        i = CRoutine.CurItemCnt
        
        If i = 0 Then
            MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
            Set CRoutine = Nothing
            Exit Sub
        End If
        
        sField01 = CRoutine.TotField01
        sField02 = CRoutine.TotField02
        sField03 = CRoutine.TotField03
                
        spdTestItem.MaxRows = 0
        
        For j = 1 To i
            spdBaseCode.MaxRows = j
            Call spdBaseCode.SetText(1, j, GetByOne(sField01, sField01) & GetByOne(sField02, sField02) & "")
            Call spdBaseCode.SetText(2, j, GetByOne(sField03, sField03) & "")
        Next
        
        Set CRoutine = Nothing
    Else
        If txtRtn = "" Then
            CRoutine.Get_RTN 0, txtPart, ""
    
            i = CRoutine.CurItemCnt
            
            If i = 0 Then
                MsgBox "관련 항목이 등록되어 있지 않습니다!!"
                txtRtn = ""
                txtRtnNm = ""
                lblRtnNm.Caption = ""
                Set CRoutine = Nothing
                Exit Sub
            End If
            
            sField01 = CRoutine.TotField01  'RoutineCd
            sField02 = CRoutine.TotField02  'RoutineNm
            
            spdTestItem.MaxRows = 0
            
            For j = 1 To i
                spdBaseCode.MaxRows = j
                Call spdBaseCode.SetText(1, j, txtPart & GetByOne(sField01, sField01) & "")
                Call spdBaseCode.SetText(2, j, GetByOne(sField02, sField02) & "")
            Next
            
            Set CRoutine = Nothing
        Else
            CRoutine.Get_RTN 3, txtPart, txtRtn
            
            i = CRoutine.CurItemCnt
            
            If i = 0 Then
                MsgBox "관련 항목이 등록되어 있지 않습니다!!"
                Set CRoutine = Nothing
                Exit Sub
            End If
            
            sField01 = CRoutine.TotField01
            
            iDisplayRow = 0
            
            For j = 1 To spdTestItem.MaxRows
                Call spdBaseCode.GetText(1, j, vRtnCd)
                
                If CStr(vRtnCd) = txtPart & txtRtn Then
                    iDisplayRow = j
                    Exit For
                End If
            Next
            
            If iDisplayRow = 0 Then
                spdTestItem.MaxRows = 0
                sField01 = CRoutine.TotField01  'RoutineName
            
                spdBaseCode.MaxRows = 1
                Call spdBaseCode.SetText(1, 1, txtPart & txtRtn & "")
                Call spdBaseCode.SetText(2, 1, sField01 & "")
            Else
                '이미 Spread에 해당자료가 있을 때
                Call spdReverse(spdBaseCode, -1, -1, iDisplayRow, iDisplayRow, RGB(255, 230, 230), 2)
            End If
            
            Set CRoutine = Nothing
        End If
    End If
    
End Sub

Private Sub cmdSelect_Click()
    
    Dim CTestItem As DCB0101
    Dim j%
    Dim i%
    Dim sTot01$, sTot02$, sTot03$, sTot04$, sTot05$
    Dim sTmp1$
    Dim sTmp2$
    Dim vTmp
    
    If txtRtn = "" Or txtRtnNm = "" Then
        MsgBox "먼저 Routine 코드와 Routine 명을 입력한 후, 검사항목을 선택하여 주십시요!!"
        Exit Sub
    End If
    
    Set CTestItem = New DCB0101
    
    If spdTestItem.MaxRows > 0 Then
        Call spdTestItem.GetText(2, 1, vTmp)
        CTestItem.Get_TESTITEM 13, Left$(CStr(vTmp), 1), Mid$(CStr(vTmp), 2, 2), Mid$(CStr(vTmp), 4, 3)
    Else
        CTestItem.Get_TESTITEM 13, txtPart.Text, "", ""
        '-- CTestItem.Get_TESTITEM 13, fCurUserPartCd, Right$(fCurUserSlipCd, 2), fCurUserSpcCd
    End If
        
    j = CTestItem.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    With CTestItem
        sTot01 = .TotField01    'PartGbn
        sTot02 = .TotField02    'SpecimenCd
        sTot03 = .TotField03    'TestItemSeq
        sTot04 = .TotField04    'SUBMCD
        sTot05 = .TotField05    'TESTITEMNM
    End With
    
    Set CTestItem = Nothing

    ReDim gCodeHlpTable(j) As CodeTBL
    
    giCodeHlpCnt = 0
    
    For i = 1 To j
        sTmp1 = GetByOne(sTot04, sTot04)
        
        If sTmp1 = "NNNN" Then
            giCodeHlpCnt = giCodeHlpCnt + 1
            gCodeHlpTable(giCodeHlpCnt).sGbn = "N"
            gCodeHlpTable(giCodeHlpCnt).sSeq = Format$(giCodeHlpCnt, "00000")
            gCodeHlpTable(giCodeHlpCnt).sCode = txtPart & GetByOne(sTot01, sTot01) & _
                                GetByOne(sTot02, sTot02) & GetByOne(sTot03, sTot03)
        
            gCodeHlpTable(giCodeHlpCnt).sCodeNm = GetByOne(sTot05, sTot05)
            
        ElseIf IsNumeric(Left$(sTmp1, 2)) = True And Left$(sTmp1, 2) = "00" Then
            'SUB 원검사만 추가
            giCodeHlpCnt = giCodeHlpCnt + 1
            gCodeHlpTable(giCodeHlpCnt).sGbn = "S" & Left$(sTmp1, 2)
            gCodeHlpTable(giCodeHlpCnt).sSeq = Format$(giCodeHlpCnt, "00000")
            gCodeHlpTable(giCodeHlpCnt).sCode = txtPart & GetByOne(sTot01, sTot01) & _
                                GetByOne(sTot02, sTot02) & GetByOne(sTot03, sTot03)
        
            gCodeHlpTable(giCodeHlpCnt).sCodeNm = GetByOne(sTot05, sTot05)
            
        Else
            'SUB 원검사 이외의 제외
            Call GetByOne(sTot01, sTot01)
            Call GetByOne(sTot02, sTot02)
            Call GetByOne(sTot03, sTot03)
            Call GetByOne(sTot05, sTot05)
            
        End If
    Next
    
    giCodeHlpMode = 1
    
    Set gCallObject = FGB0401.spdTestItem
    
    FSB0201.Left = 6000
    FSB0201.Top = 1000
    
    Load FSB0201
    FSB0201.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        'Case vbKeyF1:        Call cmdButtonPart_Click
        Case vbKeyF2:        Call cmdReg_Click
        Case vbKeyF3:        Call cmdSearch_Click
        Case vbKeyF4:        Call cmdDelete_Click
        Case vbKeyF9:        Call cmdSelect_Click
        Case vbKeyEscape:    Call cmdExit_Click
    End Select
End Sub

Private Sub Form_Load()
    
    If Me.StartUpPosition = 2 Then
    Else
        Me.Left = 0
        Me.Top = 0
        Me.Width = 11900
        Me.Height = 7900
    End If
    
    iSpdClick1 = 0
    iSpdClick2 = 0
    
    Call DisplayInit
    Call ShortKeyOrTabOrderInit
    'Call BaseCodeInit
    
    Call cmdSearch_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
    Set gCallObject = Nothing
End Sub

Private Sub spdBaseCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vField01
    Dim vField02
    Dim i%, j%
    Dim sField01$, sField02$, sField03$
    Dim CRoutine As DCB0101
    Dim sPartGbn$, sSpecimen$, sTestSeq$, sSubMCd$, sTestNm$
    Dim sTmp1$
    
    If Row = 0 Then
        Exit Sub
    End If
    
    iCurSelRow = CInt(Row)
    iSpdClick1 = 1
    iSpdClick2 = 1
    
    spdTestItem.MaxRows = 0
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    Call spdReverse(spdTestItem, -1, -1, -1, -1, 연하늘, 2)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
    
    sField01 = Left$(CStr(vField01), 1)
    sField02 = Mid$(CStr(vField01), 2, 3)
    
    txtPart = sField01
    txtRtn = sField02
    lblRtnNm.Caption = CStr(vField02)
    txtRtnNm = CStr(vField02)
    
    Set CRoutine = New DCB0101
    
    CRoutine.Get_RTN 2, sField01, sField02  'TESTITEM과 JOIN

    i = CRoutine.CurItemCnt

    If i = 0 Then
        MsgBox "아직 기초자료에 어떤 항목도 등록되어 있지 않습니다!!"
        Set CRoutine = Nothing
        Exit Sub
    End If
    
    sPartGbn = CRoutine.TotField01
    sSpecimen = CRoutine.TotField02
    sTestSeq = CRoutine.TotField03
    sSubMCd = CRoutine.TotField04
    sTestNm = CRoutine.TotField05
    
    For j = 1 To i
        With spdTestItem
            .MaxRows = j
            Call .SetText(1, j, "1")
            Call .SetText(2, j, sField01 & GetByOne(sPartGbn, sPartGbn) & GetByOne(sSpecimen, sSpecimen) & GetByOne(sTestSeq, sTestSeq) & "")
            
            sTmp1 = GetByOne(sSubMCd, sSubMCd)
                        
            If sTmp1 = "NNNN" Then
                Call .SetText(3, j, "N")
            ElseIf IsNumeric(Left$(sTmp1, 2)) = True Then
                Call .SetText(3, j, "S" & Left$(sTmp1, 2) & "")
            End If
            
            Call .SetText(4, j, GetByOne(sTestNm, sTestNm) & "")
        End With
    Next
    
    iSpdClick1 = 0
    iSpdClick2 = 0
End Sub

Private Sub spdTestItem_DblClick(ByVal Col As Long, ByVal Row As Long)
    With spdTestItem
        .Row = Row
        .Action = SS_ACTION_DELETE_ROW
        .MaxRows = .MaxRows - 1
    End With
End Sub

Private Sub txtPart_Change()
    On Error GoTo ErrHandler
    
    Dim i%
    
    If Len(txtPart) = txtPart.MaxLength Then
        'If iSpdClick1 = 1 Then
        'Else
            For i = 1 To giPartCnt
                lblPart.Caption = ""
                If gPartTable(i).sPartInit = txtPart Then
                    lblPart.Caption = gPartTable(i).sPartName
                    txtRtn = ""
                    txtRtn.SetFocus
                    Exit For
                End If
            Next
        'End If
    End If
    
ErrHandler:
End Sub

Private Sub txtPart_Click()
    Call Txt_Highlight(txtPart)
End Sub

Private Sub txtPart_GotFocus()
    Call Txt_Highlight(txtPart)
End Sub

Private Sub txtPart_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonPart_Click
    End Select
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtRtn.SetFocus
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

Private Sub txtRtn_Change()
    On Error GoTo ErrHandler
    If Len(txtRtn) = txtRtn.MaxLength Then
        If txtPart = "" Then
            MsgBox "PART 코드를 입력한 후 ROUTINE 순번을 입력하여 주십시요!!"
            Exit Sub
        Else
            If iSpdClick2 = 1 Then
            Else
                Call DisplayRtn
            End If
        End If
        
        txtRtnNm.SetFocus
    ElseIf Len(txtRtn) = 0 Then
        txtRtnNm = ""
        lblRtnNm.Caption = ""
    End If
    
ErrHandler:
End Sub

Private Sub txtRtn_Click()
    Call Txt_Highlight(txtRtn)
End Sub

Private Sub txtRtn_GotFocus()
    Call Txt_Highlight(txtRtn)
End Sub

Private Sub txtRtn_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonRtn_Click
    End Select
End Sub

Private Sub txtRtn_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtRtnNm.SetFocus
    End If
End Sub

Private Sub txtRtn_LostFocus()
    If Len(txtRtn) < txtRtn.MaxLength Then
        txtRtn = Format$(txtRtn, "000")
    End If
End Sub

Private Sub txtRtnNm_Click()
    Call Txt_Highlight(txtRtnNm)
End Sub

Private Sub txtRtnNm_GotFocus()
    Call Txt_Highlight(txtRtnNm)
End Sub
