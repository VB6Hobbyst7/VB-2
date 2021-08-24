VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FGS0401 
   BorderStyle     =   0  '없음
   Caption         =   "환자데이터조회 - DELTA CHECK"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "델타보기 설정"
      ForeColor       =   14737632
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.74
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      FloodColor      =   8421504
   End
   Begin Threed.SSFrame fraBefore 
      Height          =   3765
      Left            =   3810
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _Version        =   65536
      _ExtentX        =   14049
      _ExtentY        =   6641
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
      Begin FPSpread.vaSpread spdBefore 
         Height          =   3315
         Left            =   60
         OleObjectBlob   =   "FGS0401.frx":0000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   7845
      End
      Begin Threed.SSPanel pnlBeforeTitel 
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   3465
         _Version        =   65536
         _ExtentX        =   6112
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "   Before Data And Result ....."
         ForeColor       =   8454143
         BackColor       =   128
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
         BorderWidth     =   1
         Alignment       =   2
      End
   End
   Begin Threed.SSFrame fraDelta 
      Height          =   3765
      Left            =   3810
      TabIndex        =   2
      Top             =   3690
      Width           =   7965
      _Version        =   65536
      _ExtentX        =   14049
      _ExtentY        =   6641
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
      Begin FPSpread.vaSpread spdDelta 
         Height          =   3315
         Left            =   60
         OleObjectBlob   =   "FGS0401.frx":08E3
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   7845
      End
      Begin Threed.SSPanel pnlDeltaTitle 
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "   Delta Data And Result ....."
         ForeColor       =   8454143
         BackColor       =   128
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
         BorderWidth     =   1
         Alignment       =   2
      End
   End
   Begin Threed.SSFrame fraBasic 
      Height          =   7455
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   13150
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdDeltaItem 
         Height          =   2985
         Left            =   90
         OleObjectBlob   =   "FGS0401.frx":11C6
         TabIndex        =   8
         Top             =   2310
         Width           =   3585
      End
      Begin VB.OptionButton optView 
         Caption         =   "등록번호별"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   5850
         Width           =   1215
      End
      Begin VB.OptionButton optView 
         Caption         =   "작업번호별"
         Height          =   375
         Index           =   1
         Left            =   1275
         TabIndex        =   21
         Top             =   5850
         Width           =   1215
      End
      Begin VB.OptionButton optView 
         Caption         =   "검사항목별"
         Height          =   375
         Index           =   2
         Left            =   2490
         TabIndex        =   20
         Top             =   5850
         Width           =   1215
      End
      Begin VB.TextBox txtslipcd 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         MaxLength       =   3
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtDeltaLen 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "30"
         Top             =   1860
         Width           =   495
      End
      Begin Threed.SSPanel pnlSlip 
         Height          =   375
         Left            =   195
         TabIndex        =   11
         Top             =   210
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SLIP"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.74
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdsliph 
         Height          =   330
         Left            =   720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
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
         Picture         =   "FGS0401.frx":22F8
      End
      Begin MSComCtl2.DTPicker dtpSLabDate 
         Height          =   315
         Left            =   210
         TabIndex        =   13
         Top             =   1350
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36165
      End
      Begin Threed.SSPanel pnlLabDate 
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   960
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "접수일자"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.74
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   945
         Left            =   1350
         TabIndex        =   15
         Top             =   6300
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "조회 F3"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0401.frx":241A
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   945
         Left            =   2370
         TabIndex        =   16
         Top             =   6300
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "종료Esc"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0401.frx":2CF4
      End
      Begin Threed.SSPanel pnlDeltaLen 
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   1830
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Delta 기간"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.74
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlDeltacmt 
         Height          =   285
         Left            =   2250
         TabIndex        =   18
         Top             =   1890
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "일 이전까지"
         ForeColor       =   0
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin MSComCtl2.DTPicker dtpELabDate 
         Height          =   315
         Left            =   2010
         TabIndex        =   24
         Top             =   1350
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36165
      End
      Begin VB.Label lblWave 
         Caption         =   "~"
         Height          =   165
         Left            =   1770
         TabIndex        =   23
         Top             =   1410
         Width           =   165
      End
      Begin VB.Label lblslipnm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1020
         TabIndex        =   19
         Top             =   600
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FGS0401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DCJ0101     As DCJ0101
Dim DCS0401     As DCS0401

Dim Sys_Date    As String
Dim CodeHelp_F  As Integer

Private Sub fill_color()

    Dim iCnt        As Integer

    With spdDelta
        .Col = -1
        For iCnt = 1 To .MaxRows / 2
            
            .Row = iCnt * 2 - 1
            If iCnt Mod 2 = 0 Then
                .BackColor = 연초록
            Else
                .BackColor = 연빨강
            End If
            
            .Row = iCnt * 2
            If iCnt Mod 2 = 0 Then
                .BackColor = 연초록
            Else
                .BackColor = 연빨강
            End If
        Next
    End With
        
    With spdBefore
        .Col = -1
        For iCnt = 1 To .MaxRows / 2
            
            .Row = iCnt * 2 - 1
            If iCnt Mod 2 = 0 Then
                .BackColor = 연초록
            Else
                .BackColor = 연빨강
            End If
            
            .Row = iCnt * 2
            If iCnt Mod 2 = 0 Then
                .BackColor = 연초록
            Else
                .BackColor = 연빨강
            End If
        Next
    End With

End Sub
Private Sub search_clear(Optional position As String)
    
    Dim iCntDelta       As Integer
    Dim iCnt            As Integer
    Dim sDeltaInfo      As String
    Dim sDeltaChk       As String

    If position <> "ITEM" Then
        spdDeltaItem.Row = 1
        spdDeltaItem.Row2 = spdDeltaItem.MaxRows
        spdDeltaItem.Col = 1
        spdDeltaItem.Col2 = spdDeltaItem.MaxCols
        spdDeltaItem.BlockMode = True
        spdDeltaItem.Action = 3
        spdDeltaItem.BlockMode = False
        spdDeltaItem.MaxRows = 0
    End If

    spdBefore.Row = 3
    spdBefore.Row2 = spdBefore.MaxRows
    spdBefore.Col = -1
    spdBefore.BlockMode = True
    spdBefore.Action = 3
    spdBefore.BlockMode = False
    spdBefore.MaxRows = 2
    
    spdDelta.Row = 3
    spdDelta.Row2 = spdDelta.MaxRows
    spdDelta.Col = -1
    spdDelta.BlockMode = True
    spdDelta.Action = 3
    spdDelta.BlockMode = False
    spdDelta.MaxRows = 2
    
    If Trim(lblslipnm.Caption) <> "" Then
        Set DCS0401 = New DCS0401
            sDeltaInfo = DCS0401.Get_DeltaInfo(txtslipcd)
        Set DCS0401 = Nothing
        
        iCntDelta = Val(GetByOne(sDeltaInfo, sDeltaInfo))
        If iCntDelta = 0 Then
            ViewMsg "해당하는 SLIP에 Delta가 체크된 검사항목이 없습니다"
            spdDeltaItem.MaxRows = 0
            Exit Sub
        End If
        
        spdDeltaItem.MaxRows = iCntDelta
        
        For iCnt = 1 To iCntDelta
            Call spdDeltaItem.SetText(2, iCnt, GetByOne(sDeltaInfo, sDeltaInfo))
            Call spdDeltaItem.SetText(3, iCnt, GetByOne(sDeltaInfo, sDeltaInfo))
            sDeltaChk = GetByOne(sDeltaInfo, sDeltaInfo)
            If sDeltaChk = "1" Then
                Call spdDeltaItem.SetText(4, iCnt, "A")
            ElseIf sDeltaChk = "2" Then
                Call spdDeltaItem.SetText(4, iCnt, "%")
            End If
            Call spdDeltaItem.SetText(5, iCnt, txtslipcd + GetByOne(sDeltaInfo, sDeltaInfo))
        Next iCnt
        
    End If
        
End Sub


Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdQuery_Click()
    Dim bRtnCd          As Boolean
    Dim iCnt            As Integer
    Dim iDeltaCnt       As Integer
    Dim sDateGbn        As String
    Dim sDelta          As String
    Dim sRegNo          As String
    Dim sName           As String
    Dim sPrintNm        As String
    Dim sLabNo          As String
    Dim sOrdCd          As String
    Dim orderBy         As String
    
    Dim vDeltaChk
    Dim vOrdCd
    
' 필수항목 입력체크
    If Trim(lblslipnm.Caption) = "" Then
        ViewMsg "Slip코드를 입력하여 주십시요"
        txtslipcd.SetFocus
        Exit Sub
    End If

' 초기화 처리
    If pnlLabDate.Caption = "접수일자" Then
        sDateGbn = "LABDATE"
    ElseIf pnlLabDate.Caption = "결과완료일" Then
        sDateGbn = "RESULTDATE"
    End If

    vOrdCd = ""
    sOrdCd = " AND ( "

    For iCnt = 1 To spdDeltaItem.MaxRows
        bRtnCd = spdDeltaItem.GetText(1, iCnt, vDeltaChk)
        If vDeltaChk = "1" Then
            bRtnCd = spdDeltaItem.GetText(5, iCnt, vOrdCd)
            sOrdCd = sOrdCd _
                & " ( RESULT.SPECIMENCD   = '" & Mid(vOrdCd, 4, 3) & "' " _
                & " AND RESULT.TESTITEMSEQ  = '" & Mid(vOrdCd, 7, 3) & "' " _
                & " AND RESULT.SUBMCD       = '" & Mid(vOrdCd, 10, 4) & "' ) " _
                & " OR "
        End If
    Next iCnt

    sOrdCd = Left(sOrdCd, Len(sOrdCd) - 5) & " ) "

    If vOrdCd = "" Then
        ViewMsg "Delta 검사항목을 선택하여 주십시요"
        Exit Sub
    End If

    'sDateGbn

    For iCnt = 0 To 2
        If optView(iCnt).Value = True Then
            If iCnt = 0 Then
                 orderBy = " ORDER BY RESULT.REGNO"
            ElseIf iCnt = 1 Then
                orderBy = " ORDER BY RESULT.LABSEQ"
            ElseIf iCnt = 2 Then
                orderBy = " ORDER BY TESTITEM.PRINTORDER"
            End If
        End If
    Next

' 실 데이타 취득
    Set DCS0401 = New DCS0401
    sDelta = DCS0401.Get_Delta(sDateGbn, Format(dtpSLabDate.Value, "YYYYMMDD"), Format(dtpELabDate.Value, "YYYYMMDD"), txtslipcd, txtDeltaLen, sOrdCd, orderBy)
    Set DCS0401 = Nothing

' 데이타 표시
    iDeltaCnt = Val(GetByOne(sDelta, sDelta))

    spdBefore.MaxRows = (iDeltaCnt + 1) * 2
    spdDelta.MaxRows = (iDeltaCnt + 1) * 2
    
    Call fill_color

' 해당 Delta에 걸린 환자 정보및 결과
    For iCnt = 2 To iDeltaCnt + 1
        Call spdDelta.SetText(1, iCnt * 2 - 1, Trim(iCnt - 1)) 'Trim(iCnt * 2 - 3))       '순번
        sRegNo = GetByOne(sDelta, sDelta)
        Call spdDelta.SetText(2, iCnt * 2 - 1, sRegNo)                   '등록번호
        sName = GetByOne(sDelta, sDelta)
        Call spdDelta.SetText(3, iCnt * 2 - 1, sName)                    '이름
        Call spdDelta.SetText(4, iCnt * 2 - 1, GetByOne(sDelta, sDelta)) '진료과
        Call spdDelta.SetText(5, iCnt * 2 - 1, GetByOne(sDelta, sDelta)) '접수구분
        Call spdDelta.SetText(6, iCnt * 2 - 1, lblslipnm.Caption)        'SLIP명
        sPrintNm = GetByOne(sDelta, sDelta)
        Call spdDelta.SetText(7, iCnt * 2 - 1, sPrintNm)                 '검사명
        Call spdDelta.SetText(2, iCnt * 2, GetByOne(sDelta, sDelta))     '작업번호
        Call spdDelta.SetText(3, iCnt * 2, GetByOne(sDelta, sDelta))     ' 성별/나이
        Call spdDelta.SetText(4, iCnt * 2, GetByOne(sDelta, sDelta))     ' 병실(병동)
        Call spdDelta.SetText(5, iCnt * 2, GetByOne(sDelta, sDelta))     ' 응급
        Call spdDelta.SetText(6, iCnt * 2, GetByOne(sDelta, sDelta))     ' 검체명
        Call spdDelta.SetText(7, iCnt * 2, GetByOne(sDelta, sDelta))     ' 결과값
        
        Call spdBefore.SetText(1, iCnt * 2 - 1, Trim(iCnt - 1)) 'Trim(iCnt * 2 - 3))       '순번
        sLabNo = GetByOne(sDelta, sDelta)
        Call spdBefore.SetText(2, iCnt * 2 - 1, sRegNo)                   '등록번호
        Call spdBefore.SetText(3, iCnt * 2 - 1, sName)                    '이름
        Call spdBefore.SetText(6, iCnt * 2 - 1, lblslipnm.Caption)        'SLIP명
        Call spdBefore.SetText(7, iCnt * 2 - 1, sPrintNm) '검사명
        Call spdBefore.SetText(2, iCnt * 2, Left(sLabNo, 8) & "-" & Right(sLabNo, 5))     '과거작업번호
        Call spdBefore.SetText(7, iCnt * 2, GetByOne(sDelta, sDelta))     '과거결과
    Next iCnt
    
' 이전결과 정보
    For iCnt = 2 To iDeltaCnt + 1
        Call spdBefore.SetText(4, iCnt * 2 - 1, GetByOne(sDelta, sDelta)) '진료과
        Call spdBefore.SetText(5, iCnt * 2 - 1, GetByOne(sDelta, sDelta)) '접수구분
        Call spdBefore.SetText(3, iCnt * 2, GetByOne(sDelta, sDelta))     ' 성별/나이
        Call spdBefore.SetText(4, iCnt * 2, GetByOne(sDelta, sDelta))     ' 병실(병동)
        Call spdBefore.SetText(5, iCnt * 2, GetByOne(sDelta, sDelta))     ' 응급
        Call spdBefore.SetText(6, iCnt * 2, GetByOne(sDelta, sDelta))     ' 검체명
    Next iCnt

End Sub

Private Sub cmdsliph_Click()

    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    
    txtslipcd.SetFocus
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    j = CPart.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CPart
        sTot01 = .TotField01
        sTot02 = .TotField02
        sTot03 = .TotField03
    End With
    
    Set CPart = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01) & GetByOne(sTot02, sTot02)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot03, sTot03)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtslipcd.hwnd
    
    FSS0401.Left = 750
    FSS0401.Top = 1550
    
' Code Help Flag
    CodeHelp_F = True
    
    Load FSS0401
    FSS0401.Show vbModal

End Sub

Private Sub dtpSLabDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        txtDeltaLen.SetFocus
        KeyCode = 0
    End If
    
End Sub

Private Sub Form_Activate()

    If CodeHelp_F = False Then
        txtslipcd.Text = fCurUserSlipCd
        lblslipnm.Caption = fCurUserSlipNm
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyF3
        Call cmdQuery_Click
        KeyCode = 0
''    Case vbKeyF5
''        Call cmdPrint_Click
''        KeyCode = 0
    Case vbKeyEscape
        Call cmdExit_Click
        KeyCode = 0
    End Select
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Dim iCnt    As Integer
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        ViewMsg ""
    End If

End Sub

Private Sub Form_Load()
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11920
    Me.Height = 7950
    
    CodeHelp_F = False
    
    Set DCJ0101 = New DCJ0101
    Sys_Date = DCJ0101.Get_Date("D")
    dtpSLabDate.Value = Sys_Date
    dtpELabDate.Value = Sys_Date
    Set DCJ0101 = Nothing
    
    optView(0).Value = True
    
    txtDeltaLen.Text = "30"
    Call search_clear
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call InitRegCurFrmTitle
    
End Sub

Private Sub pnlLabDate_DblClick()
    
    If pnlLabDate.Caption = "접수일자" Then
        pnlLabDate.Caption = "결과완료일"
    ElseIf pnlLabDate.Caption = "결과완료일" Then
        pnlLabDate.Caption = "접수일자"
    End If

End Sub

Private Sub spdBefore_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    spdDelta.TopRow = NewTop
End Sub

Private Sub spdDelta_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    spdBefore.TopRow = NewTop
End Sub

Private Sub spdDeltaItem_Click(ByVal Col As Long, ByVal Row As Long)
    
    With spdDeltaItem
    
        .Col = 1
        .Row = Row

        If .Text = "1" Then
            Call .SetText(1, Row, "0")
        Else
            Call .SetText(1, Row, "1")
        End If
    End With

End Sub

'Private Sub spdDeltaItem_Click(ByVal Col As Long, ByVal Row As Long)
'
'    Dim iCnt        As Integer
'    Dim bRtnCd      As Boolean
'    Dim sDeltaChk
'
'    If Row > 0 Then
'        For iCnt = 1 To spdDeltaItem.MaxRows
'            bRtnCd = spdDeltaItem.GetText(1, iCnt, sDeltaChk)
'             If iCnt = Row And sDeltaChk = "1" Then
'                Call spdDeltaItem.SetText(1, iCnt, "0")
'                DoEvents
'                Exit Sub
'            End If
'            If sDeltaChk = "1" Then
'                Call spdDeltaItem.SetText(1, iCnt, "0")
'                Call search_clear("ITEM")
'                DoEvents
'            End If
'        Next iCnt
'        Call spdDeltaItem.SetText(1, Row, "1")
'        DoEvents
'    End If
'
'End Sub

Private Sub txtDeltaLen_Change()

    If Len(Trim(txtDeltaLen)) = txtDeltaLen.MaxLength Then
        spdDeltaItem.SetFocus
    End If
    
End Sub

Private Sub txtDeltaLen_GotFocus()

    Txt_Highlight txtDeltaLen
    
End Sub

Private Sub txtDeltaLen_LostFocus()
    
    If IsNumeric(txtDeltaLen.Text) = False Then
        ViewMsg "입력하신 Delta 기간이 잘못되었습니다."
    End If
    
End Sub

Private Sub txtslipcd_Change()

    If Len(txtslipcd.Text) = txtslipcd.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
        
        lblslipnm.Caption = DCJ0101.Get_SlipNm(txtslipcd.Text)
    
        If lblslipnm.Caption = "" Then
            ViewMsg "존재하지 않는 Slip Code입니다."
            Exit Sub
        End If
    
        Set DCJ0101 = Nothing
        
        Call search_clear
        
        If CodeHelp_F = False Then
            dtpSLabDate.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    Else
        lblslipnm.Caption = ""
    End If

End Sub

Private Sub txtslipcd_GotFocus()

    Txt_Highlight txtslipcd
    txtslipcd.Tag = txtslipcd.Text
    
End Sub

Private Sub txtslipcd_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CodeHelp_F = False

End Sub

Private Sub txtslipcd_LostFocus()

    If txtslipcd.Tag <> txtslipcd.Text Then
        'Call search_clear
    End If

End Sub
