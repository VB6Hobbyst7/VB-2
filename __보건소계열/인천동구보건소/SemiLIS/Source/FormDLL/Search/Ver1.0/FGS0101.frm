VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGS0101 
   BorderStyle     =   0  '없음
   Caption         =   "환자데이터조회 - 선택한 날짜구간의 환자데이터 조회"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame3 
      Height          =   1545
      Left            =   7140
      TabIndex        =   15
      Top             =   0
      Width           =   4635
      _Version        =   65536
      _ExtentX        =   8176
      _ExtentY        =   2725
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
      Begin VB.TextBox txtRem 
         BackColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   4455
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   150
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "병명기호(임상소견)"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1545
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   2725
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
      Begin VB.TextBox txtSlip 
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
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "H02"
         Top             =   270
         Width           =   495
      End
      Begin VB.TextBox txtEMin 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3270
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "MM"
         Top             =   1050
         Width           =   435
      End
      Begin VB.TextBox txtEHour 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2730
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "HH"
         Top             =   1050
         Width           =   405
      End
      Begin VB.TextBox txtSMin 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "MM"
         Top             =   1050
         Width           =   435
      End
      Begin VB.TextBox txtSHour 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1500
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "HH"
         Top             =   1050
         Width           =   405
      End
      Begin MSComCtl2.DTPicker dtpSLabDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   7
         Top             =   660
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
         Format          =   24379395
         CurrentDate     =   36165
      End
      Begin Threed.SSPanel pnlLabDate 
         Height          =   330
         Left            =   150
         TabIndex        =   10
         Top             =   645
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "접수일자"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
      End
      Begin MSComCtl2.DTPicker dtpELabDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   660
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
         Format          =   24379395
         CurrentDate     =   36165
      End
      Begin Threed.SSPanel pnlLabTime 
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   1035
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "접수시간"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   945
         Left            =   4920
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "조회 F3"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         RoundedCorners  =   0   'False
         Picture         =   "FGS0101.frx":0000
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   945
         Left            =   5910
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   330
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "종료 ESC"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         RoundedCorners  =   0   'False
         Picture         =   "FGS0101.frx":08DA
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Left            =   150
         TabIndex        =   21
         Top             =   255
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "SLIP"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
      End
      Begin Threed.SSCommand ssSlipHelp 
         Height          =   330
         Left            =   2010
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
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
         Picture         =   "FGS0101.frx":11B4
      End
      Begin VB.Label lblSlip 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "일반혈액검사"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   23
         Top             =   270
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3180
         TabIndex        =   20
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1950
         TabIndex        =   19
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2520
         TabIndex        =   18
         Top             =   1110
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3030
         TabIndex        =   11
         Top             =   720
         Width           =   195
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   6015
      Left            =   30
      TabIndex        =   12
      Top             =   1470
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   10610
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
      Begin FPSpread.vaSpread spdList 
         Height          =   5745
         Left            =   300
         OleObjectBlob   =   "FGS0101.frx":12D6
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   11175
      End
   End
End
Attribute VB_Name = "FGS0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slip As Boolean

Dim time As Boolean


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{ESC}"
    End If

End Sub

Private Sub cmdQuery_Click()

    Dim DCS0101 As DCS0101
    Dim count As Integer
    Dim spdRow As Integer
    Dim jubsu$, emer$, sex$
    Dim recordField01$, recordField02$, recordField03$, recordField04$
    Dim recordField05$, recordField06$, recordField07$, recordField08$
    Dim recordField09$, recordField10$, recordField11$, recordField12$

    Call SpreadClear
    
    If Trim(txtSHour.Text) = "" Then
        ViewMsg "조회시작시간이 입력되지 않았습니다."
        txtSHour.SetFocus
        Exit Sub
    End If

    If Trim(txtSMin.Text) = "" Then
        ViewMsg "조회시작시간이 입력되지 않았습니다."
        txtSMin.SetFocus
        Exit Sub
    End If

    If Trim(txtEHour.Text) = "" Then
        ViewMsg "조회종료시간이 입력되지 않았습니다."
        txtEHour.SetFocus
        Exit Sub
    End If

    If Trim(txtEMin.Text) = "" Then
        ViewMsg "조회종료시간이 입력되지 않았습니다."
        txtEMin.SetFocus
        Exit Sub
    End If

    If slip = False Then
        ViewMsg "SLIP을 정확히 입력하십시오."
        txtSlip.SetFocus
        Exit Sub
    End If
    
    MousePointer = vbHourglass

    Set DCS0101 = New DCS0101

    Call SpreadClear

    With DCS0101
        .Get_Contents txtSlip.Text, Format(dtpSLabDate.Value, "YYYYMMDD"), Format(dtpELabDate.Value, "YYYYMMDD"), _
                        txtSHour.Text & txtSMin.Text, txtEHour.Text & txtEMin.Text
        count = .Getcount
        If count = 0 Then
            ViewMsg "조회된 내용이 없습니다."
            Set DCS0101 = Nothing
            spdList.MaxRows = spdRow
            MousePointer = 0
            Exit Sub
        End If
        
        If count > 1000 Then
            If MsgBox("자료가 1000개 이상입니다. 조회하시겠습니까?", vbYesNo) = vbNo Then
                MousePointer = 0
                Set DCS0101 = Nothing
                Exit Sub
            End If
        End If

        recordField01 = .GetrecordField01
        recordField02 = .GetrecordFiled02
        recordField03 = .GetrecordFiled03
        recordField04 = .GetrecordFiled04
        recordField05 = .GetrecordFiled05
        recordField06 = .GetrecordFiled06
        recordField07 = .GetrecordFiled07
        recordField08 = .GetrecordFiled08
        recordField09 = .GetrecordFiled09
        recordField10 = .GetrecordFiled10
        recordField11 = .GetrecordFiled11
        recordField12 = .GetrecordFiled12
    End With

    Set DCS0101 = Nothing

    For spdRow = 1 To count

        With spdList
        
        .MaxRows = spdRow

        Call .SetText(1, spdRow, GetByOne(recordField01, recordField01) & "-" & txtSlip.Text & "-" & GetByOne(recordField02, recordField02))
       
        Call .SetText(2, spdRow, GetByOne(recordField03, recordField03))
        
        Call .SetText(3, spdRow, GetByOne(recordField04, recordField04))
        Call .SetText(4, spdRow, GetByOne(recordField05, recordField05))

        Select Case GetByOne(recordField06, recordField06)
            Case "1", "3": sex = "남"
            Case "2", "4": sex = "여"
        End Select
        Call .SetText(5, spdRow, sex)
        
        Select Case GetByOne(recordField07, recordField07)
            Case "0": jubsu = "외래"
            Case "1": jubsu = "입원"
            Case "2": jubsu = "수탁"
        End Select
        Call .SetText(6, spdRow, jubsu)

        Select Case GetByOne(recordField08, recordField08)
            Case "": emer = "N"
            Case "0": emer = "N"
            Case "1": emer = "Y"
        End Select
        Call .SetText(7, spdRow, emer)
        
        Call .SetText(8, spdRow, GetByOne(recordField09, recordField09))
        Call .SetText(9, spdRow, GetByOne(recordField10, recordField10))
        Call .SetText(10, spdRow, GetByOne(recordField11, recordField11))
        
        Call .SetText(11, spdRow, GetByOne(recordField12, recordField12))
        
        End With
    
    Next

    MousePointer = 0
    ViewMsg "총 " & count & "개의 자료가 조회되었습니다."


End Sub

Private Sub cmdQuery_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{F3}"
    End If

End Sub

Private Sub dtpELabDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub dtpSLabDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3: Call cmdQuery_Click       '/* 조 회 */
        Case vbKeyEscape: Unload Me             '/* 종 료 */
    End Select

End Sub

Private Sub spdList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim remark As Variant
    Call spdReverse(spdList, 1, 10, Row, Row, 연빨강, 2)
    Call spdList.GetText(11, Row, remark)
    txtRem.Text = remark

End Sub

Private Sub ssSlipHelp_Click()

    Dim i%
    Dim count%
    Dim DCS0101 As DCS0101
    Dim recordField01$, recordField02$, recordField03$
    
    Set DCS0101 = New DCS0101
    
    DCS0101.Get_SlipCD
    
    count = DCS0101.Getcount
    
    Erase gCodeHlpTable '배열 초기화

    ReDim gCodeHlpTable(count) As CodeTBL

    With DCS0101
        recordField01 = .GetrecordField01
        recordField02 = .GetrecordFiled02
        recordField03 = .GetrecordFiled03
    End With
    
    Set DCS0101 = Nothing

    For i = 1 To count
        gCodeHlpTable(i).sCode = GetByOne(recordField01, recordField01) & GetByOne(recordField02, recordField02)
        gCodeHlpTable(i).sCodeNm = GetByOne(recordField03, recordField03)
    Next
    
    giCodeHlpCnt = count

    hWndCd = txtSlip.hwnd

    FSS0101.Left = 2500
    FSS0101.Top = 1570

    Load FSS0101
    FSS0101.Show vbModal

    'txtSlip과 lblSlip에 조회내용을 표시한다.
    Call txtSlip_LostFocus
    
    txtSHour.SetFocus

End Sub


Private Sub txtEHour_Change()
    If txtEHour.SelStart = txtEHour.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtEHour_GotFocus()
    Txt_Highlight txtEHour
    txtEHour.SelStart = 0
    txtEHour.SelLength = txtEHour.MaxLength

End Sub

Private Sub txtEHour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtEHour_LostFocus()
    If txtEHour.Text > 23 Or txtEHour.Text < txtSHour.Text Then
        time = False
        Exit Sub
    End If
    
    time = True

End Sub

Private Sub txtEMin_Change()
    If txtEMin.SelStart = txtEMin.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtEMin_GotFocus()
    Txt_Highlight txtEMin
    txtEMin.SelStart = 0
    txtEMin.SelLength = txtEMin.MaxLength

End Sub

Private Sub txtEMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtEMin_LostFocus()
    If txtSHour.Text > 59 Then
        ViewMsg "조회마감시간이 범위를 넘었습니다."
        txtEMin.SetFocus
        Exit Sub
    End If
    
    If (txtSHour.Text & txtSMin.Text) > (txtEHour.Text & txtEMin.Text) Then
        ViewMsg "조회시간의 범위가 잘못지정되었습니다."
        txtEHour.SetFocus
        Exit Sub
    End If

    If time Then Call cmdQuery_Click
End Sub

Private Sub txtSHour_Change()
    If txtSHour.SelStart = txtSHour.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtSHour_GotFocus()
    Txt_Highlight txtSHour
    txtSHour.SelStart = 0
    txtSHour.SelLength = txtSHour.MaxLength

End Sub

Private Sub txtSHour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtSHour_LostFocus()
    If txtSHour.Text > 23 Then
        ViewMsg "조회시작시간이 범위를 넘었습니다."
        txtSHour.SetFocus
        Exit Sub
    End If

    If txtSHour.Text > txtEHour.Text Then
        txtEHour.Text = txtSHour.Text
    End If

End Sub

Private Sub txtSlip_LostFocus()
    
    Dim DCS0101 As DCS0101
    
    Set DCS0101 = New DCS0101
    
    DCS0101.Get_SlipName txtSlip.Text
    
    With DCS0101

        If .Getcount = 0 Then
            ViewMsg "입력하신 SLIP은 존재하지 않습니다."
            slip = False
            lblSlip.Caption = ""
            Exit Sub
        End If

        lblSlip.Caption = .GetrecordField01

    End With

    Set DCS0101 = Nothing

    txtSlip.Text = UCase$(txtSlip.Text)

    slip = True

End Sub

Private Sub txtSlip_Change()
    If txtSlip.SelStart = txtSlip.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtSlip_GotFocus()
    Txt_Highlight txtSlip
    txtSlip.SelStart = 0
    txtSlip.SelLength = txtSlip.MaxLength

End Sub

Private Sub txtSlip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11920
    Me.Height = 7950

    FGS0101.KeyPreview = True
    txtSlip.Text = fCurUserSlipCd
    lblSlip.Caption = fCurUserSlipNm
    dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
    dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    txtSHour.Text = "00"
    txtSMin.Text = "00"
    txtEHour.Text = "23"
    txtEMin.Text = "59"
    Call SpreadClear

    slip = True
    time = False
End Sub

Private Sub SpreadClear()

    '폼이 로드되거나 조회를 실행할때 스프레드 내용을 지운다.
    With spdList
        .Row = -1
        .Col = -1
        .Text = ""
        .BackColor = 연하늘
    End With

End Sub

Private Sub pnlLabDate_DblClick()
    If pnlLabDate.Caption = "접수일자" Then
        pnlLabDate.Caption = "결과완료일"
        pnlLabTime.Caption = "결과완료시간"
    ElseIf pnlLabDate.Caption = "결과완료일" Then
        pnlLabDate.Caption = "접수일자"
        pnlLabTime.Caption = "접수시간"
    End If
End Sub

Private Sub pnlLabTime_DblClick()
    If pnlLabTime.Caption = "접수시간" Then
        pnlLabDate.Caption = "결과완료일"
        pnlLabTime.Caption = "결과완료시간"
    ElseIf pnlLabTime.Caption = "결과완료시간" Then
        pnlLabDate.Caption = "접수일자"
        pnlLabTime.Caption = "접수시간"
    End If
End Sub

Private Sub txtSMin_Change()
    If txtSMin.SelStart = txtSMin.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtSMin_GotFocus()
    Txt_Highlight txtSMin
    txtSMin.SelStart = 0
    txtSMin.SelLength = txtSMin.MaxLength

End Sub

Private Sub txtSMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtSMin_LostFocus()
    If txtSMin.Text > 59 Then
        ViewMsg "조회시작시간이 범위를 넘었습니다."
        txtSMin.SetFocus
        Exit Sub
    End If

End Sub

