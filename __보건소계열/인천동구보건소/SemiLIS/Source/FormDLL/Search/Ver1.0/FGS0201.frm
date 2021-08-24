VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGS0201 
   BorderStyle     =   0  '없음
   Caption         =   "환자데이터조회 - 환자 HISTORY"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame3 
      Height          =   1545
      Left            =   8400
      TabIndex        =   17
      Top             =   0
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "이 름"
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   660
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "나 이"
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "성 별"
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
      Begin VB.Label lblName 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "남궁옥분씨애기"
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
         Left            =   1200
         TabIndex        =   23
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "130"
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
         Left            =   1200
         TabIndex        =   22
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "소아"
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
         Left            =   1200
         TabIndex        =   21
         Top             =   1020
         Width           =   495
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1545
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   8355
      _Version        =   65536
      _ExtentX        =   14737
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
      Begin VB.TextBox txtName 
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
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "오세원"
         Top             =   1050
         Width           =   825
      End
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
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "H02"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRegNo 
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
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "720121-1840518"
         Top             =   1050
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtpSLabDate 
         Height          =   315
         Left            =   1590
         TabIndex        =   6
         Top             =   645
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
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   630
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
         Left            =   3350
         TabIndex        =   7
         Top             =   645
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1035
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "등록번호"
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
      Begin Threed.SSPanel Panel3D3 
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   225
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
         Left            =   2100
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
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
         Picture         =   "FGS0201.frx":0000
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   945
         Left            =   6060
         TabIndex        =   2
         Top             =   360
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "조회 F3"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0201.frx":0122
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   945
         Left            =   7080
         TabIndex        =   4
         Top             =   360
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "종료Esc"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0201.frx":09FC
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   945
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1667
         _StockProps     =   78
         Caption         =   "인쇄 F5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FGS0201.frx":12D6
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Index           =   0
         Left            =   3285
         TabIndex        =   24
         Top             =   1035
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "이름"
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
         Height          =   330
         Left            =   2430
         TabIndex        =   16
         Top             =   240
         Width           =   2385
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
         Left            =   3120
         TabIndex        =   10
         Top             =   645
         Width           =   195
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   6045
      Left            =   30
      TabIndex        =   13
      Top             =   1470
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   10663
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
         Height          =   5775
         Left            =   30
         OleObjectBlob   =   "FGS0201.frx":1BB0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   11685
      End
   End
End
Attribute VB_Name = "FGS0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim regNo As Boolean

Dim slip As Boolean

Dim isTextRow(9) As Integer      '스프레드에 뿌릴때 마지막으로 뿌렸던 Row 기억
    
Dim workNo As Variant            '스프레드에 마지막으로 뿌려진 작업번호기억

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{ESC}"
    End If

End Sub

Private Sub cmdPrint_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{F5}"
    End If

End Sub

Private Sub cmdQuery_Click()
    Dim DCS0201 As DCS0201
    Dim count As Integer
    Dim spdRow As Integer
    Dim jubsu$, emer$
    Dim recordField01$, recordField02$, recordField03$, recordField04$
    Dim recordField05$, recordField06$, recordField07$, recordField08$, recordField09$

    If regNo = False Then
        ViewMsg "등록번호를 정확히 입력하십시오."
        txtRegNo.SetFocus
        Exit Sub
    End If

    If slip = False Then
        ViewMsg "SLIP을 정확히 입력하십시오."
        txtSlip.SetFocus
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    Set DCS0201 = New DCS0201

    Call SpreadClear

    With DCS0201
        .Get_Contents txtRegNo.text, Format(dtpSLabDate.Value, "YYYYMMDD"), Format(dtpELabDate.Value, "YYYYMMDD"), txtSlip.text
'        .Get_Contents txtName.text, Format(dtpSLabDate.Value, "YYYYMMDD"), Format(dtpELabDate.Value, "YYYYMMDD"), txtSlip.text
        count = .Getcount
        If count = 0 Then
            ViewMsg "조회된 내용이 없습니다."
            Set DCS0201 = Nothing
            spdList.MaxRows = spdRow
            MousePointer = 0
            Exit Sub
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
    End With


    For spdRow = 1 To count

        spdList.MaxRows = spdRow

        Call spdCmp(1, spdRow, GetByOne(recordField01, recordField01) & "-" & txtSlip.text & "-" & GetByOne(recordField02, recordField02))
       
        Select Case GetByOne(recordField03, recordField03)
            Case "0": jubsu = "외래"
            Case "1": jubsu = "입원"
            Case "2": jubsu = "수탁"
        End Select
        Call spdCmp(2, spdRow, jubsu)
        
        Select Case GetByOne(recordField04, recordField04)
            Case "": emer = "N"
            Case "0": emer = "N"
            Case "1": emer = "Y"
        End Select
        Call spdCmp(3, spdRow, emer)
        
        Call spdCmp(4, spdRow, GetByOne(recordField05, recordField05))
        Call spdCmp(5, spdRow, GetByOne(recordField06, recordField06))
        Call spdCmp(6, spdRow, GetByOne(recordField07, recordField07))
        Call spdCmp(7, spdRow, lblSlip.Caption)
        Call spdList.SetText(8, spdRow, GetByOne(recordField08, recordField08))
        Call spdList.SetText(9, spdRow, GetByOne(recordField09, recordField09))
    
    Next
    
    MousePointer = 0

    ViewMsg "총 " & count & "개의 자료가 조회되었습니다."

End Sub

Private Sub spdCmp(ByVal Col As Long, ByVal Row As Long, ByVal text As String)

        '위 col의 text와 비교해서 같으면 뿌리지 않는다.
    Dim data As Variant
    Dim i%

    With spdList

        If Row = 1 Then
            Call .SetText(Col, Row, text)
            isTextRow(Col) = 1
            Call .GetText(1, 1, workNo)
            Exit Sub
        Else
            If Col <> 1 Then
                Call .GetText(1, isTextRow(Col), data)
            Else
                data = text
            End If
            
            If workNo <> data Then
                If Col = 1 Then
                    For i = 1 To 9
                        isTextRow(i) = Row
                    Next
                    workNo = data
                    Call .SetText(Col, Row, text)
                End If
            Else
                If Col <> 1 Then
                    Call .GetText(Col, isTextRow(Col), data)
                    If data <> text Then
                        Call .SetText(Col, Row, text)
                        isTextRow(Col) = Row
                    End If
                End If
            End If
        End If
    End With

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
'        Case vbKeyF5: Call cmdPrint_Click       '/* 인 쇄 */
        Case vbKeyEscape: Unload Me             '/* 종 료 */
    End Select

End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11920
    Me.Height = 7950

    FGS0201.KeyPreview = True
    '폼이 로드될때 초기화
    txtSlip.text = fCurUserSlipCd
    lblSlip.Caption = fCurUserSlipNm
    txtRegNo.text = ""
    txtName.text = ""
    lblName.Caption = ""
    lblAge.Caption = ""
    lblSex.Caption = ""
    dtpSLabDate.Value = Format(Now, "yyyy-mm-dd")
    dtpELabDate.Value = Format(Now, "yyyy-mm-dd")
    Call SpreadClear
    
    regNo = False
    slip = True

End Sub

Private Sub SpreadClear()

    '폼이 로드되거나 조회를 실행할때 스프레드 내용을 지운다.
    With spdList
        .Row = -1
        .Col = -1
        .text = ""
        .BackColor = 연하늘
    End With

End Sub

Private Sub SSPanel1_DblClick(Index As Integer)

    If SSPanel1(1).Caption = "등록번호" Then
        SSPanel1(1).Caption = "이름"
        SSPanel2.Caption = "등록번호"
    Else
        SSPanel1(1).Caption = "등록번호"
        SSPanel2.Caption = "이름"
    End If
    txtRegNo.text = ""
    lblName.Caption = ""
    lblAge.Caption = ""
    lblSex.Caption = ""
    txtRegNo.SetFocus

End Sub

Private Sub ssSlipHelp_Click()
    Dim i%
    Dim count%
    Dim DCS0201 As DCS0201
    Dim recordField01$, recordField02$, recordField03$
    
    Set DCS0201 = New DCS0201
    
    DCS0201.Get_SlipCD
    
    count = DCS0201.Getcount
    
    Erase gCodeHlpTable '배열 초기화

    ReDim gCodeHlpTable(count) As CodeTBL

    With DCS0201
        recordField01 = .GetrecordField01
        recordField02 = .GetrecordFiled02
        recordField03 = .GetrecordFiled03
    End With
    
    Set DCS0201 = Nothing

    For i = 1 To count
        gCodeHlpTable(i).sCode = GetByOne(recordField01, recordField01) & GetByOne(recordField02, recordField02)
        gCodeHlpTable(i).sCodeNm = GetByOne(recordField03, recordField03)
    Next
    
    giCodeHlpCnt = count

    hWndCd = txtSlip.hwnd

    FSS0201.Left = 2500
    FSS0201.Top = 1570

    Load FSS0201
    FSS0201.Show vbModal

    'txtSlip과 lblSlip에 조회내용을 표시한다.
    Call txtSlip_LostFocus
'    txtRegNo.SetFocus

End Sub

Private Sub Text1_LostFocus()

    Dim flag As Boolean
    Dim DCS0201 As DCS0201

    If SSPanel1(1).Caption = "등록번호" Then
        Call getName
    Else
        Call getRegNo
    End If

End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtName_LostFocus()
    Dim flag As Boolean
    Dim DCS0201 As DCS0201

    If SSPanel1(0).Caption = "이름" Then
        Call getRegNo
    Else
        Call getName
    End If


End Sub


Private Sub txtRegNo_Change()
    If txtRegNo.SelStart = txtRegNo.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtRegNo_GotFocus()
    txtRegNo.SelStart = 0
    txtRegNo.SelLength = txtRegNo.MaxLength

End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSlip_LostFocus()
    
    Dim DCS0201 As DCS0201

    Set DCS0201 = New DCS0201

    DCS0201.Get_SlipNm txtSlip.text
    
    With DCS0201

        If .Getcount = 0 Then
            ViewMsg "입력하신 SLIP은 존재하지 않습니다."
            slip = False
            lblSlip.Caption = ""
            Exit Sub
        End If

        lblSlip.Caption = .GetrecordField01

    End With

    Set DCS0201 = Nothing
    
    txtSlip.text = UCase$(txtSlip.text)
    
    slip = True

End Sub

Private Sub txtSlip_Change()
    If txtSlip.SelStart = txtSlip.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub txtSlip_GotFocus()
    txtSlip.SelStart = 0
    txtSlip.SelLength = txtRegNo.MaxLength

End Sub

Private Sub txtSlip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub getName()

    '등록번호로 이름을 가져온다(Panel의 Caption이 등록번호 일때)
    Dim DCS0201 As DCS0201

    Set DCS0201 = New DCS0201
    
    DCS0201.Get_RegNo Trim(txtRegNo.text), True

    With DCS0201
    
        If .Getcount = 0 Then
            ViewMsg "입력하신 등록번호는 존재하지 않습니다."
            lblName.Caption = ""
            lblSex.Caption = ""
            lblAge.Caption = ""
            regNo = False
            Exit Sub
        End If

        lblName = .GetrecordField01

        If .GetrecordFiled02 = "1" Or .GetrecordFiled02 = "3" Then
            lblSex = "남"
        ElseIf .GetrecordFiled02 = "2" Or .GetrecordFiled02 = "4" Then
            lblSex = "여"
        End If

        If Trim(.GetrecordFiled03) <> "" Then
            lblAge = CInt(Left(DateTime.Date$, 4)) - CInt(.GetrecordFiled03)
        End If

    End With
    
    Set DCS0201 = Nothing
    
    regNo = True
    
    Call cmdQuery_Click

End Sub

Private Sub getRegNo()

    '이름으로 등록번호를 가져올때(Panel의 Caption이 이름일때)

    Dim i%
    Dim count%
    Dim DCS0201 As DCS0201
    Dim name$
    Dim recordField01$, recordField02$, recordField03$
    Dim jumpFlag As Boolean
    
    Set DCS0201 = New DCS0201
    
    DCS0201.Get_Name txtName.text
    
    count = DCS0201.Getcount
    
'    name = txtRegNo.text
    name = txtName.text

    If count = 0 Then
        ViewMsg "입력하신 이름은 존재하지 않습니다."
        lblName.Caption = ""
        lblSex.Caption = ""
        lblAge.Caption = ""
        regNo = False
        Exit Sub
    ElseIf count = 1 Then
        txtRegNo = lblName.Caption
        jumpFlag = True
        GoTo GetAgeAndSex       '검색결과가 하나이면 바로 조회
    End If
    
    Erase gCodeHlpTable '배열 초기화

    ReDim gCodeHlpTable(count) As CodeTBL

    With DCS0201
        recordField01 = .GetrecordField01
    End With
    
    Set DCS0201 = Nothing

    For i = 1 To count
        gCodeHlpTable(i).sCode = txtRegNo.text
        gCodeHlpTable(i).sCodeNm = GetByOne(recordField01, recordField01)
    Next

    giCodeHlpCnt = count

    hWndCd = txtRegNo.hwnd

    FSS0202.Left = 1620
    FSS0202.Top = 2365

    Load FSS0202
    FSS0202.Show vbModal

GetAgeAndSex:

    SSPanel1(1).Caption = "등록번호"
    SSPanel2.Caption = "이름"
    lblName.Caption = name
    
    '이름으로 정보들을 가져온다(Panel의 Caption이 이름일 때)
    Set DCS0201 = New DCS0201

    DCS0201.Get_RegNo Trim(name), False

    With DCS0201

        If jumpFlag = True Then
            txtRegNo.text = .GetrecordField01
        End If

        If .GetrecordFiled02 = "1" Or .GetrecordFiled02 = "3" Then
            lblSex = "남"
        ElseIf .GetrecordFiled02 = "2" Or .GetrecordFiled02 = "4" Then
            lblSex = "여"
        End If

        If Trim(.GetrecordFiled03) <> "" Then
            lblAge = CInt(Left(DateTime.Date$, 4)) - CInt(.GetrecordFiled03)
        End If

    End With
    
    Set DCS0201 = Nothing
    
    regNo = True
    
    Call cmdQuery_Click

End Sub
Private Sub txtRegNo_LostFocus()

    Dim flag As Boolean
    Dim DCS0201 As DCS0201

    If SSPanel1(1).Caption = "등록번호" Then
        Call getName
    Else
        Call getRegNo
    End If

End Sub

