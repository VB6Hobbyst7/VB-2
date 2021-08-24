VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm422RiPrint 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm422.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   270
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "핵의학실 결과지 출력 조건"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   75
      TabIndex        =   5
      Top             =   255
      Width           =   10740
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "외래"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   390
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "병동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   150
         Width           =   885
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00DBE6E6&
         Height          =   435
         Left            =   5880
         ScaleHeight     =   375
         ScaleWidth      =   4425
         TabIndex        =   6
         Top             =   180
         Width           =   4485
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FEF5F3&
            Caption         =   "결과보고"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  '그래픽
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FFF4FD&
            Caption         =   "일괄 재출력"
            Height          =   375
            Index           =   1
            Left            =   1485
            Style           =   1  '그래픽
            TabIndex        =   8
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00F7F7F7&
            Caption         =   "개별 재출력"
            Height          =   375
            Index           =   2
            Left            =   2955
            Style           =   1  '그래픽
            TabIndex        =   7
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker dtpVfyDt 
         Height          =   375
         Left            =   2475
         TabIndex        =   10
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   20840451
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   1275
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "보 고 일 자"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   5925
      Left            =   75
      ScaleHeight     =   5865
      ScaleWidth      =   10680
      TabIndex        =   28
      Top             =   2430
      Width           =   10740
      Begin FPSpread.vaSpread tblOrder 
         Height          =   5835
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   10292
         _StockProps     =   64
         BackColorStyle  =   3
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   20
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "frm422.frx":000C
         Appearance      =   1
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5850
         Left            =   0
         TabIndex        =   30
         Top             =   -15
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   10319
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridColor       =   14013909
         GridShowVert    =   0   'False
         MaxCols         =   46
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   16252927
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frm422.frx":0BE0
         TextTip         =   4
      End
      Begin FPSpread.vaSpread tblList 
         Height          =   5550
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   9790
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   9
         MaxRows         =   50
         OperationMode   =   1
         ShadowColor     =   15857140
         SpreadDesigner  =   "frm422.frx":1CF9
         UserResize      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   75
      TabIndex        =   11
      Top             =   900
      Width           =   10740
      Begin VB.PictureBox picESign 
         Height          =   500
         Left            =   5805
         ScaleHeight     =   435
         ScaleWidth      =   1140
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   7425
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   435
         Width           =   1320
      End
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00FEF5F3&
         Caption         =   "미리보기(&V)"
         Height          =   510
         Left            =   8745
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   435
         Width           =   1320
      End
      Begin VB.Frame fraSetWard 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   915
         Left            =   60
         TabIndex        =   14
         Top             =   195
         Width           =   6060
         Begin VB.CheckBox chkAllWard 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체병동/진료과"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2760
            TabIndex        =   20
            Top             =   135
            Width           =   1725
         End
         Begin VB.CommandButton cmdWardList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2400
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   19
            Top             =   75
            Width           =   315
         End
         Begin VB.TextBox txtWardId 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            TabIndex        =   18
            Top             =   90
            Width           =   1155
         End
         Begin VB.CommandButton cmdDoctList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2385
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   17
            Top             =   495
            Width           =   330
         End
         Begin VB.TextBox txtDoctId 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1230
            TabIndex        =   16
            Top             =   495
            Width           =   1140
         End
         Begin VB.CheckBox chkAllDoct 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2775
            TabIndex        =   15
            Top             =   540
            Width           =   705
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   45
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   90
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "병동/진료과"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   6
            Left            =   45
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "주치의"
            Appearance      =   0
         End
         Begin VB.Label lblDoctNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "DoctNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3525
            TabIndex        =   22
            Top             =   600
            Width           =   675
         End
         Begin VB.Label lblWardNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "WardNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4500
            TabIndex        =   21
            Top             =   210
            Width           =   720
         End
      End
      Begin VB.Frame fraLabNo 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   915
         Left            =   60
         TabIndex        =   23
         Top             =   195
         Visible         =   0   'False
         Width           =   6060
         Begin VB.TextBox txtPtId 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1260
            TabIndex        =   24
            Text            =   "S00"
            Top             =   225
            Width           =   1275
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   1
            Left            =   75
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   225
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "환자 ID"
            Appearance      =   0
         End
         Begin VB.Label lblPtNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "환자명1"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   2610
            TabIndex        =   27
            Top             =   330
            Width           =   630
         End
         Begin VB.Label lblSexAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "남/30"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3525
            TabIndex        =   26
            Top             =   345
            Width           =   450
         End
         Begin VB.Label lblWard 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "61W-111"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4695
            TabIndex        =   25
            Top             =   360
            Width           =   690
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "출   력 (&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   2145
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "결과지 출력 예정 리스트"
      LeftGab         =   100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 보고서 출력예정 건수 :"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   8715
      Width           =   2175
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   33
      Top             =   8670
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   " ☞ 출력대상자 리스트에서 선택하시면 출력 시 제외됩니다."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   8475
      Width           =   5955
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   570
      Index           =   0
      Left            =   90
      Shape           =   4  '둥근 사각형
      Top             =   8415
      Width           =   6255
   End
End
Attribute VB_Name = "frm422RiPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()

Private objSql       As New clsLISSqlReport
Private strStartDate As String
Private MsgFg As Boolean
Private PtFg As Boolean
Private ClearFg As Boolean

Dim blnLoadChk  As Boolean

Private Sub cmdPrint_Click()
    Dim strRstEntryType As String
    Dim strPtId         As String
    Dim strTestDiv      As String
    Dim strTable        As String
    Dim strSQL          As String
    Dim strImgPath      As String
    Dim i               As Long
    Dim j               As Long
    Dim objProgress     As jProgressBar.clsProgress
    Dim objProgress1    As jProgressBar.clsProgress
    Dim objReport       As clsBatchReport
    Dim strLastDt       As String
    Dim strLastTm       As String
    Dim strPrtDt        As String
    Dim strPrtTm        As String
    Dim lngErrCount     As Long
    
    Dim lngFileNo As Long
    
    lngFileNo = FreeFile
    
    If Printers.Count = 0 Then
        MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
        Exit Sub
    End If
    
    If Not optPrint(2).Value And Trim(txtWardId.Text) = "" Then
        MsgBox "결과지를 출력할 병동을 선택하십시오.", vbInformation, "병동선택"
        txtWardId.SetFocus
        Exit Sub
    End If
    
    If Not optPrint(2).Value And Trim(txtDoctId.Text) = "" Then
        MsgBox "주치의를 선택하십시오.", vbInformation, "주치의선택"
        txtDoctId.SetFocus
        Exit Sub
    End If
    
    If lblCnt.Caption = 0 Then
        MsgBox "출력할 대상 리스트가 없습니다.", vbInformation, "결과 출력"
        Exit Sub
    End If
    
    lngErrCount = 0
    
    MouseRunning
    
    Set objProgress = New jProgressBar.clsProgress
    
    With objProgress
        .Container = Me
        .Left = lblPrgBar.Left + 3
        .Top = lblPrgBar.Top + 3
        .Width = lblPrgBar.Width - 10
        .Height = lblPrgBar.Height - 10
        
'        .SetMyForm Me
'        .Choice = True
'        .Max = tblOrder.MaxRows
'        .Min = 0
'        .Value = 0
'        .XPos = lblPrgBar.Left + 3
'        .YPos = lblPrgBar.Top + 3
'        .XWidth = lblPrgBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblPrgBar.Height - 10 ' 260
        DoEvents
    End With
    
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        If optBussDiv(0).Value Then
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_외래.log" For Append As lngFileNo
        ElseIf optBussDiv(1).Value Then
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_병동.log" For Append As lngFileNo
        Else
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_종검.log" For Append As lngFileNo
        End If
    End If
    
'    Dim objWard As clsBasisData
    Dim strWard As String
    
    With tblOrder
        
        For i = 1 To .MaxRows
            
On Error GoTo Err_Trap1

            objProgress.Value = i
            
            .Row = i
            
            .Col = 1
            If .Value = 0 Then
                
                .TopRow = i
                
                .Col = 5    '환자명
                objProgress.Message = .Value & " 환자의 결과지를 출력하고 있습니다... ( " & i & " / " & .MaxRows & " )"

                .Col = 4    '환자ID
                strPtId = .Value
                
                .Col = 15   '전자서명 Path
                strImgPath = .Value

                .Col = 16   '보고서 종류
                strTestDiv = .Value
                
                picESign.Picture = LoadPicture(strImgPath)

                Set objReport = New clsBatchReport

                'Dictionary에 담기..레포트 출력
                .Col = 2:
'                Set objWard = Nothing
'                Set objWard = New clsBasisData
                strWard = GetWardNm(medGetP(.Value, 1, "-"))
'                Set objWard = Nothing
                
                If strWard <> "" Then
                    objReport.Ward = strWard
                    
                    If objReport.Ward <> "" Then
                        objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                    Else
                        objReport.Ward = Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                    End If
                End If
                
'                If ObjLISComCode.WardID.Exists(medgetp(.Value, 1, "-")) = True Then
'                    ObjLISComCode.WardID.KeyChange (medgetp(.Value, 1, "-"))
'                    objReport.Ward = ObjLISComCode.WardID.Fields("wardnm")
'
'                    If objReport.Ward <> "" Then
'                        objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                    Else
'                        objReport.Ward = Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                    End If
'                End If
                
                .Col = 3:  objReport.Doct = .Value
                .Col = 4:  objReport.ptid = .Value
                .Col = 5:  objReport.PtNm = .Value
                .Col = 6:  objReport.PtSex = medGetP(.Value, 1, "/")
                           objReport.PtAge = medGetP(.Value, 2, "/")
                .Col = 10: objReport.VfyDt = .Value
                '.Col = 11: objReport.VfyDt = objReport.VfyDt & " " & .Value
                .Col = 12: objReport.VfyNM = .Value
                .Col = 13: objReport.MdfDt = .Value         '수정일
                .Col = 17: objReport.ICD = .Value
                
                
                
                '병동에서 출력할때만 레포트제목에 재발행/회진용 표기
                If gUsingInWardMenu Then
                    objReport.Rouding = optPrint(3).Value       '회진레포트 여부
                    objReport.Reprint = optPrint(2).Value       '재발행 여부
                    objReport.BatchReprint = True
                Else
                    objReport.Rouding = optPrint(3).Value       '회진레포트 여부
                    objReport.Reprint = optPrint(2).Value       '재발행 여부
                    objReport.BatchReprint = optPrint(1).Value
                End If
                objReport.Special = IIf(strTestDiv = enTestDiv.TST_SpeTest, True, False)
                
                .Col = 18:
'                Set objWard = Nothing
'                Set objWard = New clsBasisData
                strWard = GetDeptNm(.Value)
'                Set objWard = Nothing
                
                If strWard <> "" Then
                    objReport.Dept = .Value
                    objReport.DeptNm = strWard
                End If
'                If ObjLISComCode.DeptCd.Exists(.Value) Then
'                    Call ObjLISComCode.DeptCd.KeyChange(.Value)
'                    objReport.Dept = .Value
'                    objReport.DeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                End If
                
                If optPrint(0).Value And (Not gUsingInWardMenu) Then
                    Print #lngFileNo, "( " & i & " / " & .MaxRows & " )  " & Now & "   " & strPtId & "," & objReport.PtNm & "," & objReport.DeptNm & "," & objReport.Ward
                End If
                
                objReport.RiPrint = "핵의학"
                Call objReport.ReportForOnePatient(strPtId, strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), _
                                                   strTestDiv, strImgPath, picESign, objProgress, strLastDt, strLastTm)
            End If
        Next
        objReport.RiPrint = ""
    End With

    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Close #lngFileNo
    End If

    MouseDefault
    
    Set objProgress = Nothing
    Set objProgress1 = Nothing
    
    If lngErrCount > 0 Then
        For i = tblOrder.DataRowCnt To 1 Step -1
            tblOrder.Row = i
            tblOrder.Col = 20
            If tblOrder.Value = "0" Then
                tblOrder.Action = ActionDeleteRow
            End If
        Next
        MsgBox "다음 환자들의 결과지 출력 중 오류가 발생했습니다. 다시 출력하십시오.", vbExclamation, "오류"
    Else
        cmdClear_Click
    End If
    
    Exit Sub
    
Err_Trap:
'==================================================================
    'DBConn.RollbackTrans
'==================================================================
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Print #lngFileNo, "DB ERROR : " & Err.Description
    End If
On Error GoTo Err_Trap
    Resume Next

Err_Trap1:
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Print #lngFileNo, "VB ERROR : " & Err.Description
    End If
On Error GoTo Err_Trap1
    Resume Next
    
End Sub

Private Sub cmdQuery_Click()

    Dim objReport   As New clsBatchReport
    Dim objESign    As clsLISElectronSign
    Dim objProgress As clsProgress
    Dim Rs          As New Recordset
    Dim rs1         As New Recordset
    Dim EmpRs       As Recordset
    Dim strWA       As String
    Dim strTable    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strReferral As String
    Dim strSex      As String
    Dim strStsCd    As String
    Dim strMsg      As String
    Dim i           As Long
    Dim strEmpId    As String
    Dim strBussDiv  As String
    Dim strChkLoad  As String
    Dim strTestDiv  As String
    Dim strKey      As String
    Dim strDOB      As String
    Dim strSQL      As String
    'Dim strSEX      As String
    
    tblOrder.MaxRows = 0
    lblCnt.Caption = 0
    
    strBussDiv = IIf(optBussDiv(0).Value, enBussDiv.BussDiv_OutPatient, enBussDiv.BussDiv_InPatient)
    
    If optPrint(2).Value = True Then
        strStartDate = Format(dtpVfyDt.Value, CS_DateDbFormat)
    Else
        strStartDate = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    End If
    
    '프로그래스바 생성..
    Set objProgress = New clsProgress
    objProgress.Container = MainFrm.stsbar
    objProgress.Message = "자료를 읽고 있습니다..."
    objProgress.Max = 100
'    objProgress.Caption = "처리중입니다."
'    objProgress.Mode = 0
'    objprogress.message = "자료를 읽고 있습니다."
'    objProgress.Max = 100
'    objProgress.Min = 0
'    objProgress.Value = 0
'    objProgress.Visible = True
    
    objSql.RiPrint = "핵의학"

    
    Dim strWard As String
    Dim strDoct As String
    
    If txtWardId.Text <> CS_AllCaption Then strWard = txtWardId.Text
    If txtDoctId.Text <> CS_AllCaption Then strDoct = txtDoctId.Text
    
    If optPrint(2).Value = True Then
        '개별재출력
        Rs.Open objSql.GetAccLAbNoLIS201(txtPtId.Text, Format(dtpVfyDt.Value, CS_DateDbFormat)), DBConn
        tblOrder.ZOrder 0
        
    ElseIf optPrint(0).Value = True Then
        '일괄출력
        Rs.Open objSql.RiReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), strBussDiv, "", strWard, strDoct), DBConn
        tblOrder.ZOrder 0
    ElseIf optPrint(1).Value = True Then
        '일괄재출력
        Rs.Open objSql.RiReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), strBussDiv, "Y", strWard, strDoct), DBConn
        tblOrder.ZOrder 0
    End If
    
    If Rs.EOF Then
        Set objProgress = Nothing
        MsgBox "해당 데이타가 없습니다.", vbInformation, "결과지 출력"
        GoTo Nodata
    End If
    
'    Dim objEmp As clsBasisData
    Dim strEmp As String
    
    strKey = ""
    With tblOrder
        
        If Rs.RecordCount > 0 Then

            '프로그래스바 생성..
            objProgress.Max = Rs.RecordCount
            objProgress.Min = 0
            objProgress.Value = 0

            .ReDraw = False
            
            i = 1
            Do Until Rs.EOF = True
            
                If strKey = "" & Rs.Fields("deptcd").Value & _
                                 Rs.Fields("ptid").Value & _
                                 Rs.Fields("testdiv").Value Then
                    '환자/진료과/보고서종류가 같은 경우엔 수정여부와 수정일만 보여주기...
                    If "" & Rs.Fields("stscd").Value = enStsCd.StsCd_LIS_Modify Then
                        .Col = 9
                        .Value = "수정"
                    End If
                    If Trim("" & Rs.Fields("mfydt").Value) <> "" Then
                        .Col = 13
                        .Value = Format(Mid("" & Rs.Fields("mfydt").Value, 3), CS_DateShortMask)
                    End If
                    GoTo Skip
                End If
                    
                .MaxRows = i
                .Row = i

                .Col = 2: .Value = "" & Rs.Fields("location").Value
                .Col = 18: .Value = "" & Rs.Fields("deptcd").Value
                
                If optPrint(2).Value Then
                    If lblWard.Caption <> "" Then
                        .Col = 2: .Value = lblWard.Caption
                        .Col = 19: .Value = lblWard.Caption
                    End If
                Else
                    If optBussDiv(1).Value Then
                        .Col = 19: .Value = "" & Rs.Fields("location").Value
                    End If
                End If
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                Set objEmp = Nothing
                
                .Col = 3: .Value = strEmp 'GetEmpName(rs.Fields("majdoct").Value & "")
                .Col = 4: .Value = "" & Rs.Fields("ptid").Value
                .Col = 5: .Value = "" & Rs.Fields("ptnm").Value
                

                If IsNumeric("" & Rs.Fields("sex").Value) Then
                    strSex = IIf(Val("" & Rs.Fields("sex").Value) Mod 2 = 1, "남", "여")
                Else
                    strSex = IIf("" & Rs.Fields("sex").Value = "M", "남", "여")
                End If
                
                .Col = 6: .Value = strSex '
                          
                           strDOB = Rs.Fields("dob").Value & ""
                           If Len(strDOB) = 6 Then strDOB = strDOB & "01"
                            .Value = .Value & "/" & DateDiff("yyyy", Format(strDOB, CS_DateMask), Now)
                
                .Col = 16
                .Value = "" & Rs.Fields("testdiv").Value
                
                .Col = 7
                Select Case "" & Rs.Fields("testdiv").Value
                    Case enTestDiv.TST_RouTest
                        .Value = "일반"
                    Case enTestDiv.TST_SpeTest
                        .Value = "기타"
                    Case enTestDiv.TST_MicTest
                        .Value = "미생물"
                End Select

                .Col = 8: .Value = 1
                .Col = 9
                Select Case "" & Rs.Fields("stscd").Value
                Case enStsCd.StsCd_LIS_MidRst
                    .Value = "중간"
                Case enStsCd.StsCd_LIS_FinRst
                    .Value = "최종"
                Case enStsCd.StsCd_LIS_Modify
                    .Value = "수정"
                End Select

                .Col = 10: .Value = Format(Mid("" & Rs.Fields("vfydt").Value, 3), CS_DateShortMask)
                .Col = 11: .Value = Format(Mid("" & Rs.Fields("vfytm").Value, 1, 4), CS_TimeShortMask)

                strEmpId = "" & Rs.Fields("vfyid").Value
                .Col = 14: .Value = strEmpId
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                Set objEmp = Nothing
                
                .Col = 12:  .Value = strEmp ' GetEmpName(strEmpId)
               


                .Col = 13: .Value = Format(Mid("" & Rs.Fields("mfydt").Value, 3), CS_DateShortMask)
                '임상진단....
                Dim objDisease  As New clsDisease
                
                objDisease.ptid = Rs.Fields("ptid").Value
                
                .Col = 17: .Value = objDisease.Disease
                
                Set objDisease = Nothing
                
                strKey = "" & Rs.Fields("deptcd").Value & _
                              Rs.Fields("ptid").Value & _
                              Rs.Fields("testdiv").Value
                              
                i = i + 1
                objProgress.Value = objProgress.Value + 1
Skip:
                Rs.MoveNext
            Loop
            Set objProgress = Nothing
            .ReDraw = True
            lblCnt.Caption = .MaxRows
        Else
            If optPrint(0).Value = True Then
                strMsg = "결과보고"
            ElseIf optPrint(1).Value = True Then
                strMsg = "일괄재출력"
            ElseIf optPrint(2).Value = True Then
                strMsg = "개별재출력"
            End If
            MsgBox strMsg & " 내역이 없습니다.", vbCritical, "결과 출력"
            medClearTable tblOrder
            tblOrder.MaxRows = 0
            lblCnt.Caption = 0
        End If

    End With

Nodata:
    Set Rs = Nothing
    Set objProgress = Nothing

End Sub
Private Sub TxtClear()
    '결과지 출력 조건
    dtpVfyDt.Value = GetSystemDate

    '결과지 출력예정리스트
    medClearTable tblOrder
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
    
    lblWard.Caption = ""
    tblOrder.MaxRows = 0
    tblOrdSheet.MaxRows = 0
    tblOrder.ZOrder 0

    txtWardId.Text = "(전체)"
    txtDoctId.Text = "(전체)"
    lblCnt.Caption = 0
    chkAllWard.Value = 1
    chkAllDoct.Value = 1
    txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    cmdPreview.Caption = "미리보기(&V)"
    cmdPreview.Tag = ""
    tblOrdSheet.Visible = False
    tblOrdSheet.ZOrder 1

End Sub
Private Sub chkAllDoct_Click()

    lblDoctNm.Caption = ""
    txtDoctId.Text = Choose(chkAllDoct.Value + 1, "", CS_AllCaption)
    txtDoctId.Enabled = Choose(chkAllDoct.Value + 1, True, False)
    cmdDoctList.Enabled = Choose(chkAllDoct.Value + 1, True, False)
End Sub

Private Sub chkAllWard_Click()

    lblWardNm.Caption = ""
    txtWardId.Text = Choose(chkAllWard.Value + 1, "", CS_AllCaption)
    txtWardId.Enabled = Choose(chkAllWard.Value + 1, True, False)
    cmdWardList.Enabled = Choose(chkAllWard.Value + 1, True, False)
End Sub
Private Sub Form_Load()
    lblWardNm.Caption = ""
    lblWard.Caption = ""
    
    optBussDiv(0).Enabled = True
    optBussDiv(0).Value = True

    blnLoadChk = False
    TxtClear

End Sub
Private Sub cmdClear_Click()
    TxtClear
End Sub
Private Sub cmdPreview_Click()
    
    Dim i As Long
    Dim strPtId    As String
    Dim strPtNm    As String
    Dim strVfyDt   As String
    Dim strTestDiv As String
    
    If cmdPreview.Tag = "1" Then
        cmdPreview.Caption = "미리보기(&V)"
        cmdPreview.Tag = ""
        tblOrdSheet.Visible = False
        tblOrdSheet.ZOrder 1
    Else
        If tblOrder.MaxRows = 0 Then Exit Sub
        
        Dim objProgress As New clsProgress
        objProgress.Container = MainFrm.stsbar
        objProgress.Message = "자료를 읽고 있습니다..."
        objProgress.Max = tblOrder.MaxRows
'        objProgress.Caption = "처리중입니다."
'        objProgress.Mode = 0
'        objprogress.message = "자료를 읽고 있습니다."
'        objProgress.Max = tblOrder.MaxRows
'        objProgress.Min = 0
'        objProgress.Value = 0
'        objProgress.Visible = True
        
        tblOrdSheet.MaxRows = 0
        
        For i = 1 To tblOrder.MaxRows
            objProgress.Value = i
            tblOrder.Row = i
            tblOrder.Col = 1
            If tblOrder.Value = 1 Then GoTo Skip
            tblOrder.Col = 4:  strPtId = tblOrder.Value
            tblOrder.Col = 5:  strPtNm = tblOrder.Value
            tblOrder.Col = 16: strTestDiv = tblOrder.Value
            strVfyDt = Format(dtpVfyDt.Value, CS_DateDbFormat)
            objProgress.Message = strPtNm & "환자의 결과내역을 읽고 있습니다."
            DoEvents
            Call DisplayOrders(strPtId, strPtNm, strVfyDt, strTestDiv)
Skip:
        Next
        cmdPreview.Caption = "닫기(&B)"
        cmdPreview.Tag = "1"
        tblOrdSheet.Visible = True
        tblOrdSheet.ZOrder 0
    End If
End Sub

Private Sub cmdWardList_Click()
'% 병동코드 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
'    Dim objWard As New clsBasisData
    Dim strCaption As String
    Dim strHead As String
    
    
    If optBussDiv(0).Value Then
        strCaption = "진료과 조회"
        strHead = "부서코드;부서명"
    Else
        strCaption = "병동 조회"
        strHead = "병동코드;병동명"
    End If
    
    With objMyList

        .FormCaption = strCaption
        .ColumnHeaderText = strHead
        .Tag = "WardID"
        Me.ScaleMode = 1
        If optBussDiv(0).Value Then
'            Call .ListPop(, 3950, 6300, ObjLISComCode.DeptCd)
            Call .LoadPopUp(GetSQLDeptList) ', 3950, 6300)
        Else
'            Call .ListPop(, 3950, 6300, ObjLISComCode.WardID)
            Call .LoadPopUp(GetSQLWardList) ', 3950, 6300)
        End If
        
        txtWardId.Text = medGetP(.SelectedString, 1, ";")
        lblWardNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objWard = Nothing
    Set objMyList = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub
Private Sub cmdDoctList_Click()

'% 주치의 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
'    Dim objDoct As New clsBasisData

    With objMyList
        .FormCaption = "주치의리스트"

        .ColumnHeaderText = "의사ID;의사명"
        .Tag = "DoctID"
        Me.ScaleMode = 1
'        Call .ListPop(getdoctlistsql, 3950, 6300)
        Call .LoadPopUp(GetSQLDoctList) ', 3950, 6300)

        txtDoctId.Text = medGetP(.SelectedString, 1, ";")
        lblDoctNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objDoct = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdExit_Click()
    Set objSql = Nothing
    Unload Me
    
    RaiseEvent FormClose
End Sub

Private Function DisplayOrders(ByVal pPtId As String, ByVal pPtNm As String, ByVal pVfyDt As String, ByVal pTestDiv As String) As Boolean

    Dim SqlStmt         As String
    Dim ColCnt          As Integer
    Dim tmpTestNm       As String
    
    Dim SvKeyDt         As String
    Dim SvSpcNm         As String
    Dim pWorkArea       As String
    Dim pAccDt          As String
    Dim pAccSeq         As String
    Dim strKeyFld       As String
    Dim strNotice       As String
    Dim strTmp          As String
    Dim i               As Integer
    Dim j               As Integer
    Dim MySql           As New clsLISSqlReview     'Sql문 클래스
    Dim tmpRs           As New Recordset
    Dim tVfyDt          As String
    
    
    Me.Enabled = False
   
    MouseRunning
    tVfyDt = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    MySql.RiPrint = "핵의학"
    '처방일/접수일 기준
    SqlStmt = MySql.SqlQueryAllResults(pPtId, "examdt", tVfyDt, pVfyDt, pTestDiv)
    
    'Query
    tmpRs.Open SqlStmt, DBConn
    
    SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    ReDim aryMesg(0)
    DisplayOrders = False
    
    If tmpRs.EOF Then GoTo Nodata
    
    With tblOrdSheet
      
        '.ReDraw = False
      
        Do Until tmpRs.EOF
         
            If Trim("" & tmpRs.Fields("RstCd").Value) = "" Then GoTo Skip
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Value = pPtId: .ForeColor = DCM_Gray
            .Col = 2: .Value = pPtNm:  .ForeColor = DCM_Gray
            
            If SvKeyDt <> Trim("" & tmpRs.Fields("KeyDate").Value) Then
                .Col = 3:   .Value = Trim("" & tmpRs.Fields("KeyDate").Value)
                            .FontBold = True: .ForeColor = vbBlack       '-- 보고일
                .Col = 4:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                            .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                SvKeyDt = Trim("" & tmpRs.Fields("KeyDate").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                .Col = 1:   .FontBold = True: .ForeColor = vbBlack
                .Col = 2:   .FontBold = True: .ForeColor = vbBlack
            Else
                .Col = 3:   .Value = "":
                            .FontBold = True: .ForeColor = vbBlack       '-- 처방일
                            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                                .Col = 4:
                                .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                            Else
                                .Col = 4:
                                .Value = "":
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                            End If
            End If
            
            .Col = 34:  .Value = Trim("" & tmpRs.Fields("KeyDate").Value)   '처방일
            .Col = 35:  .Value = Trim("" & tmpRs.Fields("SpcNm").Value)     '검체명
            
            .Col = 5:   '-- 검사명
                        .ForeColor = DCM_MidBlue
                        tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestLongNm").Value), 1, 33)
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            
                            .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), ".")
                        Else
                            .Value = Space(4) & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")
                        End If
                        
            .Col = 6:   '-- 결과명(코드일 경우..)
                        .ForeColor = DCM_Brown   '갈색
                        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
                            .Value = "미확"
                            .ForeColor = DCM_MidGray: .FontBold = False:
                        Else
                            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                                .TypeHAlign = TypeHAlignCenter
                                .Value = Trim("" & tmpRs.Fields("RstCd").Value)
                            Else
                                .CellType = CellTypeEdit
                                .TypeHAlign = TypeHAlignLeft
                                .Value = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
                            End If
                            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                                .Value = "Growth"
                            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                                .Value = Space(3)
                            End If
                        End If
                        
            .Col = 7:   '-- 결과단위
                        .Value = Trim("" & tmpRs.Fields("RstUnit").Value)
            
            .Col = 8    '-- High / Low
                        .Value = ""
                        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG:  .ForeColor = DCM_LightBlue
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                        End If
            
            .Col = 9:   '-- Delta/Panic
                        .Value = Trim("" & tmpRs.Fields("DPDiv").Value): .ForeColor = vbRed
            
            .Col = 10:   '-- 참고치
                        If Trim("" & tmpRs.Fields("RstDiv").Value) <> "*" And Trim("" & tmpRs.Fields("TestDiv").Value) < "4" Then .Value = CS_QuestionMark
            
            .Col = 11:   '-- More Result...
                        .Value = "": .ForeColor = DCM_LightBlue
                        If Trim("" & tmpRs.Fields("TxtFg").Value) > "0" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("TxtFg").Value) = "Y" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then .Value = CS_FingerMark
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            If Trim("" & tmpRs.Fields("FootNoteFg").Value) = "1" Then .Value = CS_FingerMark
                            If Trim("" & tmpRs.Fields("RmkCd").Value) <> "" Then .Value = CS_FingerMark
                        End If
                        If Trim("" & tmpRs.Fields("DcFg").Value) = "1" Then .Value = .Value & "*"
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "4" Then .Value = CS_FingerMark    '해부병리
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "5" Then .Value = CS_FingerMark    '혈액은행
         
            .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdDate").Value)        '-- 처방일
            .Col = 13: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)          '-- 처방번호
            .Col = 14: .Value = Trim("" & tmpRs.Fields("OrdDoct").Value)        '-- 처방의
            .Col = 15: .Value = Trim("" & tmpRs.Fields("ColDtTm").Value)        '-- 채혈일시
            .Col = 16: .Value = Trim("" & tmpRs.Fields("ColId").Value)          '-- 채혈자
            .Col = 17: .Value = Trim("" & tmpRs.Fields("RcvDtTm").Value)        '-- 접수일시
            .Col = 18: .Value = Trim("" & tmpRs.Fields("RcvId").Value)          '-- 접수자
            .Col = 19: .Value = Trim("" & tmpRs.Fields("WorkArea").Value):  pWorkArea = .Value  'WorkArea
            .Col = 20: .Value = Trim("" & tmpRs.Fields("AccDt").Value):     pAccDt = .Value     'AccDt
            .Col = 21: .Value = Trim("" & tmpRs.Fields("AccSeq").Value):    pAccSeq = .Value    'AccSeq
            .Col = 22: .Value = Trim("" & tmpRs.Fields("LastRst").Value)        '-- 최근결과
            .Col = 23: .Value = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)     '-- 최근결과일시
            .Col = 24: .Value = Trim("" & tmpRs.Fields("LastVfyId").Value)      '-- 최근결과 보고자
            .Col = 25: .Value = Trim("" & tmpRs.Fields("VfyDtTm").Value)        '-- 보고일시
            .Col = 26: .Value = Trim("" & tmpRs.Fields("VfyId").Value)          '-- 보고자
            .Col = 27: .Value = Trim("" & tmpRs.Fields("Sex").Value)            '-- Sex
            .Col = 28: .Value = Trim("" & tmpRs.Fields("AgeDay").Value)         '-- AgeDay
            .Col = 29: .Value = Trim("" & tmpRs.Fields("TestCd").Value)         '-- 검사코드
            .Col = 30: .Value = Trim("" & tmpRs.Fields("SpcCd").Value)          '-- 검체코드
            .Col = 31: .Value = Trim("" & tmpRs.Fields("VfyDt").Value)          '-- 보고일
            .Col = 32: .Value = Trim("" & tmpRs.Fields("TestDiv").Value)        '-- 검사구분
            .Col = 33: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)         '-- 진료과
            .Col = 36: .Value = Trim("" & tmpRs.Fields("TxtFg").Value)          '-- 소견결과여부
            .Col = 37: .Value = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
            .Col = 38: .Value = Trim("" & tmpRs.Fields("RmkCd").Value)          '-- Remark 코드
            .Col = 39: .Value = Trim("" & tmpRs.Fields("SenFg").Value)          '-- 감수성 여부
            .Col = 40: .Value = Trim("" & tmpRs.Fields("OrdDiv").Value)         '-- 처방구분
            .Col = 41: .Value = Trim("" & tmpRs.Fields("UnitQty").Value)        '-- 수혈수량
            .Col = 42: .Value = Trim("" & tmpRs.Fields("ReqDt").Value)          '-- 수혈예정일
            .Col = 43: .Value = Trim("" & tmpRs.Fields("ReqTm").Value)          '-- 수혈예정시간
            .Col = 44: .Value = Trim("" & tmpRs.Fields("WardId").Value)         '-- 병동
            .Col = 45: .Value = Trim("" & tmpRs.Fields("HosilId").Value)        '-- 호실
            .Col = 46: .Value = Trim("" & tmpRs.Fields("RoomId").Value)        '-- 호실
            
'            ReDim Preserve aryMesg(UBound(aryMesg) + 1)
'            aryMesg(UBound(aryMesg).value) = Trim("" & tmpRs.fields("Mesg").value)    '-- 진료과Remark
            If Trim("" & tmpRs.Fields("Notice").Value) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 5
                .TypeEditMultiLine = False
                .ForeColor = vbBlack
                .Value = "☞ Clinical Notice "  '& vbCrLf & Trim("" & tmpRs.fields("Notice").value)
                .RowHeight(.MaxRows) = .MaxTextRowHeight(.MaxRows)
                strNotice = Trim("" & tmpRs.Fields("Notice").Value)
                strNotice = Replace(strNotice, vbCr, "")
                strTmp = medShift(strNotice, vbLf)
                While strTmp <> ""
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 5
                    .TypeEditMultiLine = False
                    .ForeColor = &H747474
                    .Value = strTmp
                    strTmp = medShift(strNotice, vbLf)
                Wend
            End If
      
         
            DisplayOrders = True
Skip:
            tmpRs.MoveNext
        Loop
      
        .Row = -1: .Col = 6: .Col2 = 5
        .BlockMode = True
        .AllowCellOverflow = True
        .BlockMode = False
      
        .RowHeight(-1) = 11.5
        .ReDraw = True
      
        'If chkRefVal.Value = 0 Then GoTo ExitPos
        GoTo ExitPos
      
        Dim tmpTestCd As String
        Dim tmpSpcCd As String
        Dim tmpVfyDt As String
        Dim tmpSex As String
        Dim tmpAgeDay As String
        Dim tmpRs1 As New Recordset
        Dim tmpRefFromVal As Double
        Dim tmpRefToVal As Double
        Dim tmpRefCd As String
      
        DoEvents
        For i = 1 To .MaxRows
            '참고치 검색
            .Row = i
            .Col = 10: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 27:  tmpSex = Trim(.Value)
            .Col = 28:  tmpAgeDay = Trim(.Value)
            .Col = 29:  tmpTestCd = Trim(.Value)
            .Col = 30:  tmpSpcCd = Trim(.Value)
            .Col = 31:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            tmpRs1.Open SqlStmt, DBConn
            
            If tmpRs1.EOF Then
                '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
                '--> 거의 Both로 등록됨.
                SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                tmpRs1.Open SqlStmt, DBConn
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            Set tmpRs1 = Nothing
            For j = i To .MaxRows
                .Row = j
                .Col = 29   '참고치
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 10: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents

RefSkip:
        Next
      
ExitPos:
        'If .MaxRows < 20 Then .MaxRows = 20
      
    End With
   
Nodata:
    Me.Enabled = True
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Function

Private Sub optBussDiv_Click(Index As Integer)
    cmdClear_Click
End Sub

Private Sub optPrint_Click(Index As Integer)
    
    If optPrint(2).Value = True Then

        fraLabNo.Visible = True
        fraSetWard.Visible = False
        txtPtId.Text = ""
        txtPtId.SetFocus
    Else
        fraLabNo.Visible = False
        fraSetWard.Visible = True
    End If
    
    If optPrint(0).Value = True Then
        If gUsingInWardMenu Then
            chkAllWard.Value = 0
            chkAllWard.Visible = False
        End If
    End If
    

    optBussDiv(0).Enabled = True
    If gUsingInWardMenu Then
        chkAllDoct.Enabled = True
        chkAllDoct.Value = 1
    Else
        chkAllDoct.Enabled = True
        chkAllWard.Enabled = True
        chkAllDoct.Value = 1
        chkAllWard.Value = 1
    End If
    
    tblList.Visible = False
    tblOrder.Visible = True
    tblOrder.ZOrder 0

    dtpVfyDt.Value = GetSystemDate
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    '결과지 출력예정리스트
    medClearTable tblOrder

    lblCnt.Caption = 0

End Sub

Private Sub GetTestlist()
    Dim Rs As New Recordset
    Dim strTestNM As String
    Dim ii As Long
    Dim jj As Long
    
    Rs.Open objSql.GetTestReportList, DBConn
    If Rs.RecordCount > 0 Then
        ii = 0
        jj = 0
        strTestNM = ""
        With tblList
        
            .Row = 1: .Row2 = .MaxRows
            .Col = 1: .Col2 = 1
            .BlockMode = True
            .AllowCellOverflow = False
            .BlockMode = False
            
            .ReDraw = False
            .MaxRows = Rs.RecordCount + 1
            .Row = ii: .Col = 0
            .Value = "검사명/등록번호" & vbNewLine & "환자명" & vbNewLine & "병동/병실"
            ii = 1
            Rs.MoveFirst
            Do Until Rs.EOF
                ii = ii + 1
                .Row = ii
                .Col = 0
                .RowHeight(ii) = 9.5
                If Rs.Fields("panelfg").Value & "" = "D" Then strTestNM = Rs.Fields("cdval1").Value & ""
                jj = Len(strTestNM)
                If strTestNM = Mid(Rs.Fields("cdval1").Value & "", 1, jj) And jj <> "0" And strTestNM <> Rs.Fields("cdval1").Value & "" Then
                    .Value = Space(4) & Rs.Fields("field1").Value & "": .TypeHAlign = TypeHAlignLeft
                Else
                    .Value = Space(1) & Rs.Fields("field1").Value & "": .TypeHAlign = TypeHAlignLeft
                End If
                .Col = 1
                .Value = Rs.Fields("cdval1").Value & "": .ForeColor = vbWhite
                Rs.MoveNext
            Loop
            .ReDraw = True
        End With
    End If
    Set Rs = Nothing
End Sub

Private Sub tblOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If MsgFg Then Exit Sub
    
    Dim lngButtonValue As Long
    Dim i As Long
    Dim strDept As String
    
    With tblOrder
        .Row = Row
        .Col = Col
        lngButtonValue = .Value
        If .Value = 1 Then
            lblCnt.Caption = Val(lblCnt.Caption) - 1
        Else
            lblCnt.Caption = Val(lblCnt.Caption) + 1
            Exit Sub
        End If
        
        .Col = 2
        strDept = medGetP(.Value, 1, "-")
        For i = 1 To tblOrder.DataRowCnt
            MsgFg = True
            .Row = i
            .Col = 2
            If strDept = medGetP(.Value, 1, "-") Then
                .Col = 1
                .Value = lngButtonValue
            End If
            MsgFg = False
        Next
    End With
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    
    Dim objPatient As clsPatient       '환자 클래스
    
    If Not gUsingInWardMenu Then

        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If Screen.ActiveControl Is Nothing Then Exit Sub
        
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    End If
    
    If MsgFg Then Exit Sub
      
    If txtPtId.Text = "" Then
        
        Exit Sub
    End If
    
    Set objPatient = New clsPatient
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With objPatient
        If Trim(txtPtId.Text) <> "" And .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSexAge.Caption = .SEXNM & " / " & .Age & " " & .AGEDIV
            If .WardID = "" Then
                lblWard.Caption = ""
            Else
                lblWard.Caption = .WardID & "-" & .ROOMID
            End If
            PtFg = True
            ClearFg = False
        Else
            If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요..", vbInformation
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Set objPatient = Nothing
            Exit Sub
        End If
    End With
    
    Set objPatient = Nothing

    Exit Sub

End Sub
