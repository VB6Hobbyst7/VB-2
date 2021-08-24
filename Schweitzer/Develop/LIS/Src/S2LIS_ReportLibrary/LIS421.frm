VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm421TubercleReport 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11295
   ControlBox      =   0   'False
   Icon            =   "LIS421.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin Crystal.CrystalReport crtReport 
      Left            =   90
      Top             =   3765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   5925
      Left            =   75
      ScaleHeight     =   5865
      ScaleWidth      =   10680
      TabIndex        =   26
      Top             =   2445
      Width           =   10740
      Begin FPSpread.vaSpread tblOrder 
         Height          =   5505
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   10680
         _Version        =   196608
         _ExtentX        =   18838
         _ExtentY        =   9710
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
         GrayAreaBackColor=   16777215
         MaxCols         =   77
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "LIS421.frx":144A
         Appearance      =   1
      End
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   270
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   10725
      _ExtentX        =   18918
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
      Caption         =   "항산성균 감수성 결과 출력 조건"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   75
      TabIndex        =   2
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
         Left            =   375
         TabIndex        =   25
         Top             =   420
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
         Left            =   375
         TabIndex        =   24
         Top             =   165
         Width           =   885
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00DBE6E6&
         Height          =   435
         Left            =   6060
         ScaleHeight     =   375
         ScaleWidth      =   4410
         TabIndex        =   3
         Top             =   195
         Width           =   4470
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FEF5F3&
            Caption         =   "결과보고"
            Height          =   390
            Index           =   0
            Left            =   -15
            Style           =   1  '그래픽
            TabIndex        =   6
            Top             =   -30
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FFF4FD&
            Caption         =   "일괄 재출력"
            Height          =   390
            Index           =   1
            Left            =   1470
            Style           =   1  '그래픽
            TabIndex        =   5
            Top             =   -15
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00F7F7F7&
            Caption         =   "개별 재출력"
            Height          =   390
            Index           =   2
            Left            =   2940
            Style           =   1  '그래픽
            TabIndex        =   4
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker dtpVfyDt 
         Height          =   375
         Left            =   2700
         TabIndex        =   7
         Top             =   225
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
         Left            =   1515
         TabIndex        =   36
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
         Caption         =   "보 고 일 자"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   1
      Top             =   2145
      Width           =   10725
      _ExtentX        =   18918
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   75
      TabIndex        =   8
      Top             =   915
      Width           =   10740
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9165
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   405
         Width           =   1320
      End
      Begin VB.Frame fraLabNo 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   915
         Left            =   75
         TabIndex        =   19
         Top             =   165
         Visible         =   0   'False
         Width           =   6060
         Begin VB.TextBox txtPtId 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1230
            TabIndex        =   20
            Text            =   "S00"
            Top             =   225
            Width           =   1275
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   1
            Left            =   60
            TabIndex        =   37
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
            Left            =   2595
            TabIndex        =   23
            Top             =   315
            Width           =   630
         End
         Begin VB.Label lblSexAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "남/30"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3555
            TabIndex        =   22
            Top             =   315
            Width           =   450
         End
         Begin VB.Label lblWard 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "61W-111"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4650
            TabIndex        =   21
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame fraSetWard 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   915
         Left            =   75
         TabIndex        =   10
         Top             =   165
         Width           =   6060
         Begin VB.CheckBox chkAllWard 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체병동/진료과"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2715
            TabIndex        =   16
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
            Left            =   2340
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   15
            Top             =   90
            Width           =   300
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
            Left            =   1260
            TabIndex        =   14
            Top             =   105
            Width           =   1065
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
            Height          =   375
            Left            =   2340
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   13
            Top             =   495
            Width           =   300
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
            Left            =   1260
            TabIndex        =   12
            Top             =   495
            Width           =   1065
         End
         Begin VB.CheckBox chkAllDoct 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2760
            TabIndex        =   11
            Top             =   540
            Width           =   705
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   75
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   105
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
            Left            =   75
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   495
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
            TabIndex        =   18
            Top             =   585
            Width           =   675
         End
         Begin VB.Label lblWardNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "WardNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4500
            TabIndex        =   17
            Top             =   180
            Width           =   720
         End
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   " 보고서 출력예정 건수 :"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   255
      TabIndex        =   32
      Top             =   8760
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
      TabIndex        =   31
      Top             =   8715
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   " ☞ 출력대상자 리스트에서 선택하시면 출력 시 제외됩니다."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   255
      TabIndex        =   30
      Top             =   8520
      Width           =   5955
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   570
      Index           =   0
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8445
      Width           =   5190
   End
End
Attribute VB_Name = "frm421TubercleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()

Private objSql As New clsLISSqlTubercle
Private MyPatient   As New clsPatient   '환자 클래스
Private strWorkArea As String
Private MsgFg As Boolean
Private Enum TblColumn
    tcWDID = 2
    tcMAJDOCT
    tcPTID
    tcPTNM
    tcSAGE
    tcCNT
    tcSTSCD
    tcVFYDT
    tcRCVDT     '10
    tcATTR
    tcWARDID
    tcDEPTCD
    tcBACNM
    tcDLAW '15
    tcILAW
    tcDGROW
    tcRLAW
    tcALAW
    tcRGROW    '20
    tcREMARK
    tcWORKAREA
    tcACCDT
    tcACCSEQ
End Enum

Private Sub chkAllWard_Click()
    
    lblWardNm.Caption = ""
    txtWardId.Text = Choose(chkAllWard.Value + 1, "", CS_AllCaption)
    txtWardId.Enabled = Choose(chkAllWard.Value + 1, True, False)
End Sub

Private Sub chkAllDoct_Click()
    
    lblDoctNm.Caption = ""
    txtDoctId.Text = Choose(chkAllDoct.Value + 1, "", CS_AllCaption)
    txtDoctId.Enabled = Choose(chkAllDoct.Value + 1, True, False)
End Sub

Private Sub cmdClear_Click()
    TxtClear
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
    RaiseEvent FormClose
End Sub

Private Sub cmdPrint_Click()
    Dim objProgress     As jProgressBar.clsProgress
    Dim strWardId       As String
    Dim strTmp          As String
    Dim strSQL          As String
    Dim strFileNm       As String
    Dim strWA           As String
    Dim strAccDt        As String
    Dim strAccseq       As String
    Dim strRptDt        As String
    Dim strRptTm        As String
    Dim strRptId        As String
    Dim strRptNm        As String
    Dim strMyFile       As String
    Dim lngFNum         As Long
    Dim lngCnt          As Long
    Dim i               As Integer
    Dim j               As Integer
    
    
    
    If Printers.Count = 0 Then
        MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
        Exit Sub
    End If
    
    If Trim(lblCnt.Caption) = "0" Then MsgBox "출력 할 리스트가 없습니다.", vbCritical, "출력": Exit Sub

    
    strMyFile = Dir(InstallDir & "LIS\Rpt\CrystalReport.txt")
    
    If strMyFile = "" Then
        MsgBox "CrystalReport.txt 파일이 없습니다.", vbCritical, "정보확인"
        Exit Sub
    End If
    strMyFile = ""
    
    strFileNm = InstallDir & "LIS\Rpt\CrystalReport.txt"

    strMyFile = Dir(InstallDir & "LIS\Rpt\rptlab421.rpt")
    
    If strMyFile = "" Then
        MsgBox "rptlab421.rpt 파일이 없습니다.", vbCritical, "정보확인"
        Exit Sub
    End If
    
    strRptNm = InstallDir & "LIS\Rpt\rptlab421.rpt"
    
    Me.MousePointer = 13
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
    
    With tblOrder
        For i = 1 To .DataRowCnt '.MaxRows
            strTmp = ""
            objProgress.Value = i
            .Row = i
            .Col = 1
            If .Value = 0 Then
                .Col = TblColumn.tcPTID
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcPTNM
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcSAGE
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcWARDID
                    strWardId = Trim(.Value)
                .Col = TblColumn.tcDEPTCD
                    If strWardId = "" Then
                        strTmp = strTmp & Trim(.Value) & vbTab
                    Else
                        strTmp = strTmp & Trim(.Value) & Space(5) & "병동 :" & strWardId & vbTab
                    End If
                .Col = TblColumn.tcRCVDT
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcATTR
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcVFYDT
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcBACNM
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcDLAW
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcILAW
                    strTmp = strTmp & Trim(.Value) & vbTab
                    strTmp = strTmp & Trim("") & vbTab
                .Col = TblColumn.tcDGROW
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcRLAW
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcALAW
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcRGROW
                    strTmp = strTmp & Trim(.Value) & vbTab
                .Col = TblColumn.tcREMARK
                    strTmp = strTmp & Trim(.Value) & vbTab
                
                For j = 25 To .DataColCnt
                    .Col = j: strTmp = strTmp & Trim(.Value) & vbTab
                Next j
                
                strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

            
On Error GoTo ErrPrint
                lngFNum = FreeFile
                Open strFileNm For Output As #lngFNum
                Print #lngFNum, strTmp
                Close #lngFNum
                crtReport.ReportFileName = strRptNm
                crtReport.ParameterFields(0) = "title;임상병리/핵의학 검사보고서;true"
                crtReport.ParameterFields(1) = "add;" & P_HOSPITALADDR & ";true"
                crtReport.ParameterFields(2) = "HospNm;" & P_HOSPITALNAME & " 임상병리과;true"
                crtReport.RetrieveDataFiles
                crtReport.WindowState = 2 ' crptMaximized
                crtReport.Destination = crptToPrinter ' crptToWindow
                crtReport.Action = 1
                crtReport.Reset
                
                .Col = TblColumn.tcWORKAREA: strWA = .Value
                .Col = TblColumn.tcACCDT: strAccDt = .Value
                .Col = TblColumn.tcACCSEQ: strAccseq = .Value
                
                strRptDt = Format(GetSystemDate, "YYYYMMDD")
                strRptTm = Format(GetSystemDate, "mmhhss")
                strRptId = ObjMyUser.EmpId
                
                strSQL = objSql.SQLUpdateRptAFBHeader(strWA, strAccDt, strAccseq, strRptDt, strRptTm, strRptId)
                
On Error GoTo ErrUpdate
                DBConn.BeginTrans
                DBConn.Execute (strSQL)
                DBConn.CommitTrans
            End If
            
            
        Next
    End With
    Set objProgress = Nothing
    Me.MousePointer = 0
    TxtClear
    Exit Sub
    
ErrPrint:
    MsgBox "출력이 되지 않았습니다.", vbCritical
    Set objProgress = Nothing
    Me.MousePointer = 0
    Exit Sub
ErrUpdate:
    Set objProgress = Nothing
    Me.MousePointer = 0
    DBConn.RollbackTrans
    MsgBox "출력 도중 오류가 발생하였습니다.", vbExclamation
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
        strDept = Trim(medGetP(.Value, 1, "-"))
        For i = 1 To tblOrder.DataRowCnt
            MsgFg = True
            .Row = i
            .Col = 2
            If strDept = Trim(medGetP(.Value, 1, "-")) Then
                .Col = 1
                If .Value = 1 Then
                    lblCnt.Caption = Val(lblCnt.Caption) + 1
                End If
                .Value = lngButtonValue
                If .Value = 1 Then
                    lblCnt.Caption = Val(lblCnt.Caption) - 1
                End If
            End If
            MsgFg = False
        Next
    End With
   
End Sub

Private Sub cmdQuery_Click()
    Dim Rs          As New Recordset
    Dim rs1         As New Recordset
    Dim rs2         As New Recordset
    Dim objDisease  As New clsDisease
    Dim objProgress As clsProgress
    Dim strBussDiv  As String
    Dim strVfyDt    As String
    Dim strSex      As String
    Dim strLastDt   As String
    Dim strDeptCd  As String
    Dim strFieldNm  As String
    Dim strMajdoct  As String
    Dim strSQL      As String
    Dim strTmp      As String
    Dim strSAge     As String
    Dim strPtNm     As String
    Dim strAttr     As String
    Dim ii          As String
    Dim jj          As String
    
    If strWorkArea = "" Then MsgBox "항산성균 검사항목이 존재하지 않습니다.", vbCritical, "조회오류": GoTo Nodata
    strBussDiv = IIf(optBussDiv(0).Value = True, 1, 2)
    strVfyDt = Format(dtpVfyDt.Value, "YYYYMMDD")
    strLastDt = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    
    medClearTable tblOrder
    
    If strBussDiv = "1" Then
        strFieldNm = "c.deptcd"
        strDeptCd = Trim(txtWardId.Text)
    Else
        strFieldNm = "c.wardid"
        strDeptCd = Trim(txtWardId.Text)
        
    End If
    strMajdoct = Trim(txtDoctId.Text)
    If Trim(strDeptCd) = "(전체)" Then strDeptCd = ""
    If Trim(strMajdoct) = "(전체)" Then strMajdoct = ""
    
    '프로그래스바 생성..
    Me.MousePointer = 13
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
    
    
    If optPrint(0).Value = True Then
        strSQL = objSql.SQLGetAFBSensReport(strVfyDt, strLastDt, strWorkArea, strBussDiv, strFieldNm, strDeptCd, strMajdoct)
        tblOrder.ZOrder 0
    ElseIf optPrint(1).Value = True Then
        strSQL = objSql.SQLGetAFBSensReport(strVfyDt, strLastDt, strWorkArea, strBussDiv, strFieldNm, strDeptCd, strMajdoct, 2)
        tblOrder.ZOrder 0
    Else
        strSQL = objSql.SQLGetAFBSensPtIdReport(Trim(txtPtId.Text), strVfyDt, strWorkArea)
        tblOrder.ZOrder 0
    End If

    
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        Set objProgress = Nothing
        MsgBox "해당 데이타가 없습니다.", vbInformation, "결과지 출력"
        GoTo Nodata
    End If
    
'    Dim objEmp As clsBasisData
    Dim strEmp As String
    
    If Rs.RecordCount > 0 Then
         '프로그래스바 생성..
        objProgress.Max = Rs.RecordCount
        objProgress.Min = 0
        objProgress.Value = 0
        
        Rs.MoveFirst
        ii = 1
        With tblOrder
            .MaxRows = Rs.RecordCount
            lblCnt.Caption = .MaxRows
            .ReDraw = True
            Do Until Rs.EOF
                .Row = ii
                .Col = TblColumn.tcWDID
                    If optBussDiv(0).Value = True Then
                        .Value = Rs.Fields("deptcd").Value & ""
                    Else
                        .Value = Rs.Fields("wardid").Value & ""
                    End If
                .Col = TblColumn.tcMAJDOCT:
'                    Set objEmp = Nothing
'                    Set objEmp = New clsBasisData
                    strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                    Set objEmp = Nothing
                    
                    .Value = strEmp 'GetEmpName(rs.Fields("majdoct").Value & "")
                    If .Value = "" Then .Value = Rs.Fields("majdoct").Value & ""
                .Col = TblColumn.tcPTID: .Value = Rs.Fields("ptid").Value & ""
                
               
'                If MyPatient.PtntQuery(rs.Fields("ptid").Value & "") Then
                If MyPatient.GETPatient(Rs.Fields("ptid").Value & "") Then
                    strPtNm = MyPatient.PtNm
                    strSex = IIf(MyPatient.Sex = "M", "남", "여")
                    strSAge = strSex & "/" & MyPatient.Age
                    
                    '임상진단....
                    
                    objDisease.ptid = Rs.Fields("ptid").Value & ""
                    strAttr = objDisease.Disease
                End If
                    
                .Col = TblColumn.tcPTNM:    .Value = strPtNm
                .Col = TblColumn.tcSAGE:    .Value = strSAge
                .Col = TblColumn.tcCNT:     .Value = "1"
                .Col = TblColumn.tcSTSCD
                    Select Case Rs.Fields("stscd").Value & ""
                        Case StsCd_LIS_FinRst
                            .Value = "결과"
                        Case StsCd_LIS_Modify
                            .Value = "수정"
                    End Select
                        
                .Col = TblColumn.tcVFYDT:
                    .Value = Format(Mid(Rs.Fields("vfydt").Value & "", 3), "0#-##-##") & Space(1) & Format(Mid(Rs.Fields("vfytm").Value & "", 1, 4), "0#:0#")
                .Col = TblColumn.tcRCVDT:
                    .Value = Format(Mid(Rs.Fields("rcvdt").Value & "", 3), "0#-##-##") & Space(1) & Format(Mid(Rs.Fields("rcvtm").Value & "", 1, 4), "0#:0#")
                .Col = TblColumn.tcATTR:    .Value = Replace(strAttr, vbNewLine, Space(1))
                .Col = TblColumn.tcWARDID:
'                    Set objEmp = Nothing
'                    Set objEmp = New clsBasisData
                    strEmp = GetWardNm(Rs.Fields("wardid").Value & "")
'                    Set objEmp = Nothing
                    
                    If strEmp <> "" Then
                        lblWardNm.Caption = strEmp
                    Else
                        .Value = Trim(Rs.Fields("wardid").Value & "")
                    End If
                        
'                    If ObjLISComCode.WardID.Exists(Trim(rs.Fields("wardid").Value & "")) = True Then
'                        ObjLISComCode.WardID.KeyChange Trim(rs.Fields("wardid").Value & "")
'                        lblWardNm.Caption = ObjLISComCode.WardID.Fields("wardnm")
'                    Else
'                       .Value = Trim(rs.Fields("wardid").Value & "")
'                    End If
                    
                .Col = TblColumn.tcDEPTCD:
'                    Set objEmp = Nothing
'                    Set objEmp = New clsBasisData
                    strEmp = GetDeptNm(Rs.Fields("deptcd").Value & "")
'                    Set objEmp = Nothing
                    
                    If strEmp <> "" Then
                        .Value = strEmp
                    Else
                        .Value = Trim(Rs.Fields("deptcd").Value & "")
                    End If
                        
'                    If ObjLISComCode.DeptCd.Exists(Trim(rs.Fields("deptcd").Value & "")) = True Then
'                        ObjLISComCode.DeptCd.KeyChange Trim(rs.Fields("deptcd").Value & "")
'                        .Value = ObjLISComCode.DeptCd.Fields("deptnm")
'                    Else
'                        .Value = Trim(rs.Fields("deptcd").Value & "")
'                    End If
                .Col = TblColumn.tcBACNM:
                
                strTmp = objSql.SQLAFPCultureRstLoad(Rs.Fields("workarea").Value & "", Rs.Fields("accdt").Value & "", Rs.Fields("accseq").Value & "")
                
                Set rs2 = Nothing
                Set rs2 = New Recordset
                rs2.Open strTmp, DBConn
                
                If rs2.RecordCount > 0 Then
                    .Value = rs2.Fields("rstnm").Value & ""
                Else
                    .Value = Rs.Fields("bacrstcd").Value & ""
                End If
                Set rs2 = Nothing
                
                .Col = TblColumn.tcDLAW:    .Value = IIf(Rs.Fields("dilaw").Value & "" = "0", "√", "")
                .Col = TblColumn.tcILAW:    .Value = IIf(Rs.Fields("dilaw").Value & "" = "1", "√", "")
                .Col = TblColumn.tcDGROW:   .Value = Rs.Fields("dgrow").Value & ""
                .Col = TblColumn.tcRLAW:    .Value = IIf(Rs.Fields("ralaw").Value & "" = "0", "√", "")
                .Col = TblColumn.tcALAW:    .Value = IIf(Rs.Fields("ralaw").Value & "" = "1", "√", "")
                .Col = TblColumn.tcRGROW:   .Value = Rs.Fields("rgrow").Value & ""
                .Col = TblColumn.tcREMARK:  .Value = Replace(Rs.Fields("remark").Value & "", vbNewLine, Space(1))
                .Col = TblColumn.tcWORKAREA: .Value = Rs.Fields("workarea").Value & ""
                .Col = TblColumn.tcACCDT:   .Value = Rs.Fields("accdt").Value & ""
                .Col = TblColumn.tcACCSEQ:  .Value = Rs.Fields("accseq").Value & ""
                
                jj = 25
                strTmp = objSql.SQLAFPSensBodyLoad(Rs.Fields("workarea").Value & "", Rs.Fields("accdt").Value & "", Rs.Fields("accseq").Value & "")
                
                Set rs1 = Nothing
                Set rs1 = New Recordset
                
                rs1.Open strTmp, DBConn
                
                If rs1.RecordCount > 0 Then
                    rs1.MoveFirst
                    Do Until rs1.EOF
                        .Col = jj
                            .Value = rs1.Fields("rstvalue").Value & ""
                        jj = jj + 1
                        rs1.MoveNext
                    Loop
                End If
                Set rs1 = Nothing
                
                ii = ii + 1
                objProgress.Value = objProgress.Value + 1
                Rs.MoveNext
            Loop
            Set objProgress = Nothing
            .ReDraw = False
        End With
    End If
    
Nodata:
    Me.MousePointer = 0
    Set rs2 = Nothing
    Set rs1 = Nothing
    Set Rs = Nothing
    Set objDisease = Nothing
End Sub

Private Sub cmdWardList_Click()
'% 병동코드 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
    Dim strCaption As String
    Dim strHead As String
'    Dim objDept As clsBasisData
    
    chkAllWard.Value = 0
    
    If optBussDiv(0).Value Then
        strCaption = "진료과 조회"
        strHead = "부서코드;부서명"
    Else
        strCaption = "병동 조회"
        strHead = "병동코드;병동명"
    End If
    
'    Set objDept = New clsBasisData
    
    With objMyList
        .Connection = DBConn
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
    
'    Set objDept = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdDoctList_Click()

'% 주치의 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
'    Dim objDoct As New clsBasisData
    
    chkAllDoct.Value = 0
    
    With objMyList
        .Connection = DBConn
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

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
    Set MyPatient = Nothing
End Sub

Private Sub txtDoctId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDoctId_LostFocus()
'    Dim objEmp As clsBasisData
    
    If Trim(txtDoctId.Text) = "" Then Exit Sub
    
'    Set objEmp = New clsBasisData
    
    lblDoctNm.Caption = GetEmpNm(txtDoctId.Text) ' GetEmpName(txtDoctId.Text)
    If lblDoctNm.Caption = "" Then
        txtDoctId.Text = ""
        lblDoctNm.Caption = ""
    End If
'    Set objEmp = Nothing
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    
    Dim objPatient As New clsPatient     '환자 클래스
    
    If Not gUsingInWardMenu Then

        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If Screen.ActiveControl Is Nothing Then Exit Sub
        
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    End If
    
'    If MsgFg Then Exit Sub
      
    If txtPtId.Text = "" Then
        'txtPtId.SetFocus
        Exit Sub
    End If
    
    
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With objPatient
'        If Trim(txtPtId.Text) <> "" And .PtntQuery(txtPtId.Text) Then
        If Trim(txtPtId.Text) <> "" And .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSexAge.Caption = .SEXNM & " / " & .Age & " " & .AGEDIV
            If .WardID = "" Then
                lblWard.Caption = ""
            Else
                lblWard.Caption = .WardID & "-" & .ROOMID
            End If
'            PtFg = True
'            ClearFg = False
        Else
            If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
'            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요..", vbInformation
            txtPtId.SetFocus
'            MsgFg = False
'            PtFg = False
            Set objPatient = Nothing
            Exit Sub
        End If
    End With
    
    Set objPatient = Nothing

    Exit Sub

End Sub

Private Sub tblOrder_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i           As Long
    Static lngOnOff As Long
    
    If Row = 0 And Col = 1 Then
        lngOnOff = (lngOnOff + 1) Mod 2
        For i = 1 To tblOrder.MaxRows
            tblOrder.Row = i
            tblOrder.Col = 1
            tblOrder.Value = lngOnOff
        Next
        lblCnt.Caption = IIf(lngOnOff = 1, 0, tblOrder.DataRowCnt)
    End If
    
End Sub

Private Sub Form_Load()
    lblWardNm.Caption = ""
    lblWard.Caption = ""
    
    TxtClear
    
    GetWorkArea
End Sub

Private Sub GetWorkArea()
    Dim objMySql As New clsLISSqlMasters
    Dim Rs      As New Recordset
    Dim strSQL  As String
    
    strWorkArea = ""
    
    strSQL = objMySql.SqlItemQuery(P_AFBSENSCD)
    
    Rs.Open strSQL, DBConn
    If Rs.RecordCount > 0 Then
        strWorkArea = Rs.Fields("workarea").Value & ""
    End If
    
Nodata:
    Set Rs = Nothing
    Set objMySql = Nothing
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

    dtpVfyDt.Value = GetSystemDate
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    '결과지 출력예정리스트
    medClearTable tblOrder

    lblCnt.Caption = 0

End Sub

Private Sub TxtClear()
    '결과지 출력 조건
    dtpVfyDt.Value = GetSystemDate

    '결과지 출력예정리스트
    medClearTable tblOrder
    
    lblWard.Caption = ""
    tblOrder.MaxRows = 0
'    tblOrdSheet.MaxRows = 0
    tblOrder.ZOrder 0
'    chkAll.Value = 1
    txtWardId.Text = "(전체)"
    lblCnt.Caption = 0
    chkAllWard.Value = 1
    txtDoctId.Text = "(전체)"
    chkAllDoct.Value = 1
    txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtWardId_LostFocus()
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    If Trim(txtWardId.Text) = "" Then Exit Sub
    
'    Set objDept = New clsBasisData
    
    If optBussDiv(0).Value Then
        strDept = GetDeptNm(txtWardId.Text)
        
        If strDept <> "" Then
            lblWardNm.Caption = strDept
        Else
            txtWardId.Text = ""
            lblWardNm.Caption = ""
        End If
    Else
        strDept = GetWardNm(txtWardId.Text)
        
        If strDept <> "" Then
            lblWardNm.Caption = strDept
        Else
            txtWardId.Text = ""
            lblWardNm.Caption = ""
        End If
    End If
'    Set objDept = Nothing
    
'    If optBussDiv(0).Value = True Then
'        If ObjLISComCode.DeptCd.Exists(txtWardId.Text) = True Then
'            ObjLISComCode.DeptCd.KeyChange txtWardId.Text
'            lblWardNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'        Else
'            txtWardId.Text = ""
'            lblWardNm.Caption = ""
'        End If
'    Else
'        If ObjLISComCode.WardID.Exists(txtWardId.Text) = True Then
'            ObjLISComCode.WardID.KeyChange txtWardId.Text
'            lblWardNm.Caption = ObjLISComCode.WardID.Fields("wardnm")
'        Else
'            txtWardId.Text = ""
'            lblWardNm.Caption = ""
'        End If
'    End If
End Sub
