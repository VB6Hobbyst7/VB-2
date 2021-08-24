VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm463EmmaList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "EMMAList"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   14610
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   Tag             =   "45500"
   WindowState     =   2  '최대화
   Begin VB.Frame frmSMS 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SMS전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   5250
      TabIndex        =   17
      Top             =   1230
      Width           =   4545
      Begin VB.CommandButton cmdTrans 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전송"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1680
         Style           =   1  '그래픽
         TabIndex        =   28
         Tag             =   "135"
         Top             =   3840
         Width           =   1320
      End
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   3030
         Style           =   1  '그래픽
         TabIndex        =   27
         Tag             =   "135"
         Top             =   3840
         Width           =   1320
      End
      Begin VB.TextBox txtTransId 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "opt"
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtTransNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "opt"
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtTransNo 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "opt"
         Top             =   630
         Width           =   3195
      End
      Begin VB.TextBox txtDtId 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "opt"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtDtNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "opt"
         Top             =   1020
         Width           =   1875
      End
      Begin VB.TextBox txtDetpCd 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "opt"
         Top             =   1350
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDeptNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "opt"
         Top             =   1350
         Width           =   1875
      End
      Begin VB.TextBox txtDtNo 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "opt"
         Top             =   1680
         Width           =   3195
      End
      Begin VB.TextBox txtTransDt 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   18
         Tag             =   "opt"
         Top             =   3270
         Width           =   3195
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   7
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "전송자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1005
         Index           =   8
         Left            =   180
         TabIndex        =   30
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1773
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
         Caption         =   "수신자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   31
         Top             =   2070
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Caption         =   "메시지"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   32
         Top             =   3300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Caption         =   "전송일시"
         Appearance      =   0
      End
      Begin RichTextLib.RichTextBox rtfMessage 
         Height          =   1170
         Left            =   1140
         TabIndex        =   33
         Top             =   2070
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2064
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis463.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   11
         Left            =   180
         TabIndex        =   34
         Top             =   630
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
   End
   Begin VB.CheckBox chkWorkArea 
      Appearance      =   0  '평면
      BackColor       =   &H8000000D&
      Caption         =   "WorkArea"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Top             =   180
      Width           =   1275
   End
   Begin VB.ComboBox cboWorkArea 
      Height          =   300
      Left            =   5490
      Style           =   2  '드롭다운 목록
      TabIndex        =   15
      Top             =   180
      Width           =   2100
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "To &Excel"
      Height          =   510
      Left            =   11760
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "127"
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "132"
      Top             =   60
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "검 색 (&Q)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "158"
      Top             =   60
      Width           =   1320
   End
   Begin VB.Frame fraInOut 
      BackColor       =   &H00DBE6E6&
      Height          =   465
      Left            =   7815
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   3960
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "모두"
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   750
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과수정사유"
         Height          =   315
         Index           =   1
         Left            =   870
         TabIndex        =   8
         Top             =   120
         Width           =   1470
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수취소사유"
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread ssCmtList 
      Height          =   7740
      Left            =   75
      TabIndex        =   2
      Tag             =   "45506"
      Top             =   660
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   13653
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   14
      OperationMode   =   1
      Protect         =   0   'False
      ShadowColor     =   14737632
      SpreadDesigner  =   "Lis463.frx":009D
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   375
      Left            =   1050
      TabIndex        =   3
      Top             =   150
      Width           =   1425
      _ExtentX        =   2514
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
      Format          =   62324739
      CurrentDate     =   36328
   End
   Begin MSComCtl2.DTPicker dtpEndDt 
      Height          =   360
      Left            =   2730
      TabIndex        =   4
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
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
      Format          =   62324739
      CurrentDate     =   36328
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   150
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   661
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "검색기간"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   360
      Index           =   1
      Left            =   6870
      TabIndex        =   10
      Top             =   150
      Visible         =   0   'False
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "조회유형"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   750
      _Version        =   196608
      _ExtentX        =   1323
      _ExtentY        =   1191
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Lis463.frx":1D46
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      Height          =   240
      Left            =   2520
      TabIndex        =   0
      Top             =   225
      Width           =   270
   End
End
Attribute VB_Name = "frm463EmmaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Event LastFormUnload()
Private objRst As New clsPatientInfo
Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

Private Sub cmdCancle_Click()
    frmSMS.Visible = False
    txtTransId.Text = ""
    txtTransNm.Text = ""
    txtTransNo.Text = ""
    txtDtNm.Text = ""
    txtDtId.Text = ""
    txtDetpCd.Text = ""
    txtDeptNm.Text = ""
    txtDtNo.Text = ""
    rtfMessage.Text = ""
    txtTransDt.Text = ""
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp  As String
    
    If ssCmtList.DataRowCnt = 0 Then Exit Sub
    
    With ssCmtList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = Trim(strTmp)
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "EmmaList"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdStart_Click()

    Dim strSQL      As String
    Dim objProBar   As jProgressBar.clsProgress
    Dim rsGetinfo   As Recordset
    Dim RS          As Recordset
    Dim sStartDt    As String
    Dim SendDt      As String
    Dim i%
    Dim strTransNm  As String
    Dim strTAT      As String
    Dim strHH       As String
    Dim strDD       As String
    Dim strTATDt    As String
    Dim strTransFg  As String
    Dim strWorkArea As String
    Dim strSDate     As String
    Dim strEDate     As String
    Dim strNDate     As String
    Dim strEditDt    As String
    
    On Error Resume Next
    
'    sStartDt = Format(dtpStartDt.Value, CS_DateDbFormat)
'    SendDt = Format(dtpEndDt.Value, CS_DateDbFormat)

    sStartDt = Format(dtpStartDt.Value, "YYYY-MM-DD 00:00:00")
    SendDt = Format(dtpEndDt.Value, "YYYY-MM-DD 24:00:00")
    
    strSDate = Format(dtpStartDt.Value, "YYYY-MM-DD")
    strEDate = Format(dtpEndDt.Value, "YYYY-MM-DD")
    strNDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
'    strSQL = ""
'    strSQL = strSQL & vbLf & "SELECT A.*, B.RECVDATE FROM S2COM102 A, MDNOTIFT B"
'    strSQL = strSQL & vbLf & " WHERE TRANSDT >= '" & sStartDt & "'"
'    strSQL = strSQL & vbLf & "   AND TRANSDT <= '" & SendDt & "'"
'    strSQL = strSQL & vbLf & "   AND A.REMARK = B.WORKAREA(+)"

    strSQL = ""
    strSQL = strSQL & vbLf & "SELECT * FROM S2COM102 "
    strSQL = strSQL & vbLf & " WHERE TRANSDT >= '" & sStartDt & "'"
    strSQL = strSQL & vbLf & "   AND TRANSDT <= '" & SendDt & "'"
    
    If chkWorkArea.Value = 1 Then
        strSQL = strSQL & vbLf & "   AND SUBSTR(REMARK,1,2) = '" & Mid(cboWorkArea, 1, 2) & "'"
    End If
    
'    If optOption(1).Value Then
'        strSQL = strSQL & " AND  exists (SELECT * FROM " & T_LAB308 & _
'                                        " WHERE workarea = a.workarea " & _
'                                        " AND   accdt = a.accdt " & _
'                                        " AND   accseq = a.accseq ) "
'    ElseIf optOption(2).Value Then
'        strSQL = strSQL & " AND   a.stscd = '" & enStsCd.StsCd_LIS_Cancel & "'"
'    End If
                                        
    Set objProBar = New jProgressBar.clsProgress
    
    With objProBar
        .Container = Me
        .Width = ssCmtList.Width
        .Left = ssCmtList.Left
        .Top = ssCmtList.Top - 280
        .Height = 280
        .Message = "자료를 읽기 위해 준비중입니다..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = ssCmtList.Width
'        .XPos = ssCmtList.Left
'        .YPos = ssCmtList.Top - 280
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "자료를 읽기 위해 준비중입니다..."
'        .Value = 1
    End With
    
    Set rsGetinfo = New Recordset
    rsGetinfo.Open strSQL, DBConn
    
'    objProBar.Msg = ""
    
    
    If rsGetinfo.RecordCount > 0 Then
        objProBar.Max = rsGetinfo.RecordCount
    Else
        MsgBox "데이타가 없습니다.."
    End If
'    barStatus.Value = 0
    '.Fields("statfg").Value
    
    ClearssCmtList
    For i = 1 To rsGetinfo.RecordCount
'        barStatus.Value = barStatus.Value + 1
        objProBar.Value = i
        DoEvents
        With rsGetinfo
            strTransNm = GetEmpNm(Trim("" & .Fields("TRANSID").Value))
'            If Trim("" & .Fields("RCVDT").Value) = "" Then
'                strTAT = ""
'                strTATDt = ""
'            Else
'                strTAT = DateDiff("n", "" & .Fields("RCVDT").Value, "" & .Fields("TRANSDT").Value)
'                If strTAT > 60 Then
'                    strHH = strTAT \ 60
'                    strDD = strTAT - (strHH * 60)
'                    strTATDt = strHH & "시간" & " " & strDD & "분"
'                Else
'                    strTATDt = strTAT & "분"
'                End If
'            End If
            
            strTAT = DateDiff("n", "" & .Fields("TRANSDT").Value, strNDate)
            If strTAT > 60 Then
                strHH = strTAT \ 60
                strDD = strTAT - (strHH * 60)
                strTATDt = strHH & "시간" & " " & strDD & "분"
            Else
                strTATDt = strTAT & "분"
            End If

            strWorkArea = "" & .Fields("REMARK").Value
'            If strWorkArea <> "" Then
            strSQL = ""
            strSQL = " SELECT * FROM MDNOTIFT WHERE  notidate between to_date('" & strSDate & "','yyyy-mm-dd') and to_date('" & strEDate & "','yyyy-mm-dd') AND notitype = '7' and workarea = '" & strWorkArea & "' "
            Set RS = New Recordset
            RS.Open strSQL, DBConn
            
            strEditDt = "" & RS.Fields("EDITDATE").Value
            
            If RS.RecordCount = 0 Then
                strTransFg = "확인"
            Else
                If RS.Fields("RECVDATE").Value & "" = "" Then
'                    If Mid(cboWorkArea, 1, 2) = "05" Then
'                        strTransFg = "확인"
'                    Else
                        strTransFg = "미확인"
'                    End If
                Else
                    strTransFg = "확인"
                End If
            End If
            
            Dim strTestCd As String
            Dim varTestCd As Variant
            Dim strTmpDoctNm As String
            
            If .Fields("DOCTNM").Value & "" = "" Then
                strTmpDoctNm = GetEmpNm(.Fields("DOCTID").Value & "")
                If strTmpDoctNm = "" Then
                    strTmpDoctNm = .Fields("DOCTID").Value & ""
                End If
            Else
                strTmpDoctNm = .Fields("DOCTNM").Value & ""
            End If
            
            strTestCd = ""
            
            varTestCd = Split("" & .Fields("TRANSMSG"), vbCrLf)
'            If InStr(varTestCd(1), "Critical") > 0 Then
'                strTestCd = Trim(Mid(varTestCd(2), 1, InStr(varTestCd(2), ":") - 1))
'            Else
                strTestCd = Trim(Mid(varTestCd(1), 1, InStr(varTestCd(1), ":") - 1))
'            End If

            If strTestCd = "" Then
                varTestCd = Split("" & .Fields("TRANSMSG"), vbCr)
                strTestCd = Trim(Mid(varTestCd(1), 1, InStr(varTestCd(1), ":") - 1))
            End If
            
            Call DspSpd_New2("" & .Fields("TRANSDT").Value, "" & .Fields("REMARK").Value, strTransNm, "" & strTmpDoctNm & " / " & "" & .Fields("DEPTNM").Value & " (" & "" & .Fields("TELNO").Value & ")", _
                        "" & .Fields("TRANSMSG").Value, strTransFg, "" & .Fields("RCVSTAT").Value, strTATDt, "" & strTmpDoctNm, "" & .Fields("DOCTNM").Value, "" & .Fields("TELNO").Value, "" & .Fields("DEPTNM").Value, "", strTestCd & "/" & .Fields("TESTCD"), strEditDt, i)
        
'            Call DspSpd_New2("" & .Fields("TRANSDT").Value, "" & .Fields("REMARK").Value, strTransNm, "" & strTmpDoctNm & " / " & "" & .Fields("DEPTNM").Value & " (" & "" & .Fields("TELNO").Value & ")", _
'                        "" & .Fields("TRANSMSG").Value, strTransFg, "" & .Fields("RCVSTAT").Value, strTATDt, "" & strTmpDoctNm, "" & .Fields("DOCTNM").Value, "" & .Fields("TELNO").Value, "" & .Fields("DEPTNM").Value, "", strTestCd, strEditDt, i)
        
        End With
      rsGetinfo.MoveNext
    Next i
    
'    MouseDefault   2001/04/18
    
    Set rsGetinfo = Nothing
    Set objProBar = Nothing
    
End Sub

Private Sub cmdTrans_Click()
    Dim ServerName   As String
    Dim DatabaseName As String
    Dim UserName     As String
    Dim Password     As String
    Dim strTransCd   As String
    Dim strDoctCd    As String
    Dim strTransDt   As String
    Dim strTransStatus As String
    Dim strTansEtc   As String
    Dim strMessage   As String
    Dim strTransNo   As String
    Dim strDoctNo    As String
    Dim strSQL       As String
    Dim strDeptNm    As String
    Dim strTranNm    As String
    Dim strSMSIP     As String
    Dim strRcvDt     As String
    Dim strBackNo    As String
    
    Set AdoCn_ORACLE = New ADODB.Connection
    
    On Error Resume Next    '2013-09-11 PSK
    
    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
        .Properties("Persist Security Info") = True
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        .Open
    End With
           
    Set AdoRs_ORACLE = New ADODB.Recordset
        
    strSQL = ""
    strSQL = "SELECT * FROM S2lab032  "
    strSQL = strSQL + " WHERE cdindex = 'C232'"
    strSQL = strSQL + "   AND cdval1 = 'SVR1'  "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open strSQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            strSMSIP = AdoRs_ORACLE.Fields("FIELD4") & ""
        Else
            strSMSIP = "172.16.200.37"
        End If
        .Close
    End With
    
    Set AdoCn_SQL = New ADODB.Connection

    ServerName = strSMSIP
    DatabaseName = "medicalCRM_jesus"
    UserName = "jesus"
    Password = "jesus"
   
    With AdoCn_SQL
        .ConnectionTimeout = 10
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = ServerName
        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("User ID").Value = UserName
        .Properties("Password").Value = Password
        Screen.MousePointer = vbHourglass
        .Open
    End With
    Screen.MousePointer = vbDefault
    
    If txtDtNo.Text = "" Then
        MsgBox "수신번호를 입력하세요.", vbCritical + vbOKOnly, "수신번호등록 Message"
        txtDtNo.SetFocus
        Exit Sub
    End If
    
    strTransCd = ObjSysInfo.EmpId
    strTransNo = txtTransNo.Text
    strDoctCd = txtDtId.Text
    strTransDt = Format(Now, "YYYY-MM-DD HH:MM:SS")
    strDoctNo = txtDtNo.Text
    strTransStatus = "1"
    strTansEtc = "LIS"
    strDeptNm = txtDeptNm.Text
    strTranNm = txtTransNm.Text
    strMessage = rtfMessage.Text '& vbCrLf & "- " & strTranNm
    strBackNo = "063-230-8753"
    
    If Len(strMessage) > 80 Then
        MsgBox "메시지의 크기를 줄여주세요.", vbCritical + vbOKOnly, "메시지내용수정 Message"
        rtfMessage.SetFocus
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
    strSQL = strSQL & " values('" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & strBackNo & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransStatus & "' ,"
    strSQL = strSQL & "        '" & strTansEtc & "')"
    
    AdoCn_SQL.Execute strSQL
    
'    strSQL = ""
'    strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT)"
'    strSQL = strSQL & " values('" & strTransDt & "' ,"
'    strSQL = strSQL & "        '" & strTransCd & "' ,"
'    strSQL = strSQL & "        '" & strDoctNo & "' ,"
'    strSQL = strSQL & "        '" & Trim(txtDtNm.Text) & "' ,"
'    strSQL = strSQL & "        '' ,"
'    strSQL = strSQL & "        '" & strDeptNm & "' ,"
'    strSQL = strSQL & "        '" & strMessage & "' ,"
'    strSQL = strSQL & "        '정상' ,"
'    strSQL = strSQL & "        '" & strTransNo & "',"
'    strSQL = strSQL & "        '" & strRcvDt & "')"
'
'    AdoCn_ORACLE.Execute strSQL
    
'     strSQL = ""
'    strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
'    strSQL = strSQL & " (select '" & strDoctCd & "' ,"
'    strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
'    strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
'    strSQL = strSQL & "        '7' ,"
'    strSQL = strSQL & "        SYSDATE ,"
'    strSQL = strSQL & "        '[CVR(이상결과보고)]' ,"
'    strSQL = strSQL & "        '" & strMessage & "' ,"
'    strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strDoctCd & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
'
'    AdoCn_ORACLE.Execute strSQL
    
    strRcvDt = ""
    
    frmSMS.Visible = False
    Set AdoCn_SQL = Nothing
    Set AdoCn_ORACLE = Nothing
    
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim strWA As String
    
    dtpStartDt.Value = Now
    dtpEndDt.Value = Now
    ClearssCmtList
    optOption(0).Value = True
    frmSMS.Visible = False
    
    Call objRst.Load_WorkArea(cboWorkArea)
    
    '설정된 Workarea 가 있는경우 읽기
    
    strWA = GetSetting("Schweitzer2000 LIS", "Options", "UnvfyForWA", vbNullString)
    
    If strWA <> vbNullString Then
        cboWorkArea.ListIndex = Val(strWA)
    Else
        cboWorkArea.ListIndex = 0
    End If
End Sub

Private Sub DspSpd(ByVal TRANSDT As String, ByVal REMARK As String, ByVal TRANSID As String, ByVal DOCTID As String, ByVal TRANSMSG As String, ByVal RCVSTAT As String, _
                   ByVal strTAT As String, ByVal strTmp As String, ByVal lngRow As Long)
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    Dim lngMaxHeight As Long

    With ssCmtList
        .MaxRows = lngRow
        .Row = lngRow
        .Col = 1: .Text = Trim(TRANSDT)
        .Col = 2: .Text = Trim(REMARK)
        .Col = 3: .Text = Trim(TRANSID)
        .Col = 4: .Text = Trim(DOCTID)
        .Col = 5: .Text = Trim(TRANSMSG)
        .Col = 6: .Text = Trim(RCVSTAT)
        .Col = 7: .Text = Trim(strTAT)
        
'        If .MaxTextCellHeight > lngMaxHeight Then lngMaxHeight = .MaxTextCellHeight
'
'        .Col = 7: .Text = Trim(EmpNm)
'        .Col = 8: .Text = Format(Trim(RcvDt), CS_DateLongMask)
'        .Col = 9: .Text = Format(Trim(RcvTM), CS_TimeLongMask)
'        .Col = 10: .Text = Trim(OrdDocT)
'        .RowHeight(lngRow) = lngMaxHeight
    End With
    
End Sub

Private Sub DspSpd_New(ByVal TRANSDT As String, ByVal REMARK As String, ByVal TRANSID As String, ByVal DOCTID As String, ByVal TRANSMSG As String, ByVal TRANSSTAT As String, ByVal RCVSTAT As String, _
                   ByVal strTAT As String, ByVal strTmp As String, ByVal lngRow As Long)
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    Dim lngMaxHeight As Long

    With ssCmtList
        .MaxRows = lngRow
        .Row = lngRow
        .RowHeight(lngRow) = 50
        .Col = 1: .Text = Trim(TRANSDT)
        .Col = 2: .Text = Trim(REMARK)
        .Col = 3: .Text = Trim(TRANSID)
        .Col = 4: .Text = Trim(DOCTID)
        .Col = 5: .Text = Trim(TRANSMSG)
        .Col = 6: .Text = Trim(RCVSTAT)
        .Col = 7: .Text = Trim(TRANSSTAT)
        .Col = 8: .Text = Trim(strTAT)
        
'        If .MaxTextCellHeight > lngMaxHeight Then lngMaxHeight = .MaxTextCellHeight
'
'        .Col = 7: .Text = Trim(EmpNm)
'        .Col = 8: .Text = Format(Trim(RcvDt), CS_DateLongMask)
'        .Col = 9: .Text = Format(Trim(RcvTM), CS_TimeLongMask)
'        .Col = 10: .Text = Trim(OrdDocT)
'        .RowHeight(lngRow) = lngMaxHeight
    End With
    
End Sub

Private Sub DspSpd_New1(ByVal TRANSDT As String, ByVal REMARK As String, ByVal TRANSID As String, ByVal DOCTID As String, ByVal TRANSMSG As String, ByVal TRANSSTAT As String, ByVal RCVSTAT As String, _
                   ByVal strTAT As String, ByVal strDoctId As String, ByVal strDoctNm As String, ByVal strTelNo As String, ByVal strDeptNm As String, ByVal strTmp As String, ByVal lngRow As Long)
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    Dim lngMaxHeight As Long
    Dim tmpTAT       As String
    Dim varTmp
    Dim tmpMessage   As String
    
    With ssCmtList
        .MaxRows = lngRow
        .Row = lngRow
        .RowHeight(lngRow) = 50
        .Col = 1: .Text = Trim(TRANSDT)
        .Col = 2: .Text = Trim(REMARK)
        .Col = 3: .Text = Trim(TRANSID)
        .Col = 4: .Text = Trim(DOCTID)
        .Col = 5: .Text = Trim(TRANSMSG): tmpMessage = Trim(TRANSMSG)
        .Col = 6: .Text = Trim(RCVSTAT)
        .Col = 7: .Text = Trim(TRANSSTAT)
        .Col = 8: .Text = Trim(strTAT)
        
        .GetText 7, lngRow, varTmp
        
        tmpTAT = Replace(strTAT, "시간", "")
        tmpTAT = Replace(tmpTAT, "분", "")
        tmpTAT = Replace(tmpTAT, " ", "")
        
        If varTmp = "미확인" Then
            If tmpTAT > 30 Then
                .Col = 7
                .BackColor = vbRed
            Else
                .Col = 7
                .BackColor = vbWhite
            End If
        End If
        
        If InStr(tmpMessage, "VRE") > 0 Then
            .Col = 7
            .Value = "VRE"
            .BackColor = vbGreen
        End If
        
        If InStr(tmpMessage, "AFB Stain(형광법)") > 0 Then
            .Col = 7
            .Value = "AFB"
            .BackColor = vbCyan
        End If
                
        If InStr(tmpMessage, "AFB Stain(집균형광법)") > 0 Then
            .Col = 7
            .Value = "AFB"
            .BackColor = vbCyan
        End If
        
        .Col = 9: .Text = Trim(strDoctId)
        .Col = 10: .Text = Trim(strDoctNm)
        .Col = 11: .Text = Trim(strTelNo)
        .Col = 12: .Text = Trim(strDeptNm)
        
'        If .MaxTextCellHeight > lngMaxHeight Then lngMaxHeight = .MaxTextCellHeight
'
'        .Col = 7: .Text = Trim(EmpNm)
'        .Col = 8: .Text = Format(Trim(RcvDt), CS_DateLongMask)
'        .Col = 9: .Text = Format(Trim(RcvTM), CS_TimeLongMask)
'        .Col = 10: .Text = Trim(OrdDocT)
'        .RowHeight(lngRow) = lngMaxHeight
    End With
    
End Sub

Private Sub DspSpd_New2(ByVal TRANSDT As String, ByVal REMARK As String, ByVal TRANSID As String, ByVal DOCTID As String, ByVal TRANSMSG As String, ByVal TRANSSTAT As String, ByVal RCVSTAT As String, _
                   ByVal strTAT As String, ByVal strDoctId As String, ByVal strDoctNm As String, ByVal strTelNo As String, ByVal strDeptNm As String, ByVal strTmp As String, ByVal strTestNm As String, ByVal strEditDt As String, ByVal lngRow As Long)
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    Dim lngMaxHeight As Long
    Dim tmpTAT       As String
    Dim varTmp
    Dim tmpMessage   As String
    
    With ssCmtList
        .MaxRows = lngRow
        .Row = lngRow
        .RowHeight(lngRow) = 50
        .Col = 1: .Text = Trim(TRANSDT)
        .Col = 2: .Text = Trim(REMARK)
        .Col = 3: .Text = Trim(TRANSID)
        .Col = 4: .Text = Trim(DOCTID)
        .Col = 5: .Text = Trim(TRANSMSG): tmpMessage = Trim(TRANSMSG)
        .Col = 6: .Text = Trim(RCVSTAT)
        .Col = 7: .Text = Trim(TRANSSTAT)
        .Col = 8: .Text = Trim(strTAT)
        
        .GetText 7, lngRow, varTmp
        
        tmpTAT = Replace(strTAT, "시간", "")
        tmpTAT = Replace(tmpTAT, "분", "")
        tmpTAT = Replace(tmpTAT, " ", "")
        
        If varTmp = "미확인" Then
            If tmpTAT > 30 Then
                .Col = 7
                .BackColor = vbRed
            Else
                .Col = 7
                .BackColor = vbWhite
            End If
        End If
        
        If InStr(tmpMessage, "VRE") > 0 Then
            .Col = 7
            .Value = "VRE"
            .BackColor = vbGreen
        End If
        
        If InStr(tmpMessage, "CRE") > 0 Then
            .Col = 7
            .Value = "CRE"
            .BackColor = vbGreen
        End If
        
        If InStr(tmpMessage, "Strep A") > 0 Then
            .Col = 7
            .Value = "Strep A"
            .BackColor = vbGreen
        End If
        
        If InStr(tmpMessage, "CPE") > 0 Then
            .Col = 7
            .Value = "CPE PCR"
            .BackColor = vbGreen
        End If
        
        If InStr(tmpMessage, "AFB Stain(형광법)") > 0 Then
            .Col = 7
            .Value = "AFB"
            .BackColor = vbCyan
        End If
                
        If InStr(tmpMessage, "AFB Stain(집균형광법)") > 0 Then
            .Col = 7
            .Value = "AFB"
            .BackColor = vbCyan
        End If
        
        .Col = 9: .Text = Trim(strDoctId)
        .Col = 10: .Text = Trim(strDoctNm)
        .Col = 11: .Text = Trim(strTelNo)
        .Col = 12: .Text = Trim(strDeptNm)
        .Col = 13: .Text = Trim(strTestNm)
        .Col = 14: .Text = Trim(strEditDt)
'        If .MaxTextCellHeight > lngMaxHeight Then lngMaxHeight = .MaxTextCellHeight
'
'        .Col = 7: .Text = Trim(EmpNm)
'        .Col = 8: .Text = Format(Trim(RcvDt), CS_DateLongMask)
'        .Col = 9: .Text = Format(Trim(RcvTM), CS_TimeLongMask)
'        .Col = 10: .Text = Trim(OrdDocT)
'        .RowHeight(lngRow) = lngMaxHeight
    End With
    
End Sub

Private Sub ClearssCmtList()

    With ssCmtList
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With

End Sub

Private Sub dtpEndDt_Validate(Cancel As Boolean)
    ClearssCmtList
End Sub

Private Sub dtpStartDt_Validate(Cancel As Boolean)
    ClearssCmtList
End Sub


Private Sub optOption_Click(Index As Integer)
    ClearssCmtList
End Sub


Private Sub AnalysisHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    strTmp = "AnalysisList"
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 8

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("AnalysisList", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    strTmp = "조회기간 : " & Format(dtpStartDt.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpEndDt.Value, "YYYY년 MM월 DD일")
    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
    
    strTmp = "조회조건 : "
    If chkWorkArea.Value = 0 Then
        strTmp = strTmp & "     " & "(√)모두"
        strTmp = strTmp & "     " & "(  )WorkArea"
    Else
        strTmp = strTmp & "     " & "(  )모두"
        strTmp = strTmp & "     " & "(√)WorkArea"
    End If
    
    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    Call PrintString("전송일자", "접수번호", "발신자", "수신자", "Message", True)
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub
Private Sub PrintString(ByVal sAccno As String, ByVal sPtid As String, ByVal sPtnm As String, ByVal sSexAge As String, ByVal sMesg As String, _
                        Optional ByVal blnHead As Boolean = False)
    Dim arytmp()    As String
    Dim strTmp      As String
    Dim ii          As Integer
    
    
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call AnalysisHead
    End If
    
    Call Print_Setting(sAccno, PrtLeft, LineSpace, 30, "C", "C", False)
    Call Print_Setting(sPtid, 40, LineSpace, 20, "C", "C", False)
    Call Print_Setting(sPtnm, 65, LineSpace, 10, "C", "C", False)
    Call Print_Setting(sSexAge, 78, LineSpace, 40, "L", "C", False)
'    Call Print_Setting(sMesg, 75, LineSpace, 100, "L", "C", False)
    
    If sMesg <> "" Then
        If blnHead = True Then
            Call Print_Setting(sMesg, 135, LineSpace, 150, "L", "C")
        Else
'            Printer.FontBold = True
            For ii = 1 To 5
                If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                    sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                    If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                        sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                    End If
                End If
            Next

            arytmp() = Split(Trim(sMesg), vbCrLf)
            For ii = LBound(arytmp) To UBound(arytmp)
                If lngCurYPos > Printer.ScaleHeight - 6 Then
                    Printer.NewPage
                    Call AnalysisHead
                End If
                'Call Print_Setting(arytmp(ii), PrtLeft + Printer.TextWidth("소견사유 : "), LineSpace, 55, "L", "C")
                Call Print_Setting(arytmp(ii), 135, LineSpace, 150, "L", "C")
            Next
            Printer.FontBold = False
            Printer.DrawStyle = 1: Printer.DrawWidth = 2
            Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    
'    Dim strAccNo    As String
'    Dim strPtId     As String
'    Dim strPtNm     As String
'    Dim strSEXAGE   As String
'    Dim strLocation As String
'    Dim strMesg     As String
'    Dim strEntNm    As String
'    Dim strEntDT    As String
'    Dim strOrdDt    As String
'
'    Dim ii As Integer
'    If ssCmtList.DataRowCnt < 1 Then Exit Sub
'
'    Call P_PrtSet
'    Call AnalysisHead
'
'    With ssCmtList
'        For ii = 1 To .DataRowCnt
'            .Row = ii
'            .Col = 1:   strAccNo = .Value
'            .Col = 2:   strPtId = .Value
'            .Col = 3:   strPtNm = .Value
'            .Col = 4:   strSEXAGE = .Value
'            .Col = 5:   strMesg = .Value
'            Call PrintString(strAccNo, strPtId, strPtNm, strSEXAGE, strMesg)
'        Next
'    End With
'
'    Printer.EndDoc

    Dim strTitle     As String
    Dim strPrintDate As String
    Dim strTestNm    As String
    Dim strPDate     As String
    Dim tmpTitle     As String
    Dim strDate      As String
    Dim strGb        As String
    
    strGb = ""
    strPDate = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    With ssCmtList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .FontBold = False
        .FontSize = 9
        .BlockMode = False
               
        .PrintJobName = "CVR 통보 현황"

        .PrintAbortMsg = "CVR 통보 현황을 출력중입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
        
        tmpTitle = "CVR 통보 현황"
'        strTitle = "/fn""굴림체""/fz""18""/fb1/fi0/fu1/fk0/fs1" _
'              & "/f1/c" & tmpTitle & "/n/n/n"
        strTitle = "/fn""굴림체"" /fz""18"" /fb1/fi0/fu0/fk0/fs1" _
                  & "/f1/c" & tmpTitle & "/n/n/n"
        strPrintDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "출력일자 : " & strPDate & "/n/n"
        strTestNm = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "WorkArea : " & cboWorkArea.Text & "/n"
        strDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "조회기간 : " & Format(dtpStartDt.Value, "yyyy-mm-dd") & " ~ " & Format(dtpEndDt.Value, "yyyy-mm-dd") & "/n" '"   조회유형 : " & strGb & "/n"
        .PrintHeader = strTitle & strTestNm & strDate 'strPrintDate
        .PrintMarginLeft = 800
'        .PrintMarginRight = 10
'        .PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape
        .PrintOrientation = PrintOrientationLandscape 'PrintOrientationLandscape
        
        
        P_HOSPITALNAME = "예수병원 진단검사의학실"
        .PrintFooter = " /l " & String(130, Chr(6)) & "/n/l " & P_HOSPITALNAME & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintShadows = True
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint
    End With
End Sub

Private Sub ssCmtList_Click(ByVal Col As Long, ByVal Row As Long)
    Static iSortOrder As Integer
    
    With ssCmtList
        If Row = 0 Then  'Sort...
            .Row = 0: .Col = Col
            .Row = -1: .Col = -1
            .SortBy = SortByRow
            .SortKey(1) = Col
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
            End If
            .Action = ActionSort
            Exit Sub
        End If
    End With
    
End Sub

Private Sub ssCmtList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp
    Dim strFg As String
    Dim strFg1 As String
    
    strFg = ""
    
    If Col = 7 Then
        With ssCmtList
            .GetText 7, Row, varTmp: strFg = varTmp
            .GetText 5, Row, varTmp: rtfMessage.Text = varTmp
            .GetText 2, Row, varTmp: txtTransNo.Text = varTmp
            If strFg = "미확인" Then
                .GetText 9, Row, varTmp: txtDtNm.Text = varTmp
                .GetText 10, Row, varTmp: txtDtId.Text = varTmp
                .GetText 11, Row, varTmp: txtDtNo.Text = varTmp
            ElseIf strFg = "VRE" Then
'                .GetText 9, Row, varTmp: txtDtNm.Text = "손정아"
'                .GetText 10, Row, varTmp: txtDtId.Text = "003673"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-4478-1409"
'                .GetText 9, Row, varTmp: txtDtNm.Text = "문찬미"
'                .GetText 10, Row, varTmp: txtDtId.Text = "004382"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-3412-9964"
''박다야(4734 010 7434-6888)
'                .GetText 9, Row, varTmp: txtDtNm.Text = "박다야"
'                .GetText 10, Row, varTmp: txtDtId.Text = "004734"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-7434-6888"
'김선우(4734 010 7434-6888)
'                .GetText 9, Row, varTmp: txtDtNm.Text = "김선우"
'                .GetText 10, Row, varTmp: txtDtId.Text = "004728"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-5451-3354"
'                .GetText 9, Row, varTmp: txtDtNm.Text = "임송빈"
'                .GetText 10, Row, varTmp: txtDtId.Text = "003311"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-8985-1280"
                .GetText 9, Row, varTmp: txtDtNm.Text = "김현아"
                .GetText 10, Row, varTmp: txtDtId.Text = "004826"
                .GetText 11, Row, varTmp: txtDtNo.Text = "010-3152-6245"
            ElseIf strFg = "AFB" Or InStr(strFg, "CRE") > 0 Or UCase(strFg) = "STREP A" Or InStr(strFg, "CPE") > 0 Then
'                .GetText 9, Row, varTmp: txtDtNm.Text = "손정아"
'                .GetText 10, Row, varTmp: txtDtId.Text = "003673"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-4478-1409"
'                .GetText 9, Row, varTmp: txtDtNm.Text = "문찬미"
'                .GetText 10, Row, varTmp: txtDtId.Text = "004382"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-3412-9964"
'                .GetText 9, Row, varTmp: txtDtNm.Text = "박다야"
'                .GetText 10, Row, varTmp: txtDtId.Text = "004734"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-7434-6888"
'김선우(4734 010 7434-6888)
'                .GetText 9, Row, varTmp: txtDtNm.Text = "김선우"
'                .GetText 10, Row, varTmp: txtDtId.Text = "004728"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-5451-3354"
''임송빈 (3311)
''010 8985 1280
'                .GetText 9, Row, varTmp: txtDtNm.Text = "임송빈"
'                .GetText 10, Row, varTmp: txtDtId.Text = "003311"
'                .GetText 11, Row, varTmp: txtDtNo.Text = "010-8985-1280"
'김현아 (004826)
'010 3152 6245
                .GetText 9, Row, varTmp: txtDtNm.Text = "김현아"
                .GetText 10, Row, varTmp: txtDtId.Text = "004826"
                .GetText 11, Row, varTmp: txtDtNo.Text = "010-3152-6245"
                
            End If
            .GetText 12, Row, varTmp: txtDeptNm.Text = varTmp
            
            txtTransId.Text = ObjSysInfo.EmpId
            txtTransNm.Text = GetEmpNm(ObjSysInfo.EmpId)
            txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
            txtDtId.Text = txtDtNm.Text
            
            If strFg = "미확인" Or strFg = "VRE" Or strFg = "AFB" Or InStr(strFg, "CRE") > 0 Or InStr(strFg, "CPE") > 0 Or InStr(UCase(strFg), "STREP") > 0 Then
                frmSMS.Visible = True
            End If
        End With
    End If
    
End Sub
