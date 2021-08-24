VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEntryWS 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "ABO결과 일괄등록"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15030
   Icon            =   "frmEntryWS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1245
      Left            =   150
      TabIndex        =   34
      Top             =   75
      Width           =   10755
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회(&Q)"
         Height          =   945
         Left            =   9000
         Style           =   1  '그래픽
         TabIndex        =   20
         Tag             =   "135"
         Top             =   195
         Width           =   1320
      End
      Begin VB.OptionButton optCon 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수번호별"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   390
         Width           =   1275
      End
      Begin VB.OptionButton optCon 
         BackColor       =   &H00DBE6E6&
         Caption         =   "WorkSheet별"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   795
         Width           =   1515
      End
      Begin VB.Frame fraAcc 
         BackColor       =   &H00FFEBD7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   945
         Left            =   2220
         TabIndex        =   1
         Top             =   180
         Width           =   5865
         Begin VB.TextBox txtAccE 
            Height          =   345
            Left            =   4905
            TabIndex        =   6
            Top             =   300
            Width           =   600
         End
         Begin VB.TextBox txtAccS 
            Height          =   345
            Left            =   3825
            TabIndex        =   5
            Top             =   300
            Width           =   600
         End
         Begin MSComCtl2.DTPicker dtpAcc 
            Height          =   360
            Left            =   1095
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   300
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyy'-'MM'-'dd"
            Format          =   59965443
            CurrentDate     =   36287
         End
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   345
            Left            =   5520
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            BuddyControl    =   "txtAccE"
            BuddyDispid     =   196613
            OrigLeft        =   5505
            OrigTop         =   300
            OrigRight       =   5745
            OrigBottom      =   645
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MedControls1.LisLabel lbldt 
            Height          =   330
            Index           =   2
            Left            =   75
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
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
            Caption         =   "접수일자"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lbldt 
            Height          =   330
            Index           =   3
            Left            =   2820
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
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
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   345
            Left            =   4440
            TabIndex        =   36
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            BuddyControl    =   "txtAccS"
            BuddyDispid     =   196614
            OrigLeft        =   4425
            OrigTop         =   285
            OrigRight       =   4665
            OrigBottom      =   630
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Line Line2 
            X1              =   4785
            X2              =   4860
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.Frame fraWS 
         BackColor       =   &H00CDEBFE&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   945
         Left            =   2220
         TabIndex        =   8
         Top             =   180
         Width           =   5865
         Begin VB.TextBox txtWSE 
            Height          =   345
            Left            =   4905
            TabIndex        =   18
            Top             =   510
            Width           =   600
         End
         Begin VB.TextBox txtWSS 
            Height          =   345
            Left            =   3825
            TabIndex        =   16
            Top             =   510
            Width           =   600
         End
         Begin VB.TextBox txtWSCd 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1095
            TabIndex        =   10
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton cmdWSList 
            BackColor       =   &H00FFEBD7&
            Caption         =   "..."
            Height          =   330
            Left            =   2175
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   120
            Width           =   315
         End
         Begin MedControls1.LisLabel lblWSNm 
            Height          =   330
            Left            =   2520
            TabIndex        =   12
            Top             =   120
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   582
            BackColor       =   16252927
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            Appearance      =   0
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Left            =   4425
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            BuddyControl    =   "txtWSS"
            BuddyDispid     =   196618
            OrigLeft        =   4425
            OrigTop         =   510
            OrigRight       =   4665
            OrigBottom      =   855
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpWS 
            Height          =   360
            Left            =   1095
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   510
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyy'-'MM'-'dd"
            Format          =   59965443
            CurrentDate     =   36287
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   345
            Left            =   5505
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            BuddyControl    =   "txtWSE"
            BuddyDispid     =   196617
            OrigLeft        =   5505
            OrigTop         =   510
            OrigRight       =   5745
            OrigBottom      =   855
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MedControls1.LisLabel lbldt 
            Height          =   315
            Index           =   0
            Left            =   105
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
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
            Caption         =   "WS 코드"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lbldt 
            Height          =   330
            Index           =   5
            Left            =   105
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   510
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
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
            Caption         =   "작업일자"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lbldt 
            Height          =   330
            Index           =   1
            Left            =   2835
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   510
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
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
            Caption         =   "작업번호"
            Appearance      =   0
         End
         Begin VB.Line Line1 
            X1              =   4755
            X2              =   4830
            Y1              =   645
            Y2              =   645
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   13005
      Style           =   1  '그래픽
      TabIndex        =   33
      Tag             =   "128"
      Top             =   8325
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   11685
      Style           =   1  '그래픽
      TabIndex        =   32
      Tag             =   "124"
      Top             =   8325
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "실행(&S)"
      Height          =   510
      Left            =   10365
      Style           =   1  '그래픽
      TabIndex        =   31
      Tag             =   "15101"
      Top             =   8325
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1245
      Left            =   11055
      TabIndex        =   21
      Top             =   75
      Width           =   3270
      Begin VB.OptionButton optType 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Front Typing"
         Height          =   945
         Index           =   0
         Left            =   180
         Style           =   1  '그래픽
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   1305
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Back Typing"
         Height          =   945
         Index           =   1
         Left            =   1725
         Style           =   1  '그래픽
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   1305
      End
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Accession No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   135
      TabIndex        =   25
      Tag             =   "20003"
      Top             =   6165
      Width           =   14220
      Begin VB.ComboBox cboRemark 
         Height          =   300
         Left            =   7230
         Style           =   2  '드롭다운 목록
         TabIndex        =   29
         Top             =   1425
         Width           =   6915
      End
      Begin VB.CommandButton cmdCommentTemplete 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6735
         Picture         =   "frmEntryWS.frx":076A
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   1425
         Width           =   315
      End
      Begin VB.TextBox txtComment 
         Height          =   1455
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   26
         Top             =   315
         Width           =   6525
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7260
         TabIndex        =   28
         Top             =   1035
         Width           =   720
      End
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   4770
      Left            =   150
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "20001"
      Top             =   1365
      Width           =   14175
      _Version        =   196608
      _ExtentX        =   25003
      _ExtentY        =   8414
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   14737632
      GridShowVert    =   0   'False
      MaxCols         =   24
      MaxRows         =   15
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "frmEntryWS.frx":0C9C
      UserResize      =   0
      VisibleCols     =   5
      TextTip         =   2
   End
   Begin VB.Label lblErr 
      BackColor       =   &H00F7FFFF&
      Height          =   375
      Left            =   150
      TabIndex        =   30
      Top             =   8385
      Width           =   10020
   End
End
Attribute VB_Name = "frmEntryWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'입력된 접수번호의 결과등록여부 및 DC여부를 판단해야 한다.
'결과등록이 완료된 경우에는 아무런 표시도 하지 않고 다음으로 스킵
'따라서 조회할 수 있는 것들은 결과등록이 되지 않은 것들만을 대상으로 해야 한다.
'아무리 BBS303에 데이타가 없다고 하더라도..

Public Event FormClose()

Private Enum KindOfTyping
    NoUse = -1
    AnyUse = 0
    FrontUse = 1
    BackUse = 2
End Enum

Private Enum TblColumn
    tcWORKSEQ = 1
    tcPTNM
    tcSEL
    tcABO
    tcRH
    tcABOSUB
    tcRHSUB
    tcTYPE
    
    tcACCNO
    
    tcSEXAGE
    tcDEPT
    tcWARD
    
    tcLast
    tcABOCD
    tcRHCD
    tcABOSUBCD
    tcRHSUBCD
    
    tcRMKCD
    tcCOMMENT
    tcTYPECD
    
    tcWORKAREA
    tcACCDT
    tcACCSEQ
    tcToolRst
End Enum

Private MsgFg As Boolean

Private WithEvents objPopList As clsPopUpList
Attribute objPopList.VB_VarHelpID = -1

Private Sub cboRemark_Click()
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblColumn.tcRMKCD
'        .Value = medGetP(cboRemark.Text, 1, vbTab)
'    End With
End Sub

Private Sub cboRemark_LostFocus()
    With tblResult
        .Row = .ActiveRow
        .Col = TblColumn.tcRMKCD
        .Value = medGetP(cboRemark.Text, 1, vbTab)
    End With
End Sub

Private Sub cmdClear_Click()
    Call ClearAcc
    txtWSCd.Text = ""
    Call ClearWs
    Call ClearOthers
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub
'
Private Sub cmdQuery_Click()
    If optCon(0).Value Then
        Call AccQuery
    ElseIf optCon(1).Value Then
        Call WSQuery
    End If
End Sub

'Private Sub cmdSave_Click()
'    Call Save
'    If optCon(1).Value = True Then
'        Call Query
'    Else
'        Call AccQuery
'    End If
'End Sub

'Private Sub cmdWSList_Click()
'    If lstWSCode.ListCount = 0 Then
'        MsgBox "등록된 Worksheet 코드가 없습니다.", vbExclamation, "메세지"
'        Exit Sub
'    End If
'    lstWSCode.Visible = True
'    lstWSCode.ZOrder 0
'    Call CodeHelp(0, lstWSCode, txtWorkCd.Text, txtWorkCd, dptWorkDt)
'End Sub

'Private Sub dptWorkDt_Change()
'    Call SetMinMaxSeq
'End Sub

'Private Sub dtpAcc_Change()
'    Call GetAccNo
'End Sub

'Private Sub Form_Activate()
''    medMain.lblSubMenu.Caption = Me.Caption
'    If Not IsFirst Then Exit Sub
'    IsFirst = True
'
'    dptWorkDt = GetSystemDate
'    Call ClearAll
'
'    Me.Show
'    DoEvents
'    Call GetAccNo
''    Call SetRemarkCombo
''    Call SetResultTemplate
''    Call SetWorkSheetList
'End Sub

'Private Sub Form_Load()
'    IsFirst = True
'    Call LoadTestCd
'End Sub

'Private Sub LoadTestCd()
'    Dim objcom003 As clsCom003
'    Dim Rs As Recordset
'
'    Set objcom003 = New clsCom003
'    Set Rs = objcom003.OpenRecordSetDay(BC2_ABO_TEST)
'    Set objcom003 = Nothing
'
'    If Rs.EOF = False Then
'        CODE_ABOFRONT = Rs.Fields("field1").Value & ""
'        CODE_ABOBACK = Rs.Fields("field2").Value & ""
'        CODE_RH = Rs.Fields("field3").Value & ""
'        CODE_ABOSUB = Rs.Fields("field4").Value & ""
'        CODE_RHSUB = Rs.Fields("text1").Value & ""
'    End If
'
'    Set Rs = Nothing
'    Set objcom003 = Nothing
'End Sub

'Private Sub lstRstABO_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim result As String
'    Dim cd As String
'
'    With lstRstABO
'        If .ListIndex = 0 Then
'            cd = ""
'            result = ""
'        Else
'            cd = medGetP(.Text, 1, vbTab)
'            result = medGetP(.Text, 2, vbTab)
'        End If
'    End With
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblColumn.tcABO:   .Value = result
'        .Col = TblColumn.tcABOCD: .Value = cd
'    End With
'    blnClick = False
'    lstRstABO.Visible = False
'    tblResult.SetFocus
'End Sub

'Private Sub lstRstABOSUB_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim result As String
'    Dim cd As String
'
'    With lstRstABOSUB
'        If .ListIndex = 0 Then
'            cd = ""
'            result = ""
'        Else
'            cd = medGetP(.Text, 1, vbTab)
'            result = medGetP(.Text, 2, vbTab)
'        End If
'    End With
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblColumn.tcABOSUB:   .Value = result
'        .Col = TblColumn.tcABOSUBCD: .Value = cd
'    End With
'    blnClick = False
'    tblResult.SetFocus
'    lstRstABOSUB.Visible = False
'End Sub

'Private Sub lstRstRH_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim result As String
'    Dim cd As String
'
'    With lstRstRH
'        If .ListIndex = 0 Then
'            cd = ""
'            result = ""
'        Else
'            cd = medGetP(.Text, 1, vbTab)
'            result = medGetP(.Text, 2, vbTab)
'        End If
'    End With
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblColumn.tcRH:   .Value = result
'        .Col = TblColumn.tcRHCD: .Value = cd
'    End With
'    blnClick = False
'    tblResult.SetFocus
'    lstRstRH.Visible = False
'End Sub

'Private Sub lstRstRHSUB_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim result As String
'    Dim cd As String
'
'    With lstRstRHSUB
'        If .ListIndex = 0 Then
'            cd = ""
'            result = ""
'        Else
'            cd = medGetP(.Text, 1, vbTab)
'            result = medGetP(.Text, 2, vbTab)
'        End If
'    End With
'
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblColumn.tcRHSUB:   .Value = result
'        .Col = TblColumn.tcRHSUBCD: .Value = cd
'    End With
'    blnClick = False
'    tblResult.SetFocus
'    lstRstRHSUB.Visible = False
'End Sub

'Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyReturn And lstWSCode.ListIndex >= 0 Then
'        txtWorkCd.Text = Trim(Mid(lstWSCode.Text, 1, _
'                 InStr(1, lstWSCode.Text, Chr(vbKeyTab)) - 1))
'        lblWorkCdNm.Caption = medGetP(lstWSCode.Text, 2, Chr(vbKeyTab))
'        lstWSCode.Visible = False
'        dptWorkDt.SetFocus
'
'        Call SetMinMaxSeq
'    End If
'End Sub

Private Sub cmdSave_Click()
    Dim i As Long

    Dim objABOSql As clsABOSql

    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String

    Dim strVfydt As String
    Dim strVfytm As String

    Dim strABO As String
    Dim strRh As String
    Dim strABOSub As String
    Dim strRhSub As String
    Dim strComment As String
    Dim strRemark As String

    Dim blnVerify As Boolean
    Dim strTyping As String

    Dim objPrgBar As clsProgress

    strVfydt = Format(GetSystemDate, "YYYYMMDD")
    strVfytm = Format(GetSystemDate, "HHMMSS")

    Set objABOSql = New clsABOSql

    With tblResult
        Set objPrgBar = New clsProgress
        objPrgBar.Container = MainFrm.stsbar
        objPrgBar.max = .DataRowCnt

        For i = 1 To .DataRowCnt
            objPrgBar.Value = i

            .Row = i
            .Col = TblColumn.tcWORKAREA:    strWorkarea = .Value
            .Col = TblColumn.tcACCDT:       strAccdt = .Value
            .Col = TblColumn.tcACCSEQ:      strAccseq = .Value
            .Col = TblColumn.tcSEL:         blnVerify = IIf(.Value = 1, False, True)
            .Col = TblColumn.tcABOCD:       strABO = .Value
            .Col = TblColumn.tcRHCD:        strRh = .Value
            .Col = TblColumn.tcABOSUBCD:    strABOSub = .Value
            .Col = TblColumn.tcRHSUBCD:     strRhSub = .Value
            .Col = TblColumn.tcRMKCD:       strRemark = .Value
            .Col = TblColumn.tcCOMMENT:     strComment = .Value
            .Col = TblColumn.tcTYPE:        strTyping = IIf(.Value = "", -1, .Value)

            Call SaveByAccNo(objABOSql, _
                             strWorkarea, strAccdt, strAccseq, strVfydt, strVfytm, _
                             blnVerify, strTyping, _
                             strABO, strRh, strABOSub, strRhSub, _
                             strRemark, strComment)
        Next i

        Set objPrgBar = Nothing
    End With

    Set objABOSql = Nothing
End Sub

Private Function SaveByAccNo(ByVal objABOSql As clsABOSql, _
                             ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                             ByVal vvfydt As String, ByVal vVfytm As String, _
                             ByVal vVerify As Boolean, ByVal vTyping As Long, _
                             ByVal vABO As String, ByVal vRh As String, ByVal vABOSub As String, ByVal vRhSub As String, _
                             ByVal vRemark As String, ByVal vComment As String) As Boolean
    '------------------------------------------------------
    '1. bbs303에 저장시킨다.
    '2. 만일, double check완료시점이면 lab302에도 저장한다.
    '
    '* lab302에 저장시 결과유형(rsttype)은 F로 변경시킨다.
    '------------------------------------------------------

    SaveByAccNo = False

    '입력된 접수번호 검사----------------------------------------------
    If vWorkarea = "" Or vAccdt = "" Or vAccseq = "" Then Exit Function
    If vTyping = -1 Then Exit Function

    '결과값 검사-------------------------------------------------------
    If vVerify Then
        If vABO = "" And vABOSub = "" Then Exit Function
        If vRh = "" And vRhSub = "" Then Exit Function
        If vABO <> "" And vABOSub <> "" Then Exit Function
        If vRh <> "" And vRhSub <> "" Then Exit Function
    End If
    
    SaveByAccNo = objABOSql.SaveABOResult(vWorkarea, vAccdt, vAccseq, vVerify, vTyping, vABO, vRh, vABOSub, vRhSub, vRemark, vComment) ', ObjMyUser.EmpID, vfydt, vfytm)
End Function

Private Sub cmdWSList_Click()
    Dim objABOSql As clsABOSql
    Dim Rs  As Recordset
    

    Set objABOSql = New clsABOSql
    Set Rs = New Recordset
    Set Rs = objABOSql.LoadWorksheet(ObjSysInfo.buildingcd)
    Set objABOSql = Nothing

    If Rs.EOF Then
        MsgBox "워크시트 마스터가 없습니다.", vbExclamation
        Set Rs = Nothing
    End If
    
    Set objPopList = New clsPopUpList
    
    With objPopList
        .Tag = "WS"
        
        .Recordset = Rs
'        .FormWidth = 4635
'        .FormHeight = 2880
        .ColumnHeaderWidth = "1019.906;2145.26"
        .FormCaption = "워크시트 찾기"
        .ColumnHeaderText = "WS코드;워크시트 명"
        
        .LoadPopUp
    End With
        
    Set Rs = Nothing
    Set objPopList = Nothing
End Sub

Private Sub dtpAcc_Change()
    Call GetAccDuration
    If tblResult.DataRowCnt <> 0 Then Call ClearOthers
End Sub

Private Sub dtpWS_Change()
    If txtWSCd.Text <> "" Then Call GetWsDuration
    If tblResult.DataRowCnt <> 0 Then Call ClearOthers
End Sub

Private Sub Form_Load()
'    optCon(0).Value = True
        
    Call ClearAcc
    txtWSCd.Text = ""
    Call ClearWs
    Call LoadRemark(cboRemark)
        
    Call ClearOthers
End Sub

Private Sub ClearAcc()
    dtpAcc.Value = GetSystemDate
    txtAccS.Text = ""
    txtAccE.Text = ""
End Sub

Private Sub ClearWs()
    lblWSNm.Caption = ""
    dtpWS.Value = GetSystemDate
    txtWSS.Text = ""
    txtWSE.Text = ""
End Sub

Private Sub ClearOthers()
    Call medClearTable(tblResult)
    tblResult.MaxRows = 15
    tblResult.RowHeight(-1) = 12
    txtComment.Text = ""
    cboRemark.ListIndex = -1
    lblErr.Caption = ""
    optType(0).Value = False
    optType(1).Value = False
End Sub

Private Sub objPopList_SelectedItem(ByVal pSelectedItem As String)
'결과입력
    Select Case objPopList.Tag
        Case "4"
            With tblResult
                .Row = .ActiveRow
                .Col = TblColumn.tcABO:   .Value = objPopList.SelectedItems(1)
                .Col = TblColumn.tcABOCD: .Value = objPopList.SelectedItems(0)
            End With
        Case "5"
            With tblResult
                .Row = .ActiveRow
                .Col = TblColumn.tcRH:   .Value = objPopList.SelectedItems(1)
                .Col = TblColumn.tcRHCD: .Value = objPopList.SelectedItems(0)
            End With
        Case "6"
            With tblResult
                .Row = .ActiveRow
                .Col = TblColumn.tcABOSUB:   .Value = objPopList.SelectedItems(1)
                .Col = TblColumn.tcABOSUBCD: .Value = objPopList.SelectedItems(0)
            End With
        Case "7"
            With tblResult
                .Row = .ActiveRow
                .Col = TblColumn.tcRHSUB:   .Value = objPopList.SelectedItems(1)
                .Col = TblColumn.tcRHSUBCD: .Value = objPopList.SelectedItems(0)
            End With
        Case "WS"
            txtWSCd.Text = objPopList.SelectedItems(0)
            lblWSNm.Caption = objPopList.SelectedItems(1)
            
            If txtWSCd.Text <> "" Then Call GetWsDuration
    End Select
End Sub

Private Sub optCon_Click(Index As Integer)
    Select Case Index
        Case 0
            fraAcc.Enabled = True
            fraAcc.Visible = True
            fraWS.Visible = False
            fraWS.Enabled = False
            
            Call GetAccDuration
        Case 1
            fraWS.Visible = True
            fraWS.Enabled = True
            fraAcc.Enabled = False
            fraAcc.Visible = False
    End Select
    
    Call ClearOthers
End Sub

Private Sub GetAccDuration()
' 접수번호별 번호 맥스,미니멈 구하기

    Dim strTmp As String
    Dim objSQL As New clsABOSql
    
    strTmp = objSQL.GetAccNoDuration(Format(dtpAcc.Value, "yyyymmdd"))
    If strTmp <> "" Then
        txtAccS.Text = medGetP(strTmp, 2, COL_DIV)
        txtAccE.Text = medGetP(strTmp, 1, COL_DIV)
        cmdQuery.SetFocus
    Else
        txtAccS.Text = ""
        txtAccE.Text = ""
        dtpAcc.SetFocus
    End If
    Set objSQL = Nothing
    
End Sub

Private Sub GetWsDuration()
    Dim min As Long
    Dim max As Long
    Dim objABOSql As clsABOSql

    Set objABOSql = New clsABOSql
    Call objABOSql.GetWsNoDuration(Format(dtpWS.Value, "YYYYMMDD"), txtWSCd.Text, min, max)
    Set objABOSql = Nothing
    
    txtWSS.Text = IIf(min = 0, "", min)
    txtWSE.Text = IIf(max = 0, "", max)
End Sub

Private Sub optType_Click(Index As Integer)
    Dim i As Long
    
    With tblResult
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn.tcTYPECD
            If .Value = "0" Then
                .Col = TblColumn.tcTYPE
                .Value = Index
            End If
            
            .Col = TblColumn.tcABO: .Value = ""
            .Col = TblColumn.tcRH: .Value = ""
            .Col = TblColumn.tcABOSUB: .Value = ""
            .Col = TblColumn.tcRHSUB: .Value = ""
            
            .Col = TblColumn.tcABOCD: .Value = ""
            .Col = TblColumn.tcRHCD: .Value = ""
            .Col = TblColumn.tcABOSUBCD: .Value = ""
            .Col = TblColumn.tcRHSUBCD: .Value = ""
        Next i
    End With
End Sub

Private Sub tblResult_Advance(ByVal AdvanceNext As Boolean)
    Dim lngRow As Long

    lngRow = IIf(AdvanceNext, tblResult.DataRowCnt, 1)
    Call tblResult_LeaveCell(tblResult.ActiveCol, lngRow, tblResult.ActiveCol, lngRow, False)
End Sub

Private Sub tblResult_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
'Type이 바뀌면 현재 결과 지우기..

    If Col <> TblColumn.tcTYPE Then Exit Sub
    
    With tblResult
        .Col = TblColumn.tcABO: .Value = ""
        .Col = TblColumn.tcRH: .Value = ""
        .Col = TblColumn.tcABOSUB: .Value = ""
        .Col = TblColumn.tcRHSUB: .Value = ""
        
        .Col = TblColumn.tcABOCD: .Value = ""
        .Col = TblColumn.tcRHCD: .Value = ""
        .Col = TblColumn.tcABOSUBCD: .Value = ""
        .Col = TblColumn.tcRHSUBCD: .Value = ""
    End With
End Sub

Private Sub tblResult_EditChange(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    If Col > 3 And Col < 8 Then
        tblResult.Row = Row
        If Col = TblColumn.tcABO Then Col = TblColumn.tcABOCD
        If Col = TblColumn.tcRH Then Col = TblColumn.tcRHCD
        If Col = TblColumn.tcABOSUB Then Col = TblColumn.tcABOSUBCD
        If Col = TblColumn.tcRHSUB Then Col = TblColumn.tcRHSUBCD
        
        tblResult.Col = Col
        tblResult.Value = ""
    End If
End Sub

Private Sub tblResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim vValue As Variant '코드값
    Dim pValue As String '리턴받는 코드명 값
    Dim vCode As Variant '입력되어 있는 코드 값
    
    If tblResult.DataRowCnt = 0 Then Exit Sub
    If Row = 0 Then Exit Sub
    tblResult.Col = TblColumn.tcPTNM
    tblResult.Row = Row
    If tblResult.Value = "" Then Exit Sub
    
    
    Call ShowCommentandRemark(NewRow)        'Comment 및 리마크 표시

    If Col > 3 And Col < 7 Then
        If MsgFg Then Exit Sub
        
        Call tblResult.GetText(Col, Row, vValue)
        
        If vValue = "" Then Exit Sub
        
        If Col = TblColumn.tcABO Then Col = TblColumn.tcABOCD
        If Col = TblColumn.tcRH Then Col = TblColumn.tcRHCD
        If Col = TblColumn.tcABOSUB Then Col = TblColumn.tcABOSUBCD
        If Col = TblColumn.tcRHSUB Then Col = TblColumn.tcRHSUBCD
        Call tblResult.GetText(Col, Row, vCode)
        If vCode <> "" Then Exit Sub
        
        If Col = TblColumn.tcABOCD Then Col = TblColumn.tcABO
        If Col = TblColumn.tcRHCD Then Col = TblColumn.tcRH
        If Col = TblColumn.tcABOSUBCD Then Col = TblColumn.tcABOSUB
        If Col = TblColumn.tcRHSUBCD Then Col = TblColumn.tcRHSUB
        If CheckResult(Col, vValue, pValue) = False Then
            Cancel = True
            Exit Sub
        End If
        
        Call tblResult.SetText(Col, Row, pValue) '코드명으로 표시
        
        If Col = TblColumn.tcABO Then Col = TblColumn.tcABOCD
        If Col = TblColumn.tcRH Then Col = TblColumn.tcRHCD
        If Col = TblColumn.tcABOSUB Then Col = TblColumn.tcABOSUBCD
        If Col = TblColumn.tcRHSUB Then Col = TblColumn.tcRHSUBCD
    
        Call tblResult.SetText(Col, Row, vValue) '코드 저장
        tblResult.EditEnterAction = EditEnterActionDown
    ElseIf Col = 7 Then
        tblResult.Col = TblColumn.tcABO
        If tblResult.Row = tblResult.DataRowCnt Then
            tblResult.Row = 1
        Else
            tblResult.Row = Row + 1
        End If
        tblResult.Action = ActionActiveCell
    End If
End Sub

Private Sub ShowCommentandRemark(ByVal vRow As Long)
    If vRow < 1 Or vRow > tblResult.DataRowCnt Then Exit Sub
    
    With tblResult
        .Row = vRow
        .Col = TblColumn.tcCOMMENT: txtComment.Text = .Value
        .Col = TblColumn.tcRMKCD
        If .Value = "" Then
            cboRemark.ListIndex = -1
        Else
            cboRemark.ListIndex = medComboFind(cboRemark, .Value)
        End If
    End With
End Sub

Private Function CheckResult(ByVal vCol As Long, ByVal vValue As Variant, ByRef pValue As String) As Boolean
'입력된 데이타로 코드를 읽어 코드가 있는지 여부 체크
'현재 편집되는 컬럼의 코드로
'찾은 코드가 있으면 코드명으로 변환해주고, 코드는 코드값 컬럼에 저장해놓는다.
    Dim strCdNm As String
    Dim strType As String
    
    tblResult.Col = TblColumn.tcTYPE
    tblResult.Row = tblResult.ActiveRow
    strType = tblResult.Value
    
    If Val(strType) < 0 Or strType = "" Then
        Call tblResult.SetText(vCol, tblResult.ActiveRow, "")
        
        MsgFg = True
        MsgBox "결과입력 Type을 선택하십시오.", vbExclamation
        MsgFg = False
        CheckResult = False
        
        tblResult.SetFocus
        tblResult.Col = TblColumn.tcTYPE
        tblResult.Row = tblResult.ActiveRow
        tblResult.Action = ActionActiveCell
        
        Exit Function
    End If
    
    Select Case vCol
        Case TblColumn.tcABO
            If strType = "0" Then
                strCdNm = GetABOFrontNm(vValue)
            ElseIf strType = "1" Then
                strCdNm = GetABOBackNm(vValue)
            End If
        Case TblColumn.tcRH
            strCdNm = GetRHNM(vValue)
        Case TblColumn.tcABOSUB
            strCdNm = GetABOSUBNM(vValue)
        Case TblColumn.tcRHSUB
            strCdNm = GetRHSUBNM(vValue)
    End Select
    
    If strCdNm = "" Then
        CheckResult = False
        pValue = ""
        
        MsgFg = True
        MsgBox "결과코드 입력 오류입니다.", vbExclamation
        MsgFg = False
        
        tblResult.SetFocus
        tblResult.Col = vCol
        tblResult.Row = tblResult.ActiveRow
        tblResult.Action = ActionActiveCell
    Else
        CheckResult = True
        pValue = strCdNm
    End If
End Function

Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim Rs As Recordset
    
    If Row < 1 Then Exit Sub
    If Col < 4 Or Col > 7 Then Exit Sub
    
    tblResult.Col = TblColumn.tcPTNM
    tblResult.Row = Row
    If tblResult.Value = "" Then Exit Sub
    
    Set objPopList = New clsPopUpList
    Set Rs = New Recordset

    Set Rs = GetABOResult(Col)
    
    If Rs Is Nothing Then
        Set Rs = Nothing
        Set objPopList = Nothing
        Exit Sub
    End If
    
    If Rs.EOF = False Then
        With objPopList
            .Tag = Col

            .Recordset = Rs
            .HideSearchTool = True
            .SelectByClick = True
            .FormWidth = 4635
            .FormHeight = 2880
            .FormCaption = "결과코드 찾기"
            .ColumnHeaderText = "결과코드;결과코드명"
            .ColumnHeaderWidth = "1110.047;3075.024"
            .HideColumnHeaders = True
            .LoadPopUp
        End With
    End If
    Set objPopList = Nothing
End Sub

Private Function GetABOResult(ByVal vCol As String) As Recordset
    Dim strSQL As String
    Dim strType As String
        
    tblResult.Col = TblColumn.tcTYPE
    tblResult.Row = tblResult.ActiveRow
    strType = tblResult.Value
    
    If Val(strType) < 0 Or strType = "" Then
        MsgFg = True
        MsgBox "결과입력 Type을 선택하십시오.", vbExclamation
        MsgFg = False
        tblResult.SetFocus
        tblResult.Action = ActionActiveCell
        Exit Function
    End If
    
    'abo, rh, sub, rh
    Select Case vCol
        Case 4 'ABO
            If strType = "0" Then
                strSQL = GetSqlABOFrontList
            ElseIf strType = "1" Then
                strSQL = GetSqlABOBackList
            End If
        Case 5 ' Rh
            strSQL = GetSqlRhList
        Case 6 'ABO Sub
            strSQL = GetSqlABOSubList
        Case 7 'Rh Sub
            strSQL = GetSqlRhSubList
    End Select

    If strSQL <> "" Then
        Set GetABOResult = New Recordset

        GetABOResult.Open strSQL, DBConn
    End If
End Function

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row < 1 Then Exit Sub
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    tblResult.Row = Row
    tblResult.Col = TblColumn.tcLast
    If tblResult.Value = "" Then Exit Sub
    
    tblResult.Col = TblColumn.tcToolRst
    
    MultiLine = 1
    TipText = vbNewLine & "  - 최근 결과 -" & vbNewLine & vbNewLine & Replace(tblResult.Value, COL_DIV, vbNewLine)
    TipWidth = 5000
    Call tblResult.SetTextTipAppearance("굴림체", 9, True, False, &HEEFDF2, &H996666)
    ShowTip = True
End Sub

Private Sub txtAccE_Change()
    If tblResult.DataRowCnt <> 0 Then Call ClearOthers
End Sub

Private Sub txtAccS_Change()
    If tblResult.DataRowCnt <> 0 Then Call ClearOthers
End Sub

'Private Sub tblResult_Advance(ByVal AdvanceNext As Boolean)
'    'LeaveCell과 같은 역할을 한다.
'    'MsgBox tblResult.ActiveCol & "," & tblResult.ActiveRow
'    If onPgm Then Exit Sub
'
'    If TblResultLeaveCell(tblResult.ActiveCol, tblResult.ActiveRow) = True Then
'        SendKeys "{TAB}"
'    Else
'        tblResult.EditMode = True
'    End If
'End Sub

'Private Sub tblResult_EditChange(ByVal Col As Long, ByVal Row As Long)
'    tblResult.Row = Row
'    If Col = TblColumn.tcABO Then Col = TblColumn.tcABOCD
'    If Col = TblColumn.tcRH Then Col = TblColumn.tcRHCD
'    If Col = TblColumn.tcABOSUB Then Col = TblColumn.tcABOSUBCD
'    If Col = TblColumn.tcRHSUB Then Col = TblColumn.tcRHSUBCD
'    tblResult.Col = Col
'    If Col <> TblColumn.tcTYPE Then
'        tblResult.Value = ""
'    End If
'End Sub

'Private Sub tblResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim i As Long
'
'    '입력값이 정확한지 검사한다.
'    '정확하지 않으면 Clear하고 잡아둔다.
'    If onPgm Then Exit Sub
'
'
'
'
'    Dim Col1 As Integer
'    Dim strValue As String
'
'    Col1 = Col
'
'    tblResult.Row = Row
'    tblResult.Col = Col1
'
'    strValue = tblResult.Value
'
'
'
'    If Col = TblColumn.tcABO Then Col = TblColumn.tcABOCD
'    If Col = TblColumn.tcRH Then Col = TblColumn.tcRHCD
'    If Col = TblColumn.tcABOSUB Then Col = TblColumn.tcABOSUBCD
'    If Col = TblColumn.tcRHSUB Then Col = TblColumn.tcRHSUBCD
'
'    If blnClick = True Then
'        tblResult.Row = Row
'        tblResult.Col = Col
'        If tblResult.Value = "" Then
'
'            tblResult.Value = strValue
'        End If
'    End If
'    blnClick = True
'
'    If TblResultLeaveCell(Col1, Row) = False Then
'        onPgm = True
'        tblResult.Row = Row
'        tblResult.Col = Col
'        tblResult.Action = ActionActiveCell
'        onPgm = False
'        Exit Sub
'    End If
'
'    If Row = NewRow Then Exit Sub
'    If NewRow < 0 Then Exit Sub
'
'    onPgm = True
'
'    With tblResult
'        .Row = NewRow
'        .Col = TblColumn.tcCOMMENT: txtComment = .Value
'        .Col = TblColumn.tcRMKCD
'        cboRemark.ListIndex = -1
'        For i = 0 To cboRemark.ListCount - 1
'            If .Value = medGetP(cboRemark.List(i), 1, " ") Then
'                cboRemark.ListIndex = i
'                Exit For
'            End If
'        Next i
'    End With
'
'    onPgm = False
'End Sub

'Private Function TblResultLeaveCell(ByVal Col As Long, ByVal Row As Long) As Boolean
'    Dim cd As String
'    Dim result As String
'
'    Dim Col1 As Integer
'
'    Col1 = Col
'
'    If Col = TblColumn.tcABO Then Col = TblColumn.tcABOCD
'    If Col = TblColumn.tcRH Then Col = TblColumn.tcRHCD
'    If Col = TblColumn.tcABOSUB Then Col = TblColumn.tcABOSUBCD
'    If Col = TblColumn.tcRHSUB Then Col = TblColumn.tcRHSUBCD
'
'    With tblResult
'        .Row = Row
'
'        .Col = Col: cd = UCase(.Value)  '코드로 입력한다.
'
'        Select Case Col1
'            Case TblColumn.tcABO
'                If cd = "" Then
'                    .Col = TblColumn.tcABO:   .Value = ""
'                    .Col = TblColumn.tcABOCD: .Value = ""
'
'                    TblResultLeaveCell = True
'                Else
'                    result = GetABONM(cd)
'                    If result = "" Then
'                        onPgm = True
'                        MsgBox "존재하지 않는 코드입니다.", vbCritical, Me.Caption
'                        onPgm = False
'                        .Col = TblColumn.tcABO:   .Value = ""
'                        .Col = TblColumn.tcABOCD: .Value = ""
'
'                        TblResultLeaveCell = False
'                    Else
'                        .Col = TblColumn.tcABO:   .Value = result
'                        .Col = TblColumn.tcABOCD: .Value = cd
'
'                        TblResultLeaveCell = True
'                    End If
'                End If
'
'            Case TblColumn.tcABOSUB
'                If cd = "" Then
'                    .Col = TblColumn.tcABOSUB:   .Value = ""
'                    .Col = TblColumn.tcABOSUBCD: .Value = ""
'
'                     TblResultLeaveCell = True
'                Else
'                    result = GetABOSUBNM(cd)
'                    If result = "" Then
'                        onPgm = True
'                        MsgBox "존재하지 않는 코드입니다.", vbCritical, Me.Caption
'                        onPgm = False
'                        .Col = TblColumn.tcABOSUB:   .Value = ""
'                        .Col = TblColumn.tcABOSUBCD: .Value = ""
'
'                        TblResultLeaveCell = False
'                    Else
'                        .Col = TblColumn.tcABOSUB:   .Value = result
'                        .Col = TblColumn.tcABOSUBCD: .Value = cd
'
'                        TblResultLeaveCell = True
'                    End If
'                End If
'
'            Case TblColumn.tcRH
'                If cd = "" Then
'                    .Col = TblColumn.tcRH:   .Value = ""
'                    .Col = TblColumn.tcRHCD: .Value = ""
'
'                    TblResultLeaveCell = True
'                Else
'                    result = GetRHNM(cd)
'                    If result = "" Then
'                        onPgm = True
'                        MsgBox "존재하지 않는 코드입니다.", vbCritical, Me.Caption
'                        onPgm = False
'                        .Col = TblColumn.tcRH:   .Value = ""
'                        .Col = TblColumn.tcRHCD: .Value = ""
'
'                        TblResultLeaveCell = False
'                    Else
'                        .Col = TblColumn.tcRH:   .Value = result
'                        .Col = TblColumn.tcRHCD: .Value = cd
'
'                        TblResultLeaveCell = True
'                    End If
'                End If
'
'            Case TblColumn.tcRHSUB
'                If cd = "" Then
'                    .Col = TblColumn.tcRHSUB:   .Value = ""
'                    .Col = TblColumn.tcRHSUBCD: .Value = ""
'
'                    TblResultLeaveCell = True
'                Else
'                    result = GetRHSUBNM(cd)
'                    If result = "" Then
'                        onPgm = True
'                        MsgBox "존재하지 않는 코드입니다.", vbCritical, Me.Caption
'                        onPgm = False
'                        .Col = TblColumn.tcRHSUB:   .Value = ""
'                        .Col = TblColumn.tcRHSUBCD: .Value = ""
'
'                        TblResultLeaveCell = False
'                    Else
'                        .Col = TblColumn.tcRHSUB:   .Value = result
'                        .Col = TblColumn.tcRHSUBCD: .Value = cd
'
'                        TblResultLeaveCell = True
'                    End If
'                End If
'            Case Else
'                TblResultLeaveCell = True
'        End Select
'    End With
'End Function

'Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'    Dim X As Long
'    Dim Y As Long
'    Dim w As Long
'    Dim h As Long
'
''    If Row < 1 Then Exit Sub
''    If Row > tblResult.MaxRows Then Exit Sub
'
''    If lstRstABO.Visible Then Exit Sub
''    If lstRstRH.Visible Then Exit Sub
''    If lstRstABOSUB.Visible Then Exit Sub
''    If lstRstRHSUB.Visible Then Exit Sub
'
''    With tblResult
''        .Row = Row
''        .Col = Col
''        .Action = ActionActiveCell
''        Call .GetCellPos(Col, Row, X, Y, w, h)
''    End With
''
''    Select Case Col
''        Case TblColumn.tcABO
''            lstRstABO.Top = Y + h + picRst.Top
''            lstRstABO.Left = X + picRst.Left
''            lstRstABO.Visible = True
''            lstRstABO.SetFocus
''        Case TblColumn.tcABOSUB
''            lstRstABOSUB.Top = Y + h + picRst.Top
''            lstRstABOSUB.Left = X + picRst.Left
''            lstRstABOSUB.Visible = True
''            lstRstABOSUB.SetFocus
''        Case TblColumn.tcRH
''            lstRstRH.Top = Y + h + picRst.Top
''            lstRstRH.Left = X + picRst.Left
''            lstRstRH.Visible = True
''            lstRstRH.SetFocus
''        Case TblColumn.tcRHSUB
''            lstRstRHSUB.Top = Y + h + picRst.Top
''            lstRstRHSUB.Left = X + picRst.Left
''            lstRstRHSUB.Visible = True
''            lstRstRHSUB.SetFocus
''    End Select
''    blnClick = True
'
'    Dim Rs As Recordset
'
'    If optType(0).Value = False And optType(1).Value = False Then
'        Exit Sub
'    End If
'
'    Set objPopList = New clsPopUpList
'    Set Rs = New Recordset
'
'    Set Rs = GetABOResult(Col)
'
'    If Rs.EOF = False Then
'        With objPopList
'            .Tag = Col
'
'            .Recordset = Rs
'            .HideSearchTool = True
'            .SelectByClick = True
'            .FormWidth = 4635
'            .FormHeight = 2880
'            .FormCaption = "결과코드 찾기"
'            .ColumnHeaderText = "결과코드;결과코드명"
'            .ColumnHeaderWidth = "1110.047;3075.024"
'            .HideColumnHeaders = True
'            .LoadPopUp
'        End With
'    End If
'    Set objPopList = Nothing
'
'End Sub

'Private Function GetABOResult(ByVal vCol As String) As Recordset
'    Dim strSQL As String
'
'    'abo, rh, sub, rh
'    Select Case vCol
'        Case 4 'ABO
'            strSQL = GetSqlABOList
'        Case 5 ' Rh
'            strSQL = GetSqlRhList
'        Case 6 'ABO Sub
'            strSQL = GetSqlABOSubList
'        Case 7 'Rh Sub
'            strSQL = GetSqlRhSubList
'    End Select
'
'    If strSQL <> "" Then
'        Set GetABOResult = New Recordset
'
'        GetABOResult.Open strSQL, DBConn
'    End If
'End Function

Private Sub txtComment_Change()
'    With tblResult
'        .Row = .ActiveRow
'        .Col = TblColumn.tcCOMMENT
'        .Value = txtComment.Text
'    End With
End Sub

'Private Sub txtWorkCd_Change()
'
'   If txtWorkCd.Text = "" Then lblWorkCdNm.Caption = ""
'
'End Sub
'
'Private Sub txtWorkCd_GotFocus()
'   '
'   FocusMe Me.txtWorkCd
'   '
'End Sub

'Private Sub txtWorkCd_KeyPress(KeyAscii As Integer)
'Dim Char As String
'   Char = Chr(KeyAscii)
'   KeyAscii = Asc(UCase(Char))
'   If KeyAscii = vbKeyEscape Then Exit Sub
'   If KeyAscii = vbKeyReturn Then
'        Call lstWSCode_KeyDown(vbKeyReturn, 0)
'        lstWSCode.Visible = False
'        Exit Sub
'   End If
'
'   lstWSCode.Visible = True
'   lstWSCode.ZOrder 0
'   Call CodeHelp(KeyAscii, lstWSCode, txtWorkCd.Text, txtWorkCd, dptWorkDt)
'
'End Sub

'Private Sub txtWorkCd_KeyDown(KeyCode As Integer, Shift As Integer)
'    If lstWSCode.ListCount = 0 Then Exit Sub
'    If KeyCode = vbKeyDown Then
'        lstWSCode.Visible = True
'        If lstWSCode.ListIndex < lstWSCode.ListCount - 1 Then lstWSCode.ListIndex = lstWSCode.ListIndex + 1
'        lstWSCode.ZOrder 0
'        lstWSCode.SetFocus
'    End If
'End Sub

'Private Sub ClearAll()
'    txtWorkCd = ""
'    lblWorkCdNm.Caption = ""
'    txtFrNo = ""
'    txtToNo = ""
'    txtS.Text = ""
'    txtE.Text = ""
'    dtpAcc.Value = GetSystemDate
'    optCon(0).Value = True
'    optAcc.Visible = True
'    optWS.Visible = False
'
'    Call Clear
'End Sub

'Private Sub SetWorkSheetList()
'    Dim objABOSql As clsABOSql
'    Dim Rs As Recordset
'    Dim i As Long
'
'    Set objABOSql = New clsABOSql
'    Set Rs = objABOSql.LoadWorksheet(ObjSysInfo.BuildingCd)
'    Set objABOSql = Nothing
'
'    lstWSCode.Clear
'
'    If Rs Is Nothing Then Exit Sub
'
'    With Rs
'        For i = 1 To .RecordCount
'            lstWSCode.AddItem .Fields("cdval1").Value & "" & vbTab & .Fields("field1").Value & ""
'            .MoveNext
'        Next i
'    End With
'    Set Rs = Nothing
'End Sub

'Public Sub FocusMe(ctlName As Control)
'    With ctlName
'        .SelStart = 0
'        .SelLength = Len(ctlName)
'    End With
'End Sub

'Public Sub CodeHelp(iKeyAscii As Integer, lstbox As ListBox, sPreStr As String, _
'                    CurCtrl As Control, NextCtrl As Control)
'
'    Dim i%
'    Dim sMadenStr As String
'
'    sPreStr = Trim(sPreStr)
'    '***************  BackSpace 입력시 ( 나머지 문자로 Search )
'    If iKeyAscii = vbKeyBack Then
'        If Len(sPreStr) < 2 Then Exit Sub
'        sMadenStr = Mid(sPreStr, 1, Len(sPreStr) - 1)
'        For i = 0 To lstbox.ListCount
'            If sMadenStr = Mid(lstbox.List(i), 1, Len(sMadenStr)) Then
'                lstbox.Selected(i) = True
'                Exit For
'            End If
'        Next i
'    '**************  방향키 입력시
'    ElseIf iKeyAscii = vbKeyDown Then
'        lstbox.SetFocus
'        lstbox.Selected(0) = True
'
'   '***************  Return 입력시 ( 현재 Cell에 입력한 그대로의 값을 실제 리스트의
''                                     항목과 비교한후 존재하면 검체코드 로드
'    ElseIf iKeyAscii = vbKeyReturn Then
'
'        For i = 0 To lstbox.ListCount - 1
'            If sPreStr = Trim(Mid(lstbox.List(i), 1, _
'                            InStr(1, lstbox.List(i), Chr(vbKeyTab)) - 1)) Then
'
'                Exit For
'            End If
'        Next i
'
'        If i > lstbox.ListCount - 1 Then
'            'MsgBox " 존재하지 않는 코드 입니다."
'            Exit Sub
'        End If
'        NextCtrl.SetFocus
'
'   '***************  Space Bar 입력시( 현재Cell 에 입력한내용을 바탕으로 온전한
'    '                 검사항목을 찾아 Cell에 Write
'    ElseIf iKeyAscii = vbKeySpace Then
'        For i = 0 To lstbox.ListCount - 1
'            If sPreStr = Mid(lstbox.List(i), 1, Len(sPreStr)) Then
'                Exit For
'            End If
'        Next i
'
'        If i > lstbox.ListCount - 1 Then
'            MsgBox " 존재하지 않는 코드입니다."
'            Exit Sub
'        End If
'        CurCtrl.Text = Mid(lstbox.List(i), 1, _
'                             InStr(1, lstbox.List(i), Chr(vbKeyTab)) - 1)
'  '***************  기타 일반적인 문자 입력시
'    Else
'        sMadenStr = sPreStr & Chr(iKeyAscii)
'        For i = 0 To lstbox.ListCount
'            If sMadenStr = Mid(lstbox.List(i), 1, Len(sMadenStr)) Then
'                lstbox.Selected(i) = True
'                Exit For
'            Else
'                lstbox.ListIndex = -1
'            End If
'        Next i
'    End If
'End Sub

Private Function GetAccNoList(aryAccNo() As String) As Boolean
    Dim Rs As Recordset
    Dim objABOSql As clsABOSql
    Dim i As Long
    
    Erase aryAccNo
    
    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.LoadAccList(BB_WORKAREA, Format(dtpAcc.Value, "YYYYMMDD"), Val(txtAccS.Text), Val(txtAccE.Text))
    Set objABOSql = Nothing
    
    If Rs.EOF Then
        GetAccNoList = False
        MsgBox "조회할 자료가 없습니다.", vbExclamation
        Exit Function
    End If
    
    With Rs
        Do Until .EOF
            ReDim Preserve aryAccNo(i)
            aryAccNo(i) = .Fields("workarea").Value & "" & "-" & .Fields("accdt").Value & "" & "-" & .Fields("accseq").Value & ""
            
            i = i + 1
            .MoveNext
        Loop
    End With
    GetAccNoList = True
    
    Set Rs = Nothing
End Function

Private Sub AccQuery()
    Dim aryAccNo() As String
    Dim i As Long

    Dim strWorkseq    As String
    Dim strWorkarea   As String
    Dim strAccdt      As String
    Dim strAccseq     As String

    Dim objPrgBar  As clsProgress
    Dim objPrgBar2  As clsProgress
    Dim objABOSql As clsABOSql
    Dim strLastRst  As String
    Dim Row        As Long
    
    Call ClearOthers
    
    If GetAccNoList(aryAccNo) = False Then Exit Sub

    Set objPrgBar = New clsProgress
    objPrgBar.Container = MainFrm.stsbar
    objPrgBar.max = UBound(aryAccNo)
    objPrgBar.Message = "결과를 입력할 자료를 읽고 있습니다..."

    Set objABOSql = New clsABOSql
    
    tblResult.ReDraw = False
    Row = 1
    For i = LBound(aryAccNo) To UBound(aryAccNo)
        objPrgBar.Value = i
        
        If aryAccNo(i) <> "" Then
            strWorkseq = i + 1
            strWorkarea = medGetP(aryAccNo(i), 1, "-")
            strAccdt = medGetP(aryAccNo(i), 2, "-")
            strAccseq = medGetP(aryAccNo(i), 3, "-")
            
            If QueryByAccNo(objABOSql, strWorkseq, strWorkarea, strAccdt, strAccseq, Row) = True Then
                Row = Row + 1
            End If
        End If
    Next i
    Set objPrgBar = Nothing
    
    Set objPrgBar2 = New clsProgress
    objPrgBar2.Container = MainFrm.stsbar
    objPrgBar2.max = tblResult.DataRowCnt
    objPrgBar2.Message = "최근 결과를 읽고 있습니다..."

    With tblResult
        For i = 1 To .DataRowCnt
            objPrgBar2.Value = i
            strLastRst = ""
            .Row = i
            .Col = TblColumn.tcACCNO
            strLastRst = GetLastResult(Trim(.Value))
            .Col = TblColumn.tcLast: .Value = medGetP(strLastRst, 1, LINE_DIV): .ForeColor = DCM_LightRed: .FontBold = True
            .Col = TblColumn.tcToolRst: .Value = medGetP(strLastRst, 2, LINE_DIV)
        Next

        If .DataRowCnt < 15 Then
            .MaxRows = 15
        Else
            .MaxRows = .DataRowCnt
        End If
        
        .RowHeight(-1) = 12

        If .DataRowCnt > 0 Then
            .Row = 1
            .Col = TblColumn.tcABO
            .Action = ActionActiveCell
            .SetFocus
        End If
        
        .ReDraw = True
    End With

    Set objPrgBar = Nothing
    Set objPrgBar2 = Nothing
    Set objABOSql = Nothing
End Sub

Private Function GetLastResult(ByVal qAccdt As String) As String
    Dim vWorkarea As String
    Dim vAccdt As String
    Dim vAccseq As String
    Dim objSQL As New clsABO

    vWorkarea = medGetP(qAccdt, 1, "-")
    vAccdt = "20" & medGetP(qAccdt, 2, "-")
    vAccseq = medGetP(qAccdt, 3, "-")

    GetLastResult = objSQL.GetLastRst(vWorkarea, vAccdt, vAccseq)

    Set objSQL = Nothing
End Function

Private Sub WSQuery()
    Dim aryAccNo() As String
    Dim i As Long

    Dim workseq As String
    Dim vWorkarea As String
    Dim vAccdt As String
    Dim vAccseq As String

    Dim objPrgBar As clsProgress
    Dim objPrgBar2 As clsProgress
    Dim objABOSql As clsABOSql

    Dim strLastRst As String

    Dim Row As Long

    Call ClearOthers

    If GetWsNoList(aryAccNo) = False Then Exit Sub
    
    Set objPrgBar = New clsProgress
    objPrgBar.Container = MainFrm.stsbar
    objPrgBar.max = UBound(aryAccNo)
    objPrgBar.Message = "결과를 입력할 자료를 읽고 있습니다..."

    Set objABOSql = New clsABOSql
    
    tblResult.ReDraw = False
    Row = 1
    For i = LBound(aryAccNo) To UBound(aryAccNo)
        objPrgBar.Value = i

        workseq = medGetP(aryAccNo(i), 1, "-")
        vWorkarea = medGetP(aryAccNo(i), 2, "-")
        vAccdt = medGetP(aryAccNo(i), 3, "-")
        vAccseq = medGetP(aryAccNo(i), 4, "-")

        If QueryByAccNo(objABOSql, workseq, vWorkarea, vAccdt, vAccseq, Row) = True Then
            Row = Row + 1
        End If
    Next i
    
    Set objPrgBar = Nothing
    
    Set objPrgBar2 = New clsProgress
    objPrgBar2.Container = MainFrm.stsbar
    objPrgBar2.max = tblResult.DataRowCnt
    objPrgBar2.Message = "최근 결과를 읽고 있습니다..."
    
    With tblResult
        For i = 1 To .DataRowCnt
            objPrgBar2.Value = i
            
            strLastRst = ""
            .Row = i
            .Col = TblColumn.tcACCNO
            strLastRst = GetLastResult(Trim(.Value))
            .Col = TblColumn.tcLast: .Value = medGetP(strLastRst, 1, LINE_DIV): .ForeColor = DCM_LightRed: .FontBold = True
            .Col = TblColumn.tcToolRst: .Value = medGetP(strLastRst, 2, LINE_DIV)
        Next

        If .DataRowCnt < 15 Then
            .MaxRows = 15
        Else
            .MaxRows = .DataRowCnt
        End If
        .RowHeight(-1) = 12

        If .DataRowCnt > 0 Then
            .Row = 1
            .Col = TblColumn.tcABO
            .Action = ActionActiveCell
            .SetFocus
        End If
        .ReDraw = True
    End With
    
    Set objABOSql = Nothing
    Set objPrgBar = Nothing
    Set objPrgBar2 = Nothing
End Sub

Private Function GetWsNoList(aryAccNo() As String) As Boolean
    Dim Rs As Recordset
    Dim objABOSql As clsABOSql
    Dim i As Long

    Erase aryAccNo

    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.LoadWsInfo(Format(dtpWS, "YYYYMMDD"), txtWSCd, Val(txtWSS), Val(txtWSE))
    Set objABOSql = Nothing
    
    If Rs.EOF Then
        GetWsNoList = False
        MsgBox "조회할 자료가 없습니다.", vbExclamation
        Exit Function
    End If
    
    With Rs
        Do Until .EOF
            ReDim Preserve aryAccNo(i)
            aryAccNo(i) = .Fields("workseq").Value & "" & "-" & .Fields("workarea").Value & "" & "-" & .Fields("accdt").Value & "" & "-" & .Fields("accseq").Value & ""
            
            i = i + 1
            .MoveNext
        Loop
    End With
    GetWsNoList = True
    
    Set Rs = Nothing
End Function

Private Function QueryByAccNo(ByVal objABOSql As clsABOSql, _
                              ByVal vWorkseq As String, _
                              ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                              ByVal Row As Long) As Boolean
    Dim Rs As Recordset
    Dim i As Long


    Dim sex As String
    Dim ageday As Long

    Dim deptnm As String
    Dim wardnm As String
    Dim doctnm As String
    Dim ptnm As String

    Dim typing As Long
    Dim ABO As String
    Dim Rh As String
    Dim ABOSub As String
    Dim RhSub As String

    Dim RmkCd As String
    Dim comment As String

    '입력된 접수번호 검사----------------------------------------------
    If vWorkarea = "" Or vAccdt = "" Or vAccseq = "" Then
        QueryByAccNo = False
        Exit Function
    End If

    Set objABOSql = New clsABOSql

    Set Rs = objABOSql.GetAccessInfo(vWorkarea, vAccdt, vAccseq)

    If Not (Rs Is Nothing) Then
        With Rs
            If .RecordCount < 1 Then
                QueryByAccNo = False
            Else
                '이 접수번호에 ABO검사항목이 있는지 검사한다.----------------------
                If objABOSql.IsExistABO(vWorkarea, vAccdt, vAccseq) = False Then
                    QueryByAccNo = False
                    Set Rs = Nothing
                    Exit Function
                End If

                '이미 검사가 완료되었는지 검사한다.--------------------------------
                typing = CanEditABOResult(vWorkarea, vAccdt, vAccseq, ABO, Rh, ABOSub, RhSub)
                Select Case typing
                    Case KindOfTyping.NoUse
                        Set Rs = Nothing
                        Set objABOSql = Nothing
                        QueryByAccNo = False
                        Exit Function
                    Case KindOfTyping.AnyUse

                    Case KindOfTyping.FrontUse

                    Case KindOfTyping.BackUse

                    Case Else
                        Set Rs = Nothing
                        Set objABOSql = Nothing
                        Exit Function
                End Select

                sex = .Fields("sex").Value & ""
                ageday = .Fields("ageday").Value & ""

                '코드에 대한 명칭을 불러온다.--------------------------------------
                ptnm = GetPtNm(.Fields("ptid").Value & "")

                deptnm = GetDeptNm(.Fields("deptcd").Value & "")
                wardnm = GetWardNm(.Fields("wardid").Value & "")
                doctnm = GetDoctNm(.Fields("majdoct").Value & "")


                RmkCd = .Fields("rmkcd").Value & ""

                '값을 Display-------------------------------------------------------
                With tblResult
                    .Row = Row

                    If .Row > .MaxRows Then .MaxRows = .MaxRows + 1

                    .Col = TblColumn.tcWORKSEQ:  .Value = vWorkseq
                    .Col = TblColumn.tcACCNO:    .Value = vWorkarea & "-" & Mid(vAccdt, 3) & "-" & vAccseq
                    .Col = TblColumn.tcWORKAREA: .Value = vWorkarea
                    .Col = TblColumn.tcACCDT:    .Value = vAccdt
                    .Col = TblColumn.tcACCSEQ:   .Value = vAccseq

                    .Col = TblColumn.tcABOCD:    .Value = ABO
                    .Col = TblColumn.tcRHCD:     .Value = Rh
                    .Col = TblColumn.tcABOSUBCD: .Value = ABOSub
                    .Col = TblColumn.tcRHSUBCD:  .Value = RhSub

                    .Col = TblColumn.tcABO:      .Value = GetABOFrontNm(ABO)
                    .Col = TblColumn.tcRH:       .Value = GetRHNM(Rh)
                    .Col = TblColumn.tcABOSUB:   .Value = GetABOSUBNM(ABOSub)
                    .Col = TblColumn.tcRHSUB:    .Value = GetRHSUBNM(RhSub)

                    .Col = TblColumn.tcPTNM:     .Value = ptnm & "(" & Rs.Fields("ptid").Value & "" & ")": .ForeColor = DCM_LightBlue: .FontBold = True
                    .Col = TblColumn.tcSEXAGE:   .Value = sex & "/" & (ageday \ 365)
                    .Col = TblColumn.tcDEPT:     .Value = deptnm
                    .Col = TblColumn.tcWARD:     .Value = wardnm
                    .Col = TblColumn.tcRMKCD:    .Value = RmkCd

                    .Col = TblColumn.tcTYPECD:   .Value = typing
                    .Col = TblColumn.tcTYPE
                    Select Case typing
                        Case KindOfTyping.AnyUse
                            .Value = -1
                            .Col = TblColumn.tcTYPE: .Col2 = TblColumn.tcTYPE
                            .Row = Row: .Row2 = Row
                            .BlockMode = True
                            .Lock = False
'                            .Protect = False
                            .BlockMode = False
                        Case KindOfTyping.FrontUse
                            .Value = 0
                            .Col = TblColumn.tcTYPE: .Col2 = TblColumn.tcTYPE
                            .Row = Row: .Row2 = Row
                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False
                        Case KindOfTyping.BackUse
                            .Value = 1
                            .Col = TblColumn.tcTYPE: .Col2 = TblColumn.tcTYPE
                            .Row = Row: .Row2 = Row
                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False
                    End Select
                End With
            End If
        End With

        Set Rs = Nothing

        '코멘트-------------------------------------------------------------
        Set Rs = objABOSql.GetAccComment(vWorkarea, vAccdt, vAccseq)
        If Not (Rs Is Nothing) Then
            With Rs
                If .RecordCount > 0 Then
                    comment = .Fields("rsttxt").Value & ""
                End If
            End With
            Set Rs = Nothing
        End If

        With tblResult
            .Row = Row
            .Col = TblColumn.tcCOMMENT: .Value = comment
        End With
    End If

    QueryByAccNo = True
End Function

Private Function CanEditABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String, _
                                  ByRef pABO As String, ByVal pRh As String, ByRef pABOSub As String, ByVal pRhSub As String) As Long
    '-----------------------------------------------
    '반환값  < 0 : 사용할 수 없다.
    '        = 0 : 선택해서 사용할 수 있다.
    '        = 1 : Front Typing할 수 있다.
    '        = 2 : Back  Typing할 수 있다.
    '-----------------------------------------------
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset

    Dim vfydt  As String    '최종 Verify되면 set된다.
    Dim strVfyid1 As String    'Front Typing 사용자
    Dim strVfydt1 As String    'Front Typing 일자
    Dim strVfytm1 As String    'Front Typing 시간
    Dim strVfyid2 As String    'Back  Typing 사용자
    Dim strVfydt2 As String    'Back  Typing 일자
    Dim strVfytm2 As String    'Back  Typing 시간

    Set objABOSql = New clsABOSql

    Set Rs = objABOSql.GetABOResultInfo(vWorkarea, vAccdt, vAccseq)
    If Rs Is Nothing Then
        CanEditABOResult = KindOfTyping.NoUse
    Else
        With Rs
            If .RecordCount < 1 Then
'                MsgBox "결과내역을 찾을 수 없습니다.전산실로 연락하시기 바랍니다.", vbCritical, Me.Caption
                CanEditABOResult = KindOfTyping.NoUse
            Else
                vfydt = Trim(.Fields("vfydt").Value & "")
                strVfyid1 = Trim(.Fields("vfyid1").Value & "")
                strVfydt1 = Trim(.Fields("vfydt1").Value & "")
                strVfytm1 = Trim(.Fields("vfytm1").Value & "")

                strVfyid2 = Trim(.Fields("vfyid2").Value & "")
                strVfydt2 = Trim(.Fields("vfydt2").Value & "")
                strVfytm2 = Trim(.Fields("vfytm2").Value & "")

                If strVfyid1 = "0" Then strVfyid1 = ""
                If strVfyid2 = "0" Then strVfyid2 = ""

                If strVfydt1 <> "" And strVfydt2 <> "" Then
                    '----------------------------------------------------
                    '이미 최종Verify되었다. 더이상 결과를 등록할 수 없다.
                    '----------------------------------------------------
                    CanEditABOResult = KindOfTyping.NoUse

                    pABO = ""
                    pRh = ""
                    pABOSub = ""
                    pRhSub = ""
                Else
                    '현재 login한 사용자가 사용할 수 있는지 검사한다.

                    If strVfyid1 = "" And strVfyid2 = "" Then
                        '-----------------------------------------------
                        '한번도 입력한 적이 없다. 결과등록을 할 수 있다.
                        '-----------------------------------------------
                        CanEditABOResult = KindOfTyping.AnyUse

                        pABO = ""
                        pRh = ""
                        pABOSub = ""
                        pRhSub = ""
                    ElseIf strVfyid1 <> "" And strVfyid2 = "" Then
                        If strVfyid1 = ObjMyUser.EmpID Then
                            'Front Typing한 사용자이다.
                            If strVfydt1 = "" Then
                                '---------------------------------------
                                '저장만 해놓았다. 결과등록을 할 수 있다.
                                '---------------------------------------
                                CanEditABOResult = KindOfTyping.FrontUse

                                pABO = .Fields("abo1").Value & ""
                                pRh = .Fields("rh1").Value & ""
                                pABOSub = .Fields("abosub").Value & ""
                                pRhSub = .Fields("rhsub").Value & ""
                            Else
                                '---------------------------------------
                                'Verify시켰다. 결과등록을 할 수 없다.
                                '---------------------------------------
                                CanEditABOResult = KindOfTyping.NoUse

                                pABO = ""
                                pRh = ""
                                pABOSub = ""
                                pRhSub = ""
                            End If
                        Else
                            '-------------------------------------
                            'BackTyping으로 결과를 입력할 수 있다.
                            '-------------------------------------
                            CanEditABOResult = KindOfTyping.BackUse

                            pABO = .Fields("abo2").Value & ""
                            pRh = .Fields("rh2").Value & ""
                            pABOSub = .Fields("abosub").Value & ""
                            pRhSub = .Fields("rhsub").Value & ""
                        End If
                    ElseIf strVfyid1 = "" And strVfyid2 <> "" Then
                        If strVfyid2 = ObjSysInfo.EmpID Then
                            'Back Typing한 사용자이다.
                            If strVfydt2 = "" Then
                                '---------------------------------------
                                '저장만 해놓았다. 결과등록을 할 수 있다.
                                '---------------------------------------
                                CanEditABOResult = KindOfTyping.BackUse

                                '--------저장해놓았던 결과를 Display한다.
                                pABO = .Fields("abo2").Value & ""
                                pRh = .Fields("rh2").Value & ""
                                pABOSub = .Fields("abosub").Value & ""
                                pRhSub = .Fields("rhsub").Value & ""
                            Else
                                '---------------------------------------
                                'Verify시켰다. 결과등록을 할 수 없다.
                                '---------------------------------------
                                CanEditABOResult = KindOfTyping.NoUse

                                pABO = ""
                                pRh = ""
                                pABOSub = ""
                                pRhSub = ""
                            End If
                        Else
                            '-------------------------------------
                            'Front Typing으로 결과를 입력할 수 있다.
                            '-------------------------------------
                            CanEditABOResult = KindOfTyping.FrontUse

                            pABO = .Fields("abo1").Value & ""
                            pRh = .Fields("rh1").Value & ""
                            pABOSub = .Fields("abosub").Value & ""
                            pRhSub = .Fields("rhsub").Value & ""
                        End If
                    End If
                End If
            End If
        End With
        Set Rs = Nothing
    End If
    Set objABOSql = Nothing
End Function

Private Sub txtComment_LostFocus()
    With tblResult
        .Row = .ActiveRow
        .Col = TblColumn.tcCOMMENT
        .Value = txtComment.Text
    End With
End Sub

Private Sub txtWSCd_Change()
    If lblWSNm.Caption <> "" Then
        Call ClearWs
        Call ClearOthers
        
        Call GetWsDuration
    End If
End Sub

Private Sub txtWSCd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtWSCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtWSCd.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtWSCd_Validate(Cancel As Boolean)
    Dim objABOSql As clsABOSql
    Dim strWsNm As String
    
    If txtWSCd.Text = "" Then Exit Sub
    
    Set objABOSql = New clsABOSql
    
    strWsNm = objABOSql.GetWSNm(txtWSCd.Text, ObjSysInfo.buildingcd)
    
    If strWsNm = "" Then
        Cancel = True
        MsgBox "존재하지 않는 워크시트 코드입니다.", vbExclamation
    Else
        lblWSNm.Caption = strWsNm
    End If
    
    Set objABOSql = Nothing
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtWSE_Change()
    If tblResult.DataRowCnt <> 0 Then Call ClearOthers
End Sub

Private Sub txtWSS_Change()
    If tblResult.DataRowCnt <> 0 Then Call ClearOthers
End Sub
