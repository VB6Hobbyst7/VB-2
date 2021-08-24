VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm265BarPrint 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "특수화학 바코드 일괄 재출력"
   ClientHeight    =   9195
   ClientLeft      =   -45
   ClientTop       =   -135
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSQL 
      BackColor       =   &H00E0E0E0&
      Caption         =   "조회(&F)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   60
      TabIndex        =   10
      Top             =   360
      Width           =   14385
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13590
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   19
         Tag             =   "DeptCd"
         Top             =   195
         Width           =   285
      End
      Begin VB.TextBox txtTestCd 
         Height          =   315
         Left            =   11700
         TabIndex        =   18
         Top             =   195
         Width           =   1875
      End
      Begin VB.ComboBox cboWA 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   7875
         Style           =   2  '드롭다운 목록
         TabIndex        =   11
         Top             =   195
         Width           =   2565
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   765
         TabIndex        =   12
         Top             =   195
         Width           =   2610
         _ExtentX        =   4604
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
         Format          =   104660992
         CurrentDate     =   36238
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   3780
         TabIndex        =   13
         Top             =   195
         Width           =   2820
         _ExtentX        =   4974
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
         Format          =   104660992
         CurrentDate     =   36391
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   510
         Left            =   6705
         TabIndex        =   14
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
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
         Caption         =   "WorkArea"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   510
         Left            =   10530
         TabIndex        =   15
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
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
         Caption         =   "검사명"
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3405
         TabIndex        =   17
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CommandButton CmdBarcode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "재출력(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   45
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   " 특수화학 바코드 일괄 재출력"
      Appearance      =   0
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00DBE6E6&
      Height          =   7425
      Left            =   75
      TabIndex        =   0
      Top             =   975
      Width           =   14385
      Begin VB.CommandButton Command1 
         Caption         =   "바코드장수변경"
         Height          =   330
         Left            =   11970
         TabIndex        =   9
         Top             =   135
         Width           =   1455
      End
      Begin VB.ComboBox cboBarCnt 
         Height          =   300
         Left            =   13455
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   135
         Width           =   825
      End
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체선택(&A)"
         ForeColor       =   &H00553755&
         Height          =   255
         Left            =   225
         TabIndex        =   7
         Top             =   210
         Width           =   1350
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   6840
         Left            =   210
         TabIndex        =   6
         Tag             =   "10114"
         Top             =   480
         Width           =   14025
         _Version        =   196608
         _ExtentX        =   24739
         _ExtentY        =   12065
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         GridColor       =   14737632
         MaxCols         =   31
         MaxRows         =   21
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis265.frx":0000
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   21
      End
   End
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00FFF9F7&
      Height          =   2220
      Left            =   4050
      TabIndex        =   2
      Top             =   510
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "frm265BarPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRstDic As New clsDictionary
Private objMicRst As New clsLISMicResult
'Private objMicCul As New clsLISMicCulture
'Private objMicSql As New clsLISSqlMicRst
Private objSQL  As New clsLISSqlStatistic
Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
Private fWorkSheet() As tpMicWorkSheet

Private fNGCode() As Variant
Private SelFg As Boolean

Private Const fSCItem = &H8080FF          ' Worksheet List 에서 선택된 Lab-No
Private fGCItem As Long

Public Event LastFormUnload()

Private Sub GetWorkAreaCombo()
    
    Dim sSqlGetWA As String
    Dim rsGetWA As Recordset
    Dim i%
    
    Set rsGetWA = New Recordset
    rsGetWA.Open objSQL.GetWACd, DBConn
    
    cboWA.Clear
    For i = 1 To rsGetWA.RecordCount
        cboWA.AddItem "" & rsGetWA.Fields("WACd").Value & "   " & _
                            "" & rsGetWA.Fields("WANm").Value
        rsGetWA.MoveNext
    Next i

    Set rsGetWA = Nothing

End Sub

Private Sub chkSelAll_Click()
   
    Dim i As Integer
    
    SelFg = True
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            .Value = chkSelAll.Value
        Next
    End With
    SelFg = False
 
End Sub

Private Sub CmdBarcode_Click()
    Dim i As Long
    Dim jj As Long
    Dim objBar As New clsBarcode
    Dim tmpLabNo As Variant
    Dim TestNames As String
    Dim BarBuffer(1 To 15) As String
    Dim AccFg As Boolean
    
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    jj = 0
    TestNames = ""
    
    Call MouseRunning

    
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            If .Value = 1 Then
                jj = jj + 1
                    
                 Erase BarBuffer
                 .Col = 22:
                             If P_ApplyBuildingInfo Then
                                 BarBuffer(1) = Mid(.Value, 1, 2)        '건물명
                             Else
                                 BarBuffer(1) = LABName
                             End If
                 .Col = 20:  TestNames = .Value
                 .Col = 15:  BarBuffer(2) = .Value           'WorkArea
                 .Col = 18:  BarBuffer(3) = Mid(.Value, 3)   'AccDt
                 .Col = 16:  BarBuffer(4) = .Value           'AccSeq
                 .Col = 21:  BarBuffer(5) = .Value           'SpcNo
                 .Col = 4:   BarBuffer(6) = .Value           '환자ID
                 .Col = 3:   BarBuffer(7) = Mid(.Value, 1, 3)   '환자명
                 .Col = 14:  BarBuffer(8) = .Value           '검체명
                 .Col = 17:  BarBuffer(9) = .Value           '보관코드
                 .Col = 19:  BarBuffer(10) = .Value           'StatFg
                 .Col = 29:
                         If .Value = "" Then                 '진료과코드
                               .Col = 24: BarBuffer(11) = .Value
                         Else
                             BarBuffer(11) = .Value        '병동ID
                             .Col = 23
                             If Trim(.Value) <> "" Then
                                 BarBuffer(11) = BarBuffer(11) & "/" & .Value
                             End If
                         End If
                 .Col = 10:  BarBuffer(12) = Mid(.Value, 5, 2) & "/" & Mid(.Value, 7, 2)      '처방일
                 .Col = 26: BarBuffer(13) = .Value           '희망채혈일시
                  BarBuffer(14) = TestNames                  '검사명
                 .Col = 30: BarBuffer(15) = cboBarCnt.Text '.Value           '라벨출력장수
                 .Col = 25: AccFg = IIf(.Value >= enStsCd.StsCd_LIS_Accession, True, False)  'Status
            
                 Call objBar.Label_PrintOut(BarBuffer(1), BarBuffer(2), BarBuffer(3), BarBuffer(4), BarBuffer(5), BarBuffer(6), _
                                                           BarBuffer(7), BarBuffer(8), BarBuffer(9), BarBuffer(10), BarBuffer(11), _
                                                           BarBuffer(12), BarBuffer(13), BarBuffer(14), BarBuffer(15), AccFg)
            End If
        Next
    End With
    If jj = 0 Then
        MsgBox "재출력 할 리스트를 선택하여 주세요", vbCritical, "바코드 출력오류"
        MouseDefault
        Set objBar = Nothing
        Exit Sub
    End If
   
    Call objBar.Label_FormFeed
    
'    Call cmdClear_Click
    MouseDefault
'    lblMessage.Caption = ""
    Set objBar = Nothing
End Sub

Private Sub cmdClear_Click()
    ScreenClear
    chkSelAll.Value = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set objRstDic = Nothing
End Sub

Private Sub cmdHelpList_Click()
    Dim objTestDiv As New clsDictionary
    Dim objRs As Recordset
    
    If cboWA.ListIndex = -1 Then Exit Sub
    
    Set objListPop = New clsPopUpList
    
'    Call TxtClear
'    Call txtTblClear
    
    With objTestDiv
        .Clear
        .FieldInialize "검사항목코드", "검사명,구분"
        Set objRs = New Recordset
        objRs.Open objSQL.GetWAvsTest(medGetP(cboWA.Text, 1, " ")), DBConn
        While Not objRs.EOF
            .AddNew objRs.Fields("testcd").Value & "", objRs.Fields("abbrnm10").Value & COL_DIV & objRs.Fields("testdiv").Value
            objRs.MoveNext
        Wend
    End With
    Set objRs = Nothing
    
    With objListPop
        .Connection = DBConn
        .FormCaption = "검사항목 조회"
        .ColumnHeaderText = "검사항목코드;검사명;구분"
        .ColumnHeaderWidth = "1440;1260.284;750.0473"
        .FormWidth = 3900
        .LoadPopUp objSQL.GetWAvsTest(medGetP(cboWA.Text, 1, " "))
        txtTestCd.Text = medGetP(.SelectedString, 1, ";")
'        lblTestNm.Caption = medGetP(.SelectedString, 2, ";")
'        lblTestDiv.Caption = medGetP(.SelectedString, 3, ";")
'        Call GetRstCdList
    End With
    Set objListPop = Nothing
End Sub

Private Sub cmdSQL_Click()
    Dim strWork As String
    Dim strFrom As String
    Dim strTo   As String
    Dim strTestcd As String
    Dim sStartDate As String
    Dim sEndDate As String
    
    sStartDate = Format(dtpStart.Value, CS_DateDbFormat)
    sEndDate = Format(dtpEnd.Value, CS_DateDbFormat)
    strWork = Mid(cboWA, 1, 2)
    strTestcd = txtTestCd.Text
    
   Call DisplayData(sStartDate, sEndDate, strWork, strTestcd)

End Sub

Private Sub Command1_Click()
    Dim iCnt As Integer
    
    With tblOrdSheet
        For iCnt = 1 To .MaxRows
            .SetText 30, iCnt, cboBarCnt.Text
        Next
    End With
End Sub

Private Sub Form_Load()
    tblOrdSheet.Row = 1: tblOrdSheet.Col = 1: fGCItem = tblOrdSheet.ForeColor
    
    '===> 요부분입니다...
'    Call objMicRst.LoadWorkSheetCode(MWS_ForCulture, cboWSCode, fWorkSheet)
    
    cboBarCnt.Text = ""
    cboBarCnt.AddItem "1"
    cboBarCnt.AddItem "2"
    cboBarCnt.AddItem "3"
    cboBarCnt.ListIndex = 0
    
    Call ScreenClear
    
    dtpStart.Value = Format(Now, "yyyy-mm-dd")
    dtpEnd.Value = Format(Now, "yyyy-mm-dd")
        
    Call GetWorkAreaCombo
    
End Sub

Private Sub ScreenClear()

    tblOrdSheet.MaxRows = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRstDic = Nothing
    Set objSQL = Nothing
    Set objListPop = Nothing
End Sub

Private Sub DisplayData(ByVal pStart As String, ByVal pEnd As String, ByVal pWork As String, ByVal pTestCd As String)
    
    Dim strBuildDtTm As String, strRcvDtTm As String

    tblOrdSheet.MaxRows = 0
    DoEvents
    
    Call objMicRst.DisPlayReBarCodeList1(tblOrdSheet, pStart, pEnd, pWork, pTestCd)
    
End Sub
