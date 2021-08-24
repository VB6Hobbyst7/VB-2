VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm264MicBarPrint 
   BackColor       =   &H00DBE6E6&
   Caption         =   "미생물 바코드 일괄 재출력"
   ClientHeight    =   9195
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14670
   WindowState     =   2  '최대화
   Begin VB.CommandButton CmdBarcode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "재출력(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   75
      TabIndex        =   2
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
      Caption         =   " 미생물 바코드 일괄 재출력"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   810
      Left            =   75
      TabIndex        =   1
      Top             =   285
      Width           =   14370
      Begin VB.ComboBox cboWSCode 
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
         ItemData        =   "Lis264.frx":0000
         Left            =   1965
         List            =   "Lis264.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   225
         Width           =   1995
      End
      Begin VB.TextBox txtWSUnit 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   3960
         TabIndex        =   4
         Text            =   "19990005"
         Top             =   225
         Width           =   3135
      End
      Begin VB.CommandButton cmdWSList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7125
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   225
         Width           =   345
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   14
         Left            =   210
         TabIndex        =   13
         Top             =   225
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "Work Sheet Unit"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00DBE6E6&
      Height          =   7425
      Left            =   75
      TabIndex        =   0
      Top             =   1035
      Width           =   14385
      Begin VB.CommandButton Command1 
         Caption         =   "바코드장수변경"
         Height          =   330
         Left            =   11970
         TabIndex        =   16
         Top             =   135
         Width           =   1455
      End
      Begin VB.ComboBox cboBarCnt 
         Height          =   300
         Left            =   13455
         TabIndex        =   15
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
         TabIndex        =   12
         Top             =   210
         Width           =   1350
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   6840
         Left            =   210
         TabIndex        =   11
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
         SpreadDesigner  =   "Lis264.frx":0004
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   21
      End
      Begin MedControls1.LisLabel lblMedia 
         Height          =   315
         Left            =   4965
         TabIndex        =   14
         Top             =   135
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "Work Sheet Unit"
         Appearance      =   0
      End
      Begin VB.Label lblMediaList 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Bap,"
         Height          =   180
         Left            =   6735
         TabIndex        =   10
         Top             =   195
         Width           =   3615
      End
   End
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00FFF9F7&
      Height          =   2220
      Left            =   4050
      TabIndex        =   6
      Top             =   870
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "frm264MicBarPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRstDic As New clsDictionary
Private objMicRst As New clsLISMicResult
Private objMicCul As New clsLISMicCulture
Private objMicSql As New clsLISSqlMicRst

Private fWorkSheet() As tpMicWorkSheet
Private fNGCode() As Variant
Private SelFg As Boolean

Private Const fSCItem = &H8080FF          ' Worksheet List 에서 선택된 Lab-No
Private fGCItem As Long

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

Private Sub cboWSCode_Click()
    
    Dim i As Integer
    
    If cboWSCode.ListIndex < 0 Then Exit Sub
    
    ScreenClear
    
    txtWSUnit = ""
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    txtWSUnit.SetFocus

    Call objMicRst.LoadMedia(fWorkSheet(cboWSCode.ListIndex).WsCode, lblMediaList)
    
    lblMedia.Visible = IIf(lblMediaList.Caption = "", False, True)
    
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
                 .Col = 30: BarBuffer(15) = .Value           '라벨출력장수
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
    
    Call cmdClear_Click
    MouseDefault
'    lblMessage.Caption = ""
    Set objBar = Nothing
End Sub

Private Sub cmdClear_Click()
    ScreenClear
    chkSelAll.Value = 0
    lblMedia.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set objRstDic = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
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
    Call objMicRst.LoadWorkSheetCode(MWS_ForCulture, cboWSCode, fWorkSheet)
    
    cboBarCnt.Text = ""
    cboBarCnt.AddItem "1"
    cboBarCnt.AddItem "2"
    cboBarCnt.AddItem "3"
    cboBarCnt.ListIndex = 1
    
    cboWSCode.ListIndex = -1: Erase fNGCode
    txtWSUnit.Text = ""
    Call ScreenClear
    
    lblMedia.Visible = False
End Sub

Private Sub ScreenClear()

    lblMediaList.Caption = "": tblOrdSheet.MaxRows = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRstDic = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
End Sub



Private Sub txtWSUnit_KeyPress(KeyAscii As Integer)
    
    Dim iWSIndex As Integer
        
    If KeyAscii = vbKeyReturn Then
    
        iWSIndex = cboWSCode.ListIndex

        If ExistWS(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text) Then
            Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, fWorkSheet(iWSIndex).WsRstType)
        Else
            Call ScreenClear
        End If
        
    End If

End Sub

Private Sub lstWSUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iListIndex As Integer, iWSIndex As Integer
    
    MouseRunning
    
    iWSIndex = cboWSCode.ListIndex
    iListIndex = lstWSUnit.ListIndex
    
    If Button = vbLeftButton And iListIndex >= 0 Then
        txtWSUnit.Text = medGetP(lstWSUnit.List(iListIndex), 1, " ")
        Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, fWorkSheet(iWSIndex).WsRstType)
    End If
    
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    
    MouseDefault

End Sub

Private Sub cmdWSList_Click()
    
    
    Dim sWsCd As String

    If cboWSCode.ListIndex < 0 Then Exit Sub

    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    objMicRst.LoadMicWorkList sWsCd, lstWSUnit
    If lstWSUnit.ListCount <= 0 Then Exit Sub
    
    lstWSUnit.ListIndex = 0
    lstWSUnit.Visible = True
    lstWSUnit.ZOrder

End Sub

Private Sub DisplayData(ByVal pWsCd As String, ByVal pWsUnit As String, ByVal sRTs As String)
    
    Dim strBuildDtTm As String, strRcvDtTm As String

    tblOrdSheet.MaxRows = 0
    DoEvents
    
    Call objMicRst.DisPlayReBarCodeList(tblOrdSheet, pWsCd, pWsUnit, sRTs)
    
End Sub
