VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS816 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Donor Screening 검사 항목 마스터"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS816.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstGroup 
      Appearance      =   0  '평면
      Height          =   1650
      ItemData        =   "frmBBS816.frx":076A
      Left            =   1635
      List            =   "frmBBS816.frx":077D
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtGroup 
      Height          =   315
      Left            =   1635
      TabIndex        =   0
      Top             =   285
      Width           =   2490
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   2640
      Left            =   435
      TabIndex        =   3
      Top             =   720
      Width           =   5175
      _Version        =   196608
      _ExtentX        =   9128
      _ExtentY        =   4657
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   3
      MaxRows         =   10
      OperationMode   =   2
      RestrictRows    =   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS816.frx":07AD
      StartingColNumber=   2
      ScrollBarTrack  =   3
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2820
      Left            =   5685
      TabIndex        =   12
      Top             =   555
      Width           =   4575
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "종료(&X)"
         Height          =   420
         Left            =   3120
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   2040
         Width           =   1260
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   420
         Left            =   1860
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   2040
         Width           =   1260
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
         Caption         =   "삭제(&D)"
         Height          =   420
         Left            =   180
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   2040
         Width           =   1260
      End
      Begin VB.CommandButton cmdPopupList 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frmBBS816.frx":0B82
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   540
         Width           =   300
      End
      Begin VB.ComboBox cboSpcCd 
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
         Left            =   1140
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   1380
         Width           =   3285
      End
      Begin VB.TextBox txtTestCd 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사명"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label lblTestName 
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Height          =   330
         Left            =   1140
         TabIndex        =   15
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체코드"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   4995
      Left            =   435
      TabIndex        =   4
      Top             =   3360
      Width           =   9825
      Begin MSComctlLib.TabStrip tabAppDt 
         Height          =   390
         Left            =   600
         TabIndex        =   5
         Top             =   750
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   688
         Style           =   2
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread tblReference 
         Height          =   2940
         Left            =   585
         TabIndex        =   6
         Tag             =   "35304"
         Top             =   1800
         Width           =   7125
         _Version        =   196608
         _ExtentX        =   12568
         _ExtentY        =   5186
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         BackColorStyle  =   1
         ColHeaderDisplay=   0
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
         MaxCols         =   7
         MaxRows         =   50
         OperationMode   =   1
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmBBS816.frx":110C
         VirtualRows     =   7
      End
      Begin MSComCtl2.DTPicker dtpAppDt 
         Height          =   330
         Left            =   1515
         TabIndex        =   7
         Top             =   1395
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
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
         Format          =   20840451
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpExpDt 
         Height          =   330
         Left            =   4815
         TabIndex        =   8
         Top             =   1380
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd"
         DateIsNull      =   -1  'True
         Format          =   20840451
         CurrentDate     =   36328
      End
      Begin VB.Label lblTestItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사 적격치 정보"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   660
         TabIndex        =   11
         Tag             =   "35131"
         Top             =   300
         Width           =   1860
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폐 기 일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3975
         TabIndex        =   10
         Tag             =   "35214"
         Top             =   1470
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "적 용 일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   9
         Tag             =   "35210"
         Top             =   1455
         Width           =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   585
         X2              =   7560
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   585
         X2              =   7560
         Y1              =   1155
         Y2              =   1155
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사그룹"
      Height          =   180
      Left            =   795
      TabIndex        =   21
      Top             =   345
      Width           =   720
   End
End
Attribute VB_Name = "frmBBS816"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form명   : frmBBS811
'|  2. 기  능   : Doner Screening 검사 항목 마스터
'|  3. 작성자   : 김 동열
'|  4. 작성일   : 2000.11.28
'|
'|  CopyRight(C) 2000 대련엠티에스
'+--------------------------------------------------------------------------------------+
Option Explicit


Private objSql As clsBBSMSTStatement
Private WithEvents objListpop As clsPopUpList
Attribute objListpop.VB_VarHelpID = -1

Private onPgm As Boolean


Private Sub cmdDelete_Click()
    Dim objDonorTest As clsDonorTest
    Dim strTmp       As VbMsgBoxResult
    
    strTmp = MsgBox("삭제하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbCancel Then Exit Sub
    
    Set objDonorTest = New clsDonorTest
    
On Error GoTo cmdDelete_error

    DBConn.BeginTrans
    
    If objDonorTest.Delete(txtGroup, txtTestCd) = False Then GoTo cmdDelete_error
    
    DBConn.CommitTrans
    
    Call TblDisplay
    
    Exit Sub
    
cmdDelete_error:
    DBConn.RollbackTrans
'
'    'Dim Rs As New Recordset
'    Dim strTmp As VbMsgBoxResult
'
'    strTmp = MsgBox("삭제하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
'    If strTmp = vbCancel Then
'        Exit Sub
'    Else '저장
'        Set objSql = New clsBBSMSTStatement
''        objSql.setDbConn DbConn
'        If objSql.DeleteB003(Trim(txtTestCd)) = True Then
'            MsgBox "삭제하였습니다..", vbInformation, Me.Caption
'            TblDisplay
'            Clear
'            tblClear
'            txtTestCd.SetFocus
'            Exit Sub
'        End If
'    End If
'    Set objSql = Nothing
End Sub

Private Sub cmdExit_Click()
    If Not objSql Is Nothing Then
        Set objSql = Nothing
    End If
    Unload Me
End Sub

Private Sub cmdPopupList_Click()
    'Dim Rs As New Recordset
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    '리스트 팝업을 불러오자...
    Set objListpop = New clsPopUpList
    objListpop.Connection = DBConn
'    objListpop.BackColor = Me.BackColor
    objListpop.Tag = "TestCd"
    objListpop.FormCaption = "검사코드 찾기"
    Call objListpop.LoadPopup(objSql.LoadPopup("0")) ', 3200, 10700)
    
    Set objSql = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim objDonorTest As clsDonorTest
    Dim strSpecCd    As String
    Dim strSpecNm    As String
    
    
    Set objDonorTest = New clsDonorTest
    
On Error GoTo cmdSave_error

    DBConn.BeginTrans
    
    strSpecCd = medGetP(cboSpcCd.Text, 1, " ")
    strSpecNm = Mid(cboSpcCd.Text, Len(strSpecCd) + 1)
    
    If objDonorTest.Save(txtGroup, txtTestCd, strSpecCd, strSpecNm) = True Then
        Call TblDisplay
    Else
        GoTo cmdSave_error
    End If
    
    DBConn.CommitTrans
    
    Exit Sub
    
cmdSave_error:
    DBConn.RollbackTrans
    
'    Dim Rs As New Recordset
'    Dim strTmp As VbMsgBoxResult
'    Dim strTmp1 As VbMsgBoxResult
'
'    Set objSql = New clsBBSMSTStatement
''    objSql.setDbConn DbConn
'    Set Rs = objSql.LoadB003(Trim(txtTestCd))
'    If Rs.EOF = False Then
'        strTmp1 = MsgBox("수정하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
'        If strTmp1 = vbCancel Then
'            Set Rs = Nothing
'            Set objSql = Nothing
'            Exit Sub
'        Else '수정
'            If objSql.InsertB003(Trim(txtTestCd), medGetP(cboSpcCd.Text, 1, " "), Trim(lblTestName), False) = True Then
'                MsgBox "수정하였습니다.", vbInformation, Me.Caption
'                TblDisplay
'                Call cboSpcCd_Click
'            End If
'        End If
'    Else
'    '저장여부 확인...
'        strTmp = MsgBox("저장하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
'        If strTmp = vbCancel Then
'            Set objSql = Nothing
'            Exit Sub
'        Else '저장
'            If objSql.InsertB003(Trim(txtTestCd), medGetP(cboSpcCd.Text, 1, " "), Trim(lblTestName), True) = True Then
'                MsgBox "저장성공하였습니다.", vbInformation, Me.Caption
'                TblDisplay
'                Clear
'                tblClear
'            End If
'        End If
'    End If
'    Set Rs = Nothing
'    Set objSql = Nothing
End Sub

Private Sub GetGroup()
    Dim objDonorTest As clsDonorTest
    Dim strGroup()   As String
    Dim icnt         As Long
    Dim i            As Long
    
    Set objDonorTest = New clsDonorTest
    icnt = objDonorTest.GetGroup(strGroup)
    
    lstGroup.Clear
    For i = 1 To icnt
        lstGroup.AddItem strGroup(i - 1)
    Next i
    
    Set objDonorTest = Nothing
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
    
    Call GetGroup
'    Call TblDisplay
End Sub

Private Sub Form_Load()
    dtpAppDt.Enabled = False
    dtpExpDt.Enabled = False
    Clear
End Sub

Private Sub lstGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    onPgm = True
    txtGroup = lstGroup.Text
    txtGroup.SelStart = Len(txtGroup)
    txtGroup.SelLength = Len(txtGroup)
    onPgm = False
    
    lstGroup.Visible = False
    
    Call TblDisplay
End Sub

Private Sub objListpop_SendCode(ByVal SelString As String)
    Dim i As Integer
    Dim RS As Recordset
    '리스트박스에 있는내용을 가져오자..
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    txtTestCd.Text = medGetP(SelString, 1, ";")
    lblTestName.Caption = medGetP(SelString, 2, ";")
    If txtTestCd.Text = "" Then Exit Sub
        lblTestName.Caption = ""
        Set RS = objSql.GetTestNm(Trim(txtTestCd.Text))
        If Not RS.EOF Then
            lblTestName.Caption = RS.Fields("testnm").Value & ""
            Call LoadSpeimen(Trim(txtTestCd.Text))
            If cboSpcCd.ListCount > 0 Then
               cboSpcCd.ListIndex = 0
               cboSpcCd.SetFocus
            End If
        Else
            MsgBox "검사항목이 없습니다.", vbInformation, Me.Caption
            txtTestCd.Text = ""
            txtTestCd.SetFocus
        End If
    If Not RS Is Nothing Then
        Set RS = Nothing
    End If
    Set objSql = Nothing
    Set objListpop = Nothing
End Sub

Public Function TblDisplay()
    Dim objDonorTest As clsDonorTest
    Dim strTestCd()  As String
    Dim icnt         As Long
    Dim i            As Long
    
    With tblList
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClear
        .BlockMode = False
    End With
    
    If txtGroup = "" Then Exit Function
    
    Set objDonorTest = New clsDonorTest
    icnt = objDonorTest.GetTestOfGroup(txtGroup, strTestCd)
    
    With tblList
         .MaxRows = icnt
        For i = 1 To icnt
            .Row = i: .Row2 = i
            .Col = 1: .Col2 = 3
            .BlockMode = True
            .Clip = strTestCd(i - 1)
            .BlockMode = False
        Next i
    End With
    
    Set objDonorTest = Nothing
End Function

Private Sub TblQuery()
    '스프레드 내용을 가져오자..
    On Error Resume Next
    
    With tblList
        .Row = .ActiveRow
        .Col = 1: txtTestCd = .Value
        .Col = 2: lblTestName = .Value
        Call LoadSpeimen(Trim(txtTestCd.Text))
        If cboSpcCd.ListCount > 0 Then
           cboSpcCd.ListIndex = 0
           cboSpcCd.SetFocus
        End If
       .Col = 3:
        cboSpcCd.ListIndex = medComboFind(cboSpcCd, medGetP(.Value, 1, " "))
    End With
End Sub

Public Sub LoadSpeimen(sTestCd As String)
    Dim RS As New Recordset
    Dim i As Integer
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.getSpcs(sTestCd)
    cboSpcCd.Clear
    Do Until RS.EOF
        cboSpcCd.AddItem "" & RS.Fields("spccd").Value & Space(1) & RS.Fields("spcnm").Value & ""    ', Val(OraDS.Fields("Seq").Value) - 1
        RS.MoveNext
    Loop
End Sub


Private Sub objListpop_SelectedItem(ByVal pSelectedItem As String)
    Dim i As Integer
    Dim RS As Recordset
    '리스트박스에 있는내용을 가져오자..
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    txtTestCd.Text = medGetP(pSelectedItem, 1, ";")
    lblTestName.Caption = medGetP(pSelectedItem, 2, ";")
    If txtTestCd.Text = "" Then Exit Sub
        lblTestName.Caption = ""
        Set RS = objSql.GetTestNm(Trim(txtTestCd.Text))
        If Not RS.EOF Then
            lblTestName.Caption = RS.Fields("testnm").Value & ""
            Call LoadSpeimen(Trim(txtTestCd.Text))
            If cboSpcCd.ListCount > 0 Then
               cboSpcCd.ListIndex = 0
               cboSpcCd.SetFocus
            End If
        Else
            MsgBox "검사항목이 없습니다.", vbInformation, Me.Caption
            txtTestCd.Text = ""
            txtTestCd.SetFocus
        End If
    If Not RS Is Nothing Then
        Set RS = Nothing
    End If
    Set objSql = Nothing
    Set objListpop = Nothing
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    TblQuery
End Sub

Private Sub Clear()
    '깨끗이..
    txtTestCd.Text = ""
    lblTestName.Caption = ""
    cboSpcCd.Clear
End Sub

Private Sub ClearTable()
    With tblList
        .Row = -1
        .Col = -1
        .Text = ""
    End With
End Sub

Private Sub txtGroup_Change()
    Dim Index As Long
    
    
    '사용자의 키입력이 아닌경우에는 제외
    If onPgm = True Then Exit Sub
    
    If lstGroup.Visible = False Then lstGroup.Visible = True
    
    '입력하는데로 리스트 박스에서 해당 줄을 찾는다.
    If txtGroup = "" Then
        lstGroup.ListIndex = -1
    Else
        Index = medListFind(lstGroup, txtGroup)
        lstGroup.ListIndex = Index
    End If
End Sub

Private Sub txtGroup_GotFocus()
    If lstGroup.Visible = False Then lstGroup.Visible = True
End Sub

Private Sub txtGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Index As Long
    
    Select Case KeyCode
        Case vbKeyDown
            If lstGroup.Visible = False Then lstGroup.Visible = True
            With lstGroup
                If .ListIndex < (.ListCount - 1) Then .ListIndex = .ListIndex + 1
                onPgm = True
                txtGroup = .Text
                txtGroup.SelStart = Len(txtGroup)
                txtGroup.SelLength = Len(txtGroup)
                onPgm = False
            End With
            KeyCode = 0
        Case vbKeyUp
            If lstGroup.Visible = False Then lstGroup.Visible = True
            With lstGroup
                If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
                onPgm = True
                txtGroup = .Text
                txtGroup.SelStart = Len(txtGroup)
                txtGroup.SelLength = Len(txtGroup)
                onPgm = False
            End With
            KeyCode = 0
        Case vbKeyReturn
            If txtGroup = "" Then Exit Sub
            
            With lstGroup
                '신규
                If .ListIndex < 0 Then
                    If MsgBox("새로운 그룹으로 등록하시겠읍니까?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                        txtGroup.SelStart = 0
                        txtGroup.SelLength = Len(txtGroup)
                        KeyCode = 0
                        Exit Sub
                    End If
                    .AddItem txtGroup
                    .ListIndex = medListFind(lstGroup, txtGroup)
                End If
                    
                Call lstGroup_MouseDown(0, 0, 0, 0)
            End With
            KeyCode = 0
    End Select
End Sub

Private Sub txtGroup_LostFocus()
    If lstGroup.Visible = True Then lstGroup.Visible = False
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTestCd_LostFocus()
    Dim RS As Recordset
    
    If txtTestCd.Text = "" Then Exit Sub
    lblTestName.Caption = ""
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.GetTestNm(Trim(txtTestCd.Text))
    If Not RS.EOF Then
        lblTestName.Caption = RS.Fields("testnm").Value & ""
        Call LoadSpeimen(Trim(txtTestCd.Text))
        If cboSpcCd.ListCount > 0 Then
           cboSpcCd.ListIndex = 0
'           cboSpcCd.SetFocus
        End If
    Else
        MsgBox "검사항목이 없습니다.", vbInformation, Me.Caption
        txtTestCd.Text = ""
'        txtTestCd.SetFocus
    End If
    If Not RS Is Nothing Then
        Set RS = Nothing
    End If
End Sub

Private Sub cboSpcCd_Click()
    Dim strSql As String
    Dim strSpcCd As String

    strSpcCd = medGetP(cboSpcCd.Text, 1, " ")
    Call LoadLab003AppDt(txtTestCd.Text, strSpcCd)
    If tabAppDt.Tabs.Count > 0 Then
        tabAppDt.Tabs(1).Selected = True
    Else
        tblClear
    End If
End Sub

Private Sub tblClear()
    tabAppDt.Tabs.Clear
    With tblReference
            .Row = -1
            .Col = -1
            .Text = ""
    End With
End Sub

Public Sub LoadLab003AppDt(sTestCd As String, sSpcCd As String)
    Dim RS As Recordset
    Dim i As Integer
    Dim strKey As String
    Dim strCaption As String
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.getApplydt(sTestCd, sSpcCd)
    
    i = 0
    tabAppDt.Tabs.Clear
    Do Until RS.EOF
        i = i + 1
        strKey = RS.Fields("applydt").Value & ""
        strCaption = Format(strKey, "##-##-##")
        tabAppDt.Tabs.Add i, , strCaption
        RS.MoveNext
    Loop
    If Not RS Is Nothing Then
        Set RS = Nothing
    End If
End Sub

Private Sub tabAppDt_Click()
   Dim strSql As String
   Dim strAppDt As String
   Dim strSpcCd As String
   
   strAppDt = Format(tabAppDt.SelectedItem.Caption, CS_DateDbFormat)
   strSpcCd = medGetP(cboSpcCd.Text, 1, " ")
   If Trim(txtTestCd.Text) = "" Then
        Exit Sub
   End If
   Call LoadLab003(txtTestCd.Text, strSpcCd, strAppDt)
End Sub

Public Sub LoadLab003(sTestCd As String, sSpcCd As String, sAppDt As String)
    Dim i As Integer
    Dim MyReference As New clsBBSMSTStatement
    Dim RS As Recordset
    Dim sgTmp As Single
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.getReference(sTestCd, sSpcCd, sAppDt)
    Call medClearTable(tblReference, False, False)
    i = 0
    dtpAppDt.Value = Format(CStr(RS.Fields("applydt").Value & ""), "##-##-##")
    
    With tblReference
        .MaxRows = 50
        .Row = 0
        Do Until RS.EOF
            i = i + 1
            If .Row = .MaxRows Then .MaxRows = .MaxRows + i
            .Row = .Row + 1
            '.TypeHAlign = TypeHAlignCenter
            .Col = 1:
            Select Case "" & RS.Fields("applysex").Value
                Case "M":
                .TypeComboBoxCurSel = 0
                Case "F":
                .TypeComboBoxCurSel = 1
                Case "B":
                .TypeComboBoxCurSel = 2
                Case "U":
                .TypeComboBoxCurSel = 3
            End Select
            .Col = 2: .Value = "" & RS.Fields("agefrom").Value
            .Col = 3: .Value = "" & RS.Fields("ageto").Value
            .Col = 4: .Value = "" & RS.Fields("refvalfrom").Value
            .Col = 5: .Value = "" & RS.Fields("refvalto").Value
            .Col = 6: .Value = "" & RS.Fields("refcd").Value
            sgTmp = .MaxTextRowHeight(.Row)
            If sgTmp > 13.3 Then
                .RowHeight(.Row) = sgTmp
                Else
                .RowHeight(.Row) = 13.3
            End If
            dtpAppDt.Value = Format(RS.Fields("applydt").Value & "", "##-##-##")
            If Trim(RS.Fields("expdt").Value & "") = "" Then
                dtpExpDt.Value = ""
                dtpExpDt.Enabled = False
            Else
                dtpExpDt.Value = Format(RS.Fields("expdt").Value & "", "##-##-##")
                dtpExpDt.Enabled = False
            End If
            RS.MoveNext
        Loop
    End With
    
NoData:
    'MyItem.RemoveParameters
    If Not RS Is Nothing Then
    Set RS = Nothing
    Set objSql = Nothing
    End If
    Exit Sub
End Sub





