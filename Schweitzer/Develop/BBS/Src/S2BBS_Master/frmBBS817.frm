VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS817 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Donor Screening 검사 적격치 마스터"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS817.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   5520
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "128"
      Top             =   7860
      Width           =   1230
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   420
      Left            =   4200
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "128"
      Top             =   7860
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1635
      Left            =   1680
      TabIndex        =   11
      Top             =   360
      Width           =   7455
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
         Left            =   1860
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   1080
         Width           =   5265
      End
      Begin VB.CheckBox chkTestCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "적격치가 설정된 검사항목만 찾기"
         Height          =   195
         Left            =   3960
         TabIndex        =   18
         Top             =   300
         Width           =   3135
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1860
         TabIndex        =   13
         Top             =   660
         Width           =   1395
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
         Left            =   1365
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frmBBS817.frx":076A
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체코드"
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
         Left            =   420
         TabIndex        =   20
         Tag             =   "35303"
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label lblTestName 
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Height          =   330
         Left            =   3315
         TabIndex        =   15
         Top             =   660
         Width           =   3795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
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
         Left            =   420
         TabIndex        =   14
         Tag             =   "35302"
         Top             =   705
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   5715
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   7485
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가(&A)"
         Height          =   420
         Left            =   3780
         Style           =   1  '그래픽
         TabIndex        =   3
         Tag             =   "35301"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "수정(&E)"
         Height          =   420
         Left            =   4920
         Style           =   1  '그래픽
         TabIndex        =   2
         Tag             =   "135"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소(&U)"
         Height          =   420
         Left            =   6060
         Style           =   1  '그래픽
         TabIndex        =   1
         Tag             =   "35301"
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.TabStrip tabAppDt 
         Height          =   390
         Left            =   270
         TabIndex        =   4
         Top             =   795
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
         Height          =   3240
         Left            =   225
         TabIndex        =   5
         Tag             =   "35304"
         Top             =   1860
         Width           =   7125
         _Version        =   196608
         _ExtentX        =   12568
         _ExtentY        =   5715
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
         SpreadDesigner  =   "frmBBS817.frx":0CF4
         VirtualRows     =   7
      End
      Begin MSComCtl2.DTPicker dtpAppDt 
         Height          =   330
         Left            =   1155
         TabIndex        =   6
         Top             =   1455
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
         Format          =   62521347
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpExpDt 
         Height          =   330
         Left            =   4455
         TabIndex        =   7
         Top             =   1440
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
         Format          =   62521347
         CurrentDate     =   36328
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   225
         X2              =   7200
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   225
         X2              =   7200
         Y1              =   1230
         Y2              =   1230
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
         Left            =   270
         TabIndex        =   10
         Tag             =   "35210"
         Top             =   1515
         Width           =   705
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
         Left            =   3615
         TabIndex        =   9
         Tag             =   "35214"
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "나이계산 단축키 :    D-입력된 값을 일령으로,  Y-연령으로,  M-최대값(364635)"
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
         Left            =   240
         TabIndex        =   8
         Tag             =   "35214"
         Top             =   5280
         Width           =   6900
      End
   End
End
Attribute VB_Name = "frmBBS817"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form명   : frmBBS810
'|  2. 기  능   : Doner Screening 검사 적격치 마스터
'|  3. 작성자   : 김 동열
'|  4. 작성일   : 2000.11.24
'|
'|  CopyRight(C) 2000 대련엠티에스
'+--------------------------------------------------------------------------------------+
Option Explicit
Private objSql                  As clsBBSMSTStatement
Attribute objSql.VB_VarHelpID = -1
'Private WithEvents mnuPopup     As Menu
'Private WithEvents mnuDelete    As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private WithEvents objListpop   As clsPopUpList
Attribute objListpop.VB_VarHelpID = -1
Private SvApplyDt               As String
Private InsertFlag              As Integer
Private UpdateFlag              As Integer
Private ClearFg                 As Boolean
Private lngRow                  As Long
Private Sub cboSpcCd_Click()
    Dim strSql As String
    Dim strSpcCd As String

    strSpcCd = medGetP(cboSpcCd.Text, 1, " ")
    Call LoadLab003AppDt(txtTestCd.Text, strSpcCd)
    If tabAppDt.Tabs.Count > 0 Then
        Call ClearRtn
        tabAppDt.Tabs(1).Selected = True
    Else
        ClearTable
        cmdNew.Enabled = True
        InsertFlag = 0
        Call cmdNew_Click
    End If
End Sub

Public Sub LoadLab003AppDt(sTestCd As String, sSpcCd As String)
    Dim RS          As Recordset
    Dim i           As Integer
    Dim strKey      As String
    Dim strCaption  As String
    
    Set objSql = New clsBBSMSTStatement
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

Public Sub LoadSpeimen(sTestCd As String)
    Dim RS  As Recordset
    Dim i   As Integer
    
    Set objSql = New clsBBSMSTStatement
    Set RS = objSql.getSpcs(sTestCd)
    cboSpcCd.Clear
    Do Until RS.EOF
        cboSpcCd.AddItem "" & RS.Fields("spccd").Value & "   " & RS.Fields("spcnm").Value & ""    ', Val(OraDS.Fields("Seq").Value) - 1
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub cboSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cboSpcCd_Click
End Sub

Private Sub cmdEdit_Click()
    Dim MyReference As clsBBSMSTStatement
    Dim strSex      As String
    Dim lngFrom     As Double
    Dim lngTo       As Double
    Dim i           As Long
    
    Set MyReference = New clsBBSMSTStatement
    If UpdateFlag = 1 Then   ' Update
        If dtpExpDt.Value <> 0 And Format(dtpExpDt, PRESENTDATE_FORMAT) < Format(GetSystemDate, PRESENTDATE_FORMAT) Then
            MsgBox "이전날짜는 사용할 수 없습니다! 폐기일을 수정하세요..", vbInformation, Me.Caption
            dtpExpDt.SetFocus
            GoTo SKIP
        ElseIf dtpExpDt.Value <> 0 And Format(dtpExpDt, PRESENTDATE_FORMAT) < Format(dtpAppDt, PRESENTDATE_FORMAT) Then
            MsgBox "적용일 이전에 폐기할 수 없습니다! 폐기일을 수정하세요..", vbInformation, Me.Caption
            dtpExpDt.SetFocus
            GoTo SKIP
        End If
        cmdEdit.Caption = "수정"
        Set MyReference = New clsBBSMSTStatement
        With tblReference
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = 1: strSex = Trim(.Value)
                .Col = 2: lngFrom = Val(.Value)
                .Col = 3: lngTo = Val(.Value)
                If chkDup(i, strSex, lngFrom, lngTo) = False Then
                    MsgBox " 중복된 검사치가 있습니다.수정하여 주세요", vbCritical, Me.Caption
                    cmdEdit.Caption = "저장"
                    GoTo SKIP
                End If
            Next
        End With
        DBConn.BeginTrans
        Call MyReference.DeleteBBS002(Trim(txtTestCd), medGetP(cboSpcCd.Text, 1, " "), Format(dtpAppDt, PRESENTDATE_FORMAT))
        With tblReference
            For i = 1 To .DataRowCnt
                Call Lab005Move(MyReference, i)
                .Row = i
                .Col = 7
                If MyReference.ApplySex = "" Then
                    MsgBox "성별을 입력하여 주세요", vbInformation
                    cmdEdit.Caption = "저장"
                    GoTo SKIP
                End If
                Call MyReference.InsertBBS002(MyReference)
            Next
            DBConn.CommitTrans
        End With
        
        Call cboSpcCd_Click
        Call LockRtn(1, True)
        dtpExpDt.Enabled = False
        cmdNew.Enabled = True
        cmdCancel.Enabled = False
    Else    ' Edit
        dtpAppDt.Enabled = False
        cmdEdit.Caption = "저장"
        UpdateFlag = 1
        Call LockRtn(2, False)
        cmdNew.Enabled = False
        cmdCancel.Enabled = True
    End If
    Set objSql = Nothing
SKIP:

End Sub

Private Sub Lab005Move(ByRef MyReference As clsBBSMSTStatement, ByVal Row As Long)
   With tblReference
         .Row = Row
         MyReference.TestCd = txtTestCd.Text
         MyReference.SpcCd = medGetP(cboSpcCd.Text, 1, " ")
         MyReference.ApplyDt = Format(dtpAppDt.Value, PRESENTDATE_FORMAT)
         If IsNull(dtpExpDt.Value) Then
            MyReference.ExpDt = ""
         Else
            MyReference.ExpDt = Format(dtpExpDt.Value, PRESENTDATE_FORMAT)
         End If
         .Col = 1:
            If .TypeComboBoxCurSel = -1 Then
                MyReference.ApplySex = ""
            Else
                MyReference.ApplySex = Choose(.TypeComboBoxCurSel + 1, "M", "F", "B", "U")
            End If
         .Col = 2: MyReference.AgeFrom = Val(.Value)
         .Col = 3: MyReference.AgeTo = Val(.Value)
         .Col = 4: MyReference.RefValFrom = Val(.Value)
         .Col = 5: MyReference.RefValTo = Val(.Value)
         .Col = 6: MyReference.RefCd = .Value
   End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set objSql = Nothing
End Sub

Private Sub cmdNew_Click()
    Dim MyReference As clsBBSMSTStatement
    Dim strSex      As String
    Dim lngFrom     As Double
    Dim lngTo       As Double
    Dim i           As Long
    
    Set MyReference = New clsBBSMSTStatement
    
    
    '새로운 데이타 추가
    If InsertFlag = 1 Then  ' Insert
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 선택하세요.", vbInformation, Me.Caption
            txtTestCd.SetFocus
            GoTo SKIP
        End If
        If Trim(medGetP(cboSpcCd.Text, 1, " ")) = "" Then
            MsgBox "검체코드를 선택하세요.", vbInformation, Me.Caption
            cboSpcCd.SetFocus
            GoTo SKIP
        End If
        If dtpAppDt.Value <> "" And Format(dtpAppDt.Value, PRESENTDATE_FORMAT) < Format(GetSystemDate, PRESENTDATE_FORMAT) Then
            MsgBox "이전날짜는 사용할 수 없습니다! 폐기일을 수정하세요..", vbInformation, Me.Caption
            dtpAppDt.SetFocus
            GoTo SKIP
        End If
        cmdNew.Caption = "추가"
        With tblReference
             For i = 1 To .DataRowCnt
                .Row = i
                .Col = 1: strSex = Trim(.Value)
                .Col = 2: lngFrom = Val(.Value)
                .Col = 3: lngTo = Val(.Value)
                If chkDup(i, strSex, lngFrom, lngTo) = False Then
                    MsgBox " 중복된 검사치가 있습니다. 수정하여 주세요", vbCritical, Me.Caption
                    cmdNew.Caption = "저장"
                    Exit Sub
                End If
            Next
            
            DBConn.BeginTrans
            For i = 1 To .DataRowCnt
            Call Lab005Move(MyReference, i)
            .Row = i
            .Col = 7
            If MyReference.ApplySex = "" Then
                MsgBox "성별을 입력하여 주세요", vbInformation
                cmdNew.Caption = "저장"
                Exit Sub
            End If
             Call MyReference.InsertBBS002(MyReference)
            Next
            DBConn.CommitTrans
        End With
        
        InsertFlag = 0
        Call cboSpcCd_Click
        Call LockRtn(1, True)
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False
        tblReference.OperationMode = OperationModeRead
        SvApplyDt = ""
    Else
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 선택하세요.", vbInformation
            txtTestCd.SetFocus
            Exit Sub
        End If
        If Trim(medGetP(cboSpcCd.Text, 1, " ")) = "" Then
            MsgBox "검체코드를 선택하세요."
            cboSpcCd.SetFocus
            Exit Sub
        End If
        cmdNew.Caption = "저장"
        dtpExpDt.Enabled = False
        InsertFlag = 1
        Call ClearTable
        Call LockRtn(1, False)
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        tblReference.OperationMode = OperationModeNormal
        If tabAppDt.Tabs.Count > 0 Then
            SvApplyDt = Format(dtpAppDt.Value, CS_DateDbFormat)
        Else
            SvApplyDt = ""
        End If
        dtpAppDt.Value = Format(Now, "YYYY-MM-DD")
        dtpAppDt.SetFocus
    End If
SKIP:
    Set MyReference = Nothing
End Sub

Private Sub cmdCancel_Click()
    InsertFlag = 0
    UpdateFlag = 0

    Call LockRtn(1, True)
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdNew.Caption = "추가"
    cmdEdit.Caption = "수정"
    cmdCancel.Enabled = False
    If tabAppDt.Tabs.Count > 0 Then tabAppDt.Tabs(1).Selected = True
End Sub

Private Sub cmdClear_Click()
    tabAppDt.Tabs.Clear
    Call ClearRtn
    cboSpcCd.Clear
    lblTestName.Caption = ""
    txtTestCd.Text = ""
    txtTestCd.SetFocus
End Sub

Private Sub ClearRtn()
    '깨끗이..
    InsertFlag = 0
    UpdateFlag = 0
    cmdNew.Caption = "추가"
    cmdNew.Enabled = True
    cmdEdit.Caption = "수정"
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False
    Call ClearTable
End Sub

Private Sub ClearTable()
    dtpExpDt.Value = ""
    With tblReference
        .Row = -1
        .Col = -1
        .Text = ""
    End With
End Sub

Private Sub LockRtn(ByVal intPart As Integer, ByVal LockValue As Boolean)
    Dim EnableValue As Boolean
    
    If LockValue Then
        EnableValue = False
        tblReference.OperationMode = OperationModeRead
    Else
        EnableValue = True
        tblReference.OperationMode = OperationModeNormal
    End If
    
    If intPart = 1 Then
        With tblReference
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 1: .Col2 = 3
            .BlockMode = True
            .Lock = False
            .BlockMode = False
        End With
        dtpAppDt.Enabled = EnableValue
    End If
    If intPart = 2 Then
        With tblReference
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 1: .Col2 = 3
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        dtpExpDt.Enabled = EnableValue
    End If
End Sub

Private Sub cmdPopupList_Click()
    Dim RS As Recordset
    
    '리스트 팝업을 불러오자...
    Set objSql = New clsBBSMSTStatement
    Set objListpop = New clsPopUpList
    objListpop.Connection = DBConn
    
'    objListpop.BackColor = Me.BackColor
    If chkTestCd.Value = 0 Then
        objListpop.Tag = "TestCd"
        objListpop.FormCaption = "검사코드 찾기"
        Call objListpop.LoadPopup(objSql.LoadPopup("0")) ', 3400, 7100)
    Else
        Set RS = objSql.GetBBS002()
        If RS.EOF = True Then
            MsgBox "등록된 검사항목이 없습니다.", vbInformation, Me.Caption
            Set RS = Nothing
            Set objListpop = Nothing
            Set objSql = Nothing
            Exit Sub
        End If
        Set RS = Nothing
        objListpop.Tag = "DTestCd"
        objListpop.FormCaption = "적격치 검사코드 찾기"
        Call objListpop.LoadPopup(objSql.LoadPopup("1")) ', 1860, 8000)
    End If
    Set objSql = Nothing
End Sub

Private Sub dtpAppDt_LostFocus()
    Dim RS As Recordset
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    If dtpAppDt.Enabled = True Then
        Set RS = objSql.ChkBBS002(Trim(txtTestCd), medGetP(cboSpcCd.Text, 1, " "), Format(dtpAppDt.Value, PRESENTDATE_FORMAT))
        If RS.EOF = True Then
            Set RS = Nothing
            Set objSql = Nothing
            Exit Sub
        Else
            Call LoadLab003(txtTestCd.Text, medGetP(cboSpcCd.Text, 1, " "), Format(dtpAppDt.Value, PRESENTDATE_FORMAT))
            cmdNew.Enabled = True
            cmdNew.Caption = "추가"
            cmdEdit.Enabled = True
            Call LockRtn(1, True)
            InsertFlag = 0
        End If
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
    
    txtTestCd.SetFocus
    dtpAppDt.Value = Format(Now, "YYYY-MM-DD")
    dtpExpDt.Enabled = False
    dtpExpDt.Value = Format(Now, "YYYY-MM-DD")
End Sub

Private Sub txtAppDt_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      With tblReference
         .SetFocus
         .Row = .DataRowCnt + 1
         .Col = 1
         .Action = ActionActiveCell
      End With
   End If
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            With tblReference
                .Row = lngRow
                .Col = -1
                .Action = 5
            End With
    End Select
End Sub

'Private Sub mnuDelete_Click()
'    With tblReference
'        .Row = lngRow
'        .Col = -1
'        .Action = 5
'    End With
'End Sub

Private Sub tblReference_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not (tblReference.Col = 2 Or tblReference.Col = 3) Then Exit Sub
    Select Case KeyCode
        Case vbKeyY:   '연령으로
            tblReference.Value = tblReference.Value / 365
        Case vbKeyD:   '일령으로
            tblReference.Value = tblReference.Value * 365
        Case vbKeyM:   'Maximun
            tblReference.Value = 364635
    End Select
End Sub

Private Sub objListpop_SendCode(ByVal SelString As String)
    Dim RS  As Recordset
    Dim i   As Integer
    
    '리스트박스에 있는내용을 가져오자..
    Set objSql = New clsBBSMSTStatement
    txtTestCd.Text = medGetP(SelString, 1, ";")
    lblTestName.Caption = medGetP(SelString, 2, ";")
    If txtTestCd.Text = "" Then Exit Sub
        lblTestName.Caption = ""
        ClearRtn
        tabAppDt.Tabs.Clear
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
    Set objMyList = Nothing
    Set objListpop = Nothing
End Sub

Private Sub tblReference_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    tblReference.Row = NewRow
    tblReference.Col = NewCol
    If Col = 6 Then
        tblReference.RowHeight(Row) = tblReference.MaxTextRowHeight(Row)
    End If
End Sub

Private Sub txtTestCd_Change()
   If Not ClearFg Then
      Call ClearRtn
      ClearFg = True
   End If
End Sub
Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    Dim RS As Recordset
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If txtTestCd.Text = "" Then Exit Sub
        lblTestName.Caption = ""
        ClearRtn
        tabAppDt.Tabs.Clear
        Set objSql = New clsBBSMSTStatement
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
    End If
End Sub
Private Sub tblReference_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
   
    If tblReference.OperationMode = OperationModeRead Then Exit Sub
    lngRow = Row
    tblReference.Row = lngRow
    tblReference.Col = -1
    tblReference.BackColor = &HC0C0C0
    lngRow = Row
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        .PopupMenus Me.hWnd
    End With
    
    Set objPop = Nothing
'    Set mnuPopup = frmControl.mnuPopup
'    Set mnuDelete = frmControl.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
    
    If lngRow Mod 2 = 0 Then
        tblReference.BackColor = 14411494
    Else
        tblReference.BackColor = -2147483643
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
   cmdNew.Enabled = True
   cmdNew.Caption = "추가"
   cmdEdit.Enabled = True
   cmdEdit.Caption = "수정"
   Call LockRtn(1, True)
End Sub

Public Sub LoadLab003(sTestCd As String, sSpcCd As String, sAppDt As String)
    Dim i As Integer
    Dim MyReference As New clsBBSMSTStatement
    Dim RS As Recordset
    Dim sgTmp As Single
    
    Set MyReference = New clsBBSMSTStatement
    Set RS = MyReference.getReference(sTestCd, sSpcCd, sAppDt)
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
    Set RS = Nothing
    Set MyReference = Nothing
End Sub

Public Sub Raise_TestCd_Keypress()
   Call txtTestCd_KeyPress(13)
End Sub

Public Sub Raise_cboSpcCd_Click()
   Call cboSpcCd_Click
End Sub

Private Sub txtTestCd_LostFocus()
    Dim RS As Recordset
    
    If txtTestCd.Text = "" Then Exit Sub
    lblTestName.Caption = ""
    ClearRtn
    tabAppDt.Tabs.Clear
    Set objSql = New clsBBSMSTStatement
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
    Set RS = Nothing
End Sub

Private Function chkDup(ByVal Prow As Long, ByVal psex As String, ByVal pfrom As Double, ByVal pto As Double) As Boolean
    Dim i As Long
    Dim strSex As String
    Dim lngFrom As Double
    Dim lngTo As Double
    
    With tblReference
        For i = Prow + 1 To .DataRowCnt
            .Row = i
            .Col = 1: strSex = Trim(.Value)
            .Col = 2: lngFrom = Val(.Value)
            .Col = 3: lngTo = Val(.Value)
            
            If psex = strSex And pfrom = lngFrom And pto = lngTo Then chkDup = False: Exit Function
        Next
    End With
    chkDup = True
End Function




