VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm353Reference 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "Lis353.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "128"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "128"
      Top             =   8190
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   6495
      Left            =   30
      TabIndex        =   7
      Top             =   1545
      Width           =   10830
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   7095
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   25
         Text            =   "Lis353.frx":038A
         Top             =   2880
         Visible         =   0   'False
         Width           =   3630
      End
      Begin MSComctlLib.TabStrip tabRef 
         Height          =   360
         Left            =   120
         TabIndex        =   24
         Top             =   1695
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   635
         MultiRow        =   -1  'True
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "General Reference"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Panic/Critical Reference"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "AMR Reference"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   405
         Left            =   210
         TabIndex        =   22
         Top             =   930
         Width           =   10560
         Begin MSComctlLib.TabStrip tabAppDt 
            Height          =   300
            Left            =   60
            TabIndex        =   23
            Top             =   75
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   529
            Style           =   2
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H0085A3A3&
            BorderWidth     =   2
            Height          =   360
            Left            =   45
            Top             =   45
            Width           =   10485
         End
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0CFC2&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   5445
         Style           =   1  '그래픽
         TabIndex        =   21
         Tag             =   "35301"
         Top             =   195
         Width           =   1320
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소(&U)"
         Height          =   510
         Left            =   9405
         Style           =   1  '그래픽
         TabIndex        =   11
         Tag             =   "35301"
         Top             =   195
         Width           =   1320
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "수정(&E)"
         Height          =   510
         Left            =   8085
         Style           =   1  '그래픽
         TabIndex        =   9
         Tag             =   "135"
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가(&A)"
         Height          =   510
         Left            =   6765
         Style           =   1  '그래픽
         TabIndex        =   8
         Tag             =   "35301"
         Top             =   195
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker txtAppDt 
         Height          =   300
         Left            =   8970
         TabIndex        =   3
         Top             =   1395
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
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
         Format          =   87556099
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker txtExpDt 
         Height          =   300
         Left            =   8970
         TabIndex        =   4
         Top             =   1725
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd"
         DateIsNull      =   -1  'True
         Format          =   87556099
         CurrentDate     =   36328
      End
      Begin FPSpread.vaSpread tblReference 
         Height          =   3255
         Left            =   135
         TabIndex        =   18
         Tag             =   "35304"
         Top             =   2070
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   5741
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   50
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis353.frx":03BE
         VirtualRows     =   7
         VisibleCols     =   8
         VisibleRows     =   9
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "나이계산 단축키 :  D-입력된 값을 일령으로,  Y-연령으로,  M-최대값(50000)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   15
         Tag             =   "35214"
         Top             =   5415
         Width           =   6945
      End
      Begin VB.Label lblExpDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폐 기 일"
         Height          =   180
         Left            =   8175
         TabIndex        =   13
         Tag             =   "35214"
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label lblAppDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "적 용 일"
         Height          =   180
         Left            =   8175
         TabIndex        =   12
         Tag             =   "35210"
         Top             =   1455
         Width           =   660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   225
         X2              =   8500
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   285
         X2              =   10725
         Y1              =   1350
         Y2              =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1275
      Left            =   30
      TabIndex        =   0
      Top             =   225
      Width           =   10905
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00EBEBEB&
         Caption         =   "(&N) >>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   7635
         Style           =   1  '그래픽
         TabIndex        =   20
         Tag             =   "124"
         Top             =   525
         Width           =   1320
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00EBEBEB&
         Caption         =   "<< (&P)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   6315
         Style           =   1  '그래픽
         TabIndex        =   19
         Tag             =   "124"
         Top             =   525
         Width           =   1320
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
         Left            =   1305
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis353.frx":0C43
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   360
         Width           =   300
      End
      Begin VB.ComboBox cboSpcCd 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   720
         Width           =   4680
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1620
         TabIndex        =   1
         Top             =   345
         Width           =   1395
      End
      Begin MedControls1.LisLabel lblTestName 
         Height          =   330
         Left            =   3015
         TabIndex        =   17
         Top             =   345
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   582
         BackColor       =   14411494
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
         TabIndex        =   6
         Tag             =   "35303"
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label1 
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
         TabIndex        =   5
         Tag             =   "35302"
         Top             =   405
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm353Reference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1

Private WithEvents objHelpList As clsPopUpList
Attribute objHelpList.VB_VarHelpID = -1
Private MyItem As New clsItem
Private MySqlStmt As New clsLISSqlStatement    ' SQL 클래스

Private InsertFlag As Integer
Private UpdateFlag As Integer
Private SvApplyDt As String
Private ClearFg As Boolean


'% 검체선택하면 기준치정보 Display

Private Sub cboSpcCd_Click()

    Dim tmpSql As String
    Dim tmpSpcCd As String

    tmpSpcCd = medGetP(cboSpcCd.Text, 1, " ")
    tmpSql = MySqlStmt.SqlLAB005AppDt(txtTestCd.Text, tmpSpcCd)
    Call Lab005AppDt(tmpSql)
    If tabAppDt.Tabs.Count > 0 Then
    'If tabAppDt.Pages.Count > 0 Then
        Call ClearRtn
        'tabAppDt.Value = 0: Call tabAppDt_Click(0)
        tabAppDt.Tabs(1).Selected = True
    Else
        cmdNew.Enabled = True
        InsertFlag = 0
        Call cmdNew_Click
    End If

End Sub

Private Sub cboSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then Call cboSpcCd_Click

End Sub


Private Sub cmdCancel_Click()

    Call CancelRoutine
    Call cboSpcCd_Click
    If tabAppDt.Tabs.Count > 0 Then tabAppDt.Tabs(1).Selected = True
    'If tabAppDt.Pages.Count > 0 Then tabAppDt.Value = 0: Call tabAppDt_Click(0) ' tabAppDt.Tabs(1).Selected = True

End Sub

Private Sub CancelRoutine()

    If Not ConfirmExit Then Exit Sub
    
    InsertFlag = 0
    UpdateFlag = 0

    Call LockRtn(2, True)

    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdNew.Caption = "추가"
    cmdEdit.Caption = "수정"
    cmdCancel.Enabled = False

End Sub

Private Sub cmdClear_Click()

    If Not ConfirmExit Then Exit Sub

    Call ClearRtn
    txtTestCd.Text = ""
    txtTestCd.SetFocus
End Sub

Private Sub cmdDelete_Click()
    
    Dim Resp As VbMsgBoxResult
    
    Resp = MsgBox("해당 적용일의 데이타를 모두 삭제하시겠습니까?", vbQuestion, "참고치 등록")
    If Resp = vbNo Then Exit Sub
    
    Dim MyReference As New clsReference

    Call Lab005Move(MyReference, 1)
    Call MyReference.RefDelete(True)
    
    Call txtTestCd_KeyPress(vbKeyReturn)
    
End Sub

Private Sub cmdEdit_Click()

    Dim MyReference As New clsReference
    Dim i As Long
    Dim Resp As VbMsgBoxResult
    Dim intYesNo As VbMsgBoxResult
    
    intYesNo = MsgBox("자료가 수정되었습니다." & vbNewLine & "수정된 자료를 저장하시겠습니까?", vbYesNo, "참고치등록")
    If intYesNo = vbNo Then Exit Sub

'    Resp = MsgBox("해당 적용일의 데이타를 수정하시겠습니까?", vbQuestion, "참고치 수정")
'    If Resp = vbNo Then Exit Sub

    If UpdateFlag = 1 Then  ' Update

        cmdEdit.Caption = "수정"
        With tblReference
            For i = 1 To .DataRowCnt
                Call Lab005Move(MyReference, i)
                .Row = i
                .Col = 7
                Select Case .Value
                    Case "":  MyReference.RefInsert
                    Case "1": MyReference.RefUpdate
                    Case "2": MyReference.RefDelete
                End Select
                'MyReferences.Add Format(txtSpcAppDt.Text, CS_DateDbFormat), MyReference
            Next
        End With
        Set MyReference = Nothing
        UpdateFlag = 0
        Call cboSpcCd_Click
        Call LockRtn(1, True)
        cmdNew.Enabled = True
        cmdCancel.Enabled = False

    Else    ' Edit

        txtAppDt.Enabled = False
        cmdEdit.Caption = "저장"
        UpdateFlag = 1
        Call LockRtn(2, False)
        cmdNew.Enabled = False
        cmdCancel.Enabled = True
        'Call txtAppDt_KeyPress(13)

    End If

End Sub

Private Sub cmdExit_Click()

    If Not ConfirmExit Then Exit Sub

    Unload Me
'   Set frm353Reference = Nothing
End Sub

Private Sub cmdFind_Click(Index As Integer)
    
    Dim i As Integer
    
    If txtTestCd.Text = "" Then Exit Sub
    If Not ConfirmExit Then Exit Sub

'    I = medListFind(lstItemList, txtTestCd.Text)
    If Not lstItemList.Exists(txtTestCd.Text) Then Exit Sub
    Call lstItemList.KeyChange(txtTestCd.Text)

'    If I < 0 Then Exit Sub
    Select Case Index
        Case 0:   'Previous
            'If I <= 0 Then Exit Sub
            'txtTestCd.Text = lstItemList.List(I - 1)
            lstItemList.MovePrevious
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
        Case 1:   'Next
'            If I >= lstItemList.ListCount - 1 Then Exit Sub
'            txtTestCd.Text = lstItemList.List(I + 1)
            lstItemList.MoveNext
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
    End Select
    Call txtTestCd_KeyPress(vbKeyReturn)

End Sub

Private Sub cmdNew_Click()

    Dim i As Integer
    Dim MyReference As New clsReference
    Dim Resp As VbMsgBoxResult
    Dim intYesNo As VbMsgBoxResult
    
    intYesNo = MsgBox("자료가 수정되었습니다." & vbNewLine & "수정된 자료를 저장하시겠습니까?", vbYesNo, "참고치등록")
    If intYesNo = vbNo Then Exit Sub
    
'    Resp = MsgBox("해당 데이타를 모두 저장하시겠습니까?", vbQuestion, "참고치 등록")
'    If Resp = vbNo Then Exit Sub


    If InsertFlag = 1 Then  ' Insert
        If SvApplyDt <> "" And SvApplyDt >= Format(txtAppDt.Value, CS_DateDbFormat) Then
            MsgBox "적용일을 수정하세요.."
            txtAppDt.SetFocus
            Exit Sub
        End If
        cmdNew.Caption = "추가"

        With tblReference
            For i = 1 To .DataRowCnt
                Call Lab005Move(MyReference, i)
                .Row = i
                .Col = 7
                Select Case .Value
                Case "":
                   MyReference.RefInsert
                End Select
                'MyReferences.Add Format(txtSpcAppDt.Text, CS_DateDbFormat), MyReference
            Next
        End With
        Set MyReference = Nothing
        InsertFlag = 0
        Call cboSpcCd_Click
        Call LockRtn(1, True)
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False
        tblReference.OperationMode = OperationModeRead
        SvApplyDt = ""

    Else    ' New

        cmdNew.Caption = "저장"
        InsertFlag = 1
        Call ClearTable
        Call LockRtn(1, False)
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        tblReference.OperationMode = OperationModeNormal
        'If tabAppDt.Pages.Count > 0 Then
        If tabAppDt.Tabs.Count > 0 Then
            SvApplyDt = Format(txtAppDt.Value, CS_DateDbFormat)
        Else
            SvApplyDt = ""
        End If
        txtAppDt.Value = Format(Now, CS_DateLongFormat)
        
        On Error Resume Next
        
        txtAppDt.SetFocus

    End If

End Sub

Private Sub cmdPopupList_Click()

    Dim tmpSql As String

    If Not ConfirmExit Then Exit Sub

    Set objHelpList = New clsPopUpList
    With objHelpList
        .Connection = DBConn
        .FormCaption = "Test Code List.."
        .Tag = "TestCode"
        .ColumnHeaderText = "검사코드;검사명"
        tmpSql = MySqlStmt.SqlLAB001CodeList
        .LoadPopUp tmpSql '(, Me.Top + txtTestCd.Top + txtTestCd.Height, Me.Left + txtTestCd.Left, lstItemList)
        'Call .ListPop(tmpSql, Me.Top + txtTestCd.Top + txtTestCd.Height, Me.Left + txtTestCd.Left)
        txtTestCd.Text = Trim(medShift(.SelectedString, ";"))
        Call txtTestCd_KeyPress(vbKeyReturn)
    End With

End Sub


Private Sub Form_Deactivate()
    Set objHelpList = Nothing

End Sub

Private Sub Form_Load()

    Call medAlwaysOn(frm353Reference, 1)

    txtAppDt.Value = Format(Now, CS_DateLongFormat)
    txtExpDt.Value = ""
    lblTestName.Caption = ""

    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdCancel.Enabled = False
    ClearFg = True

    InsertFlag = 0
    UpdateFlag = 0
    
    tabAppDt.Tabs.Clear
    'tabAppDt.Pages.Clear
    
    Call MyItem.GetItemList(lstItemList): DoEvents
    cmdDelete.Enabled = ObjMyUser.isdeveloper 'gIsDeveloper
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
    
    Set objHelpList = Nothing
    Set MyItem = Nothing
    Set MySqlStmt = Nothing

End Sub

Private Sub objHelpList_SelectedItem(ByVal pSelectedItem As String)
    Set objHelpList = Nothing
    txtTestCd.Text = medShift(pSelectedItem, ";")
    lblTestName.Caption = medShift(pSelectedItem, ";")
    Call txtTestCd_KeyPress(vbKeyReturn)
    Me.Enabled = True
End Sub

'Private Sub mnuDelete_Click()
'    tblReference.Col = 7
'    If tblReference.Value = "1" Then
'       tblReference.Value = "2"
'    ElseIf tblReference.Value = "" Then
'       tblReference.Value = "3"
'    End If
'
'    tblReference.RowHidden = True
'
'End Sub

'Private Sub objHelpList_SendCode(ByVal SelString As String)
'
'    Set objHelpList = Nothing
'    txtTestCd.Text = medShift(SelString, ";")
'    lblTestName.Caption = medShift(SelString, ";")
'    Call txtTestCd_KeyPress(vbKeyReturn)
'    Me.Enabled = True
'
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            tblReference.Col = 7
            If tblReference.Value = "1" Then
               tblReference.Value = "2"
            ElseIf tblReference.Value = "" Then
               tblReference.Value = "3"
            End If
            
            tblReference.RowHidden = True
    End Select
End Sub

'Private Sub tabAppDt_Click(Index As Long)
Private Sub tabAppDt_Click()

    Dim tmpSql As String
    Dim tmpAppDt As String
    Dim tmpSpcCd As String

    'lstSpcName.ListIndex = cboSpcCd.ListIndex
    
    Call CancelRoutine

    tabRef.Tabs(1).Selected = True

'    tmpAppDt = Format(tabAppDt.SelectedItem.Caption, CS_DateDbFormat)
'    tmpSpcCd = medGetP(cboSpcCd.Text, 1, " ")
'
'    tmpSql = MySqlStmt.SqlLAB005Read(txtTestCd.Text, tmpSpcCd, tmpAppDt)
'
'    Call Lab005Load(tmpSql)


    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    Call LockRtn(2, True)

End Sub



Private Sub tabRef_Click()
    Dim RS          As Recordset
    Dim tmpStr      As String
    Dim tmpSex      As String
    Dim tmpSql      As String
    Dim tmpAppDt    As String
    Dim tmpSpcCd    As String
    Dim i           As Integer
    
'On Error GoTo Error_Trap
    
    If tabRef.SelectedItem.Index = 1 Then
        tmpStr = "성별" & vbTab & "일령(From)" & vbTab & "일령(To)" & vbTab & "From Value" & vbTab & _
                 "To Value" & vbTab & "Alpha Value" & vbTab & "Flag" & vbTab & "다중참고치" & vbTab & _
                 "Auto From" & vbTab & "Auto To"
        tmpSex = "남자" & vbTab & "여자" & vbTab & "Both" & vbTab & "중성"
        With tblReference
            .Row = 0: .Row2 = 0: .Col = 1: .Col2 = 10
            .BlockMode = True: .Clip = tmpStr: .BlockMode = False
            .Row = 1: .Row2 = 50: .Col = 1: .Col2 = 1: .BlockMode = True:
            .CellType = CellTypeComboBox
            .TypeComboBoxList = tmpSex
            .BlockMode = False
        End With
    ElseIf tabRef.SelectedItem.Index = 2 Then
        tmpStr = "성별" & vbTab & "일령(From)" & vbTab & "일령(To)" & vbTab & "Panic From" & vbTab & _
                 "Panic To" & vbTab & "Alpha Value" & vbTab & "Flag" & vbTab & "다중참고치" & vbTab & _
                 "Critical From" & vbTab & "Critical To"
        tmpSex = "P남" & vbTab & "P여" & vbTab & "P_Both"
        With tblReference
            .Row = 0: .Row2 = 0: .Col = 1: .Col2 = 10
            .BlockMode = True: .Clip = tmpStr: .BlockMode = False
            .Row = 1: .Row2 = 50: .Col = 1: .Col2 = 1: .BlockMode = True:
            .CellType = CellTypeComboBox
            .TypeComboBoxList = tmpSex
            .BlockMode = False
        End With
    Else
        tmpStr = "성별" & vbTab & "일령(From)" & vbTab & "일령(To)" & vbTab & "AMR From" & vbTab & _
                 "AMR To" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                 "" & vbTab & ""
        tmpSex = "M남" & vbTab & "M여" & vbTab & "M_Both"
        With tblReference
            .Row = 0: .Row2 = 0: .Col = 1: .Col2 = 10
            .BlockMode = True: .Clip = tmpStr: .BlockMode = False
            .Row = 1: .Row2 = 50: .Col = 1: .Col2 = 1: .BlockMode = True:
            .CellType = CellTypeComboBox
            .TypeComboBoxList = tmpSex
            .BlockMode = False
        End With
    End If
    
    If tabAppDt.Tabs.Count < 1 Then Exit Sub
    Call medClearTable(tblReference, False, False)
    tmpAppDt = Format(tabAppDt.SelectedItem.Caption, CS_DateDbFormat)
    tmpSpcCd = medGetP(cboSpcCd.Text, 1, " ")
    tmpSql = MySqlStmt.SqlLAB005Read(txtTestCd.Text, tmpSpcCd, tmpAppDt)
    
    Set RS = New Recordset
    RS.Open tmpSql, DBConn
    
    If RS.EOF Then GoTo NoData
    i = 0
    
    If tabRef.SelectedItem.Index = 1 Then
        txtAppDt.Value = Format(CStr(RS.Fields("ApplyDt").Value), CS_DateMask)
        With tblReference
            .MaxRows = 50
            .Row = 0
            While (RS.EOF = False)
                If RS.Fields("applysex").Value & "" = "M" Or RS.Fields("applysex").Value & "" = "F" Or _
                   RS.Fields("applysex").Value & "" = "B" Or RS.Fields("applysex").Value & "" = "U" Then
                    i = i + 1
                    If .Row = .MaxRows Then .MaxRows = .MaxRows + i
                    .Row = .Row + 1: .RowHidden = False
                    .TypeHAlign = TypeHAlignCenter
                    .Col = 1:
                    Select Case RS.Fields("ApplySex").Value & ""
                       Case "M": .TypeComboBoxCurSel = 0
                       Case "F": .TypeComboBoxCurSel = 1
                       Case "B": .TypeComboBoxCurSel = 2
                       Case "U": .TypeComboBoxCurSel = 3
                    End Select
                    
                    .Col = 2: .Value = RS.Fields("AgeFrom").Value & ""
                    .Col = 3: .Value = RS.Fields("AgeTo").Value & ""
                    .Col = 4: .Value = RS.Fields("RefValFrom").Value & ""
                    .Col = 5: .Value = RS.Fields("RefValTo").Value & ""
                    .Col = 6: .Value = RS.Fields("RefCd").Value & ""
                    .Col = 7: .Value = "1"
                    '.Col=8:.RS.Fields("RefText").Value & ""
                    .Col = 9: .Value = RS.Fields("arefvalfrom").Value & ""
                    .Col = 10: .Value = RS.Fields("arefvalto").Value & ""
                    
                    If RS.Fields("RefText").Value & "" <> "" Then
                        .Col = 8: .Value = "Y": .ForeColor = DCM_LightRed: .FontBold = True
                                                .TypeHAlign = TypeHAlignCenter
                        .Col = 11: .Value = RS.Fields("RefText").Value & ""
                    End If
                    
                    txtAppDt.Value = Format(RS.Fields("ApplyDt").Value & "", CS_DateMask)
                    If Trim(RS.Fields("ExpDt").Value) = "" Then
                        txtExpDt.Value = ""
                    Else
                        txtExpDt.Value = Format(RS.Fields("ExpDt").Value & "", CS_DateMask)
                    End If
                End If
                RS.MoveNext
            Wend
            .RowHeight(-1) = 13.3
        End With
    ElseIf tabRef.SelectedItem.Index = 2 Then
        txtAppDt.Value = Format(CStr(RS.Fields("ApplyDt").Value), CS_DateMask)
        With tblReference
            .MaxRows = 50
            .Row = 0
            While (RS.EOF = False)
                If RS.Fields("applysex").Value & "" = "X" Or RS.Fields("applysex").Value & "" = "Y" Or _
                                                             RS.Fields("applysex").Value & "" = "Z" Then
                    i = i + 1
                    If .Row = .MaxRows Then .MaxRows = .MaxRows + i
                    .Row = .Row + 1: .RowHidden = False
                    .TypeHAlign = TypeHAlignCenter
                    .Col = 1:
                    Select Case RS.Fields("ApplySex").Value & ""
                       Case "X": .TypeComboBoxCurSel = 0
                       Case "Y": .TypeComboBoxCurSel = 1
                       Case "Z": .TypeComboBoxCurSel = 2
                    End Select
                    
                    .Col = 2: .Value = RS.Fields("AgeFrom").Value & ""
                    .Col = 3: .Value = RS.Fields("AgeTo").Value & ""
                    .Col = 4: .Value = RS.Fields("panicfrval").Value & ""
                    .Col = 5: .Value = RS.Fields("panictoval").Value & ""
                    .Col = 6: .Value = RS.Fields("RefCd").Value & ""
                    .Col = 7: .Value = "1"
                    '.Col=8:.RS.Fields("RefText").Value & ""
                    .Col = 9: .Value = RS.Fields("arletfrval").Value & ""
                    .Col = 10: .Value = RS.Fields("arlettoval").Value & ""
                        
                    If RS.Fields("RefText").Value & "" <> "" Then
                        .Col = 8: .Value = "Y": .ForeColor = DCM_LightRed: .FontBold = True
                                                .TypeHAlign = TypeHAlignCenter
                        .Col = 11: .Value = RS.Fields("RefText").Value & ""
                    End If
                    
                    txtAppDt.Value = Format(RS.Fields("ApplyDt").Value & "", CS_DateMask)
                    If Trim(RS.Fields("ExpDt").Value) = "" Then
                        txtExpDt.Value = ""
                    Else
                        txtExpDt.Value = Format(RS.Fields("ExpDt").Value & "", CS_DateMask)
                    End If
                End If
                RS.MoveNext
            Wend
            .RowHeight(-1) = 13.3
        End With
    Else
        txtAppDt.Value = Format(CStr(RS.Fields("ApplyDt").Value), CS_DateMask)
        With tblReference
            .MaxRows = 50
            .Row = 0
            While (RS.EOF = False)
                If RS.Fields("applysex").Value & "" = "X" Or RS.Fields("applysex").Value & "" = "Y" Or _
                                                             RS.Fields("applysex").Value & "" = "Z" Then
                    i = i + 1
                    If .Row = .MaxRows Then .MaxRows = .MaxRows + i
                    .Row = .Row + 1: .RowHidden = False
                    .TypeHAlign = TypeHAlignCenter
                    .Col = 1:
                    Select Case RS.Fields("ApplySex").Value & ""
                       Case "X": .TypeComboBoxCurSel = 0
                       Case "Y": .TypeComboBoxCurSel = 1
                       Case "Z": .TypeComboBoxCurSel = 2
                    End Select
                    
                    .Col = 2: .Value = RS.Fields("AgeFrom").Value & ""
                    .Col = 3: .Value = RS.Fields("AgeTo").Value & ""
                    .Col = 4: .Value = RS.Fields("AMRfrval").Value & ""
                    .Col = 5: .Value = RS.Fields("AMRtoval").Value & ""
'                    .Col = 6: .Value = RS.Fields("RefCd").Value & ""
                    .Col = 7: .Value = "1"
                    '.Col=8:.RS.Fields("RefText").Value & ""
'                    .Col = 9: .Value = RS.Fields("arletfrval").Value & ""
'                    .Col = 10: .Value = RS.Fields("arlettoval").Value & ""
'
'                    If RS.Fields("RefText").Value & "" <> "" Then
'                        .Col = 8: .Value = "Y": .ForeColor = DCM_LightRed: .FontBold = True
'                                                .TypeHAlign = TypeHAlignCenter
'                        .Col = 11: .Value = RS.Fields("RefText").Value & ""
'                    End If
                    
                    txtAppDt.Value = Format(RS.Fields("ApplyDt").Value & "", CS_DateMask)
                    If Trim(RS.Fields("ExpDt").Value) = "" Then
                        txtExpDt.Value = ""
                    Else
                        txtExpDt.Value = Format(RS.Fields("ExpDt").Value & "", CS_DateMask)
                    End If
                End If
                RS.MoveNext
            Wend
            .RowHeight(-1) = 13.3
        End With
    End If
NoData:
    Set RS = Nothing

End Sub

Private Sub tblReference_Click(ByVal Col As Long, ByVal Row As Long)
    Dim Wdt As Long
    Dim Hgt As Long
    Dim X   As Long
    Dim Y   As Long
    Dim Ret As Boolean
    

    With tblReference
        .Row = Row
        Select Case Col
            Case 8
                Ret = .GetCellPos(8, Row, X, Y, Wdt, Hgt)
                If Row <> .DataRowCnt Then
                     Y = Y + Hgt
                Else
                     Y = Y
                End If
                
                If .Height - Y < txtComment.Height Or Y < 0 Then
                       Ret = .GetCellPos(8, Row, X, Y, Wdt, Hgt)
                       txtComment.Top = .Top + Y - txtComment.Height + MainFrm.picMain.Height + 950
                       txtComment.Left = .Left + X
                
                Else
                   txtComment.Left = .Left + X
                   txtComment.Top = .Top + Y
                End If
                .Col = 11
                txtComment.Text = .Value
                txtComment.Tag = Row
                txtComment.Visible = True
                If txtComment.Enabled Then txtComment.SetFocus
                If cmdEdit.Caption = "수정" Then
                     txtComment.Enabled = False
                Else
                    txtComment.Enabled = True
                End If
        Case Else
            txtComment.Visible = False
        End Select
    End With
End Sub
Private Sub txtComment_KeyDown(KeyCode As Integer, Shift As Integer)
    If Val(txtComment.Tag) < 1 Or Val(txtComment.Tag) > tblReference.MaxRows Then Exit Sub
    If KeyCode = vbKeyReturn Then
        With tblReference
            .Row = txtComment.Tag
            .Col = 11: .Value = txtComment.Text
            If .Value <> "" Then
                .Col = 8: .Value = "Y": .ForeColor = DCM_LightRed: .FontBold = True
                                        .TypeHAlign = TypeHAlignCenter
            Else
                .Col = 8: .Value = ""
            End If
        End With
        txtComment.Visible = False
        Call tblReference.SetFocus
    End If
    
End Sub
Private Sub tblReference_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Not (tblReference.Col = 2 Or tblReference.Col = 3) Then Exit Sub
    Select Case KeyCode
        Case vbKeyY:   '연령으로
            tblReference.Value = tblReference.Value / 365
        Case vbKeyD:   '일령으로
            tblReference.Value = tblReference.Value * 365
        Case vbKeyM:   'Maximun
            tblReference.Value = 50000
    End Select

End Sub

Private Sub tblReference_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If tblReference.OperationMode = OperationModeRead Then Exit Sub
    
    Dim lngOldColor As Long

    tblReference.Row = Row
    tblReference.Col = -1
    lngOldColor = tblReference.BackColor
    tblReference.BackColor = &HC0C0C0
    
    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        
        .PopupMenus Me.hWnd
    End With
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup

    tblReference.Row = Row
    tblReference.Col = -1
    tblReference.BackColor = lngOldColor
    
    Set objPop = Nothing
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing

End Sub

Private Sub tblReference_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    tblReference.Row = NewRow
    tblReference.Col = NewCol
    If Col = 8 Then
'        tblReference.RowHeight(Row) = tblReference.MaxTextRowHeight(Row)
    End If
End Sub


Private Sub txtAppDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With tblReference
            '.MaxRows = .DataRowCnt + 1
            .SetFocus
            .Row = .DataRowCnt + 1
            .Col = 1
            .Action = ActionActiveCell
            'Call tblReference_Click(1, .Row)
        End With
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

Private Sub txtTestCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objHelpList Is Nothing Then Call cmdPopupList_Click
'        Call clsCodeList.SetFocus(2)
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    Dim tmpSql As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If Not ConfirmExit Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = vbKeyReturn Then
        If txtTestCd.Text = "" Then Exit Sub
    
        Call ClearRtn
        'tabAppDt.Pages.Clear
        tabAppDt.Tabs.Clear
        ' 선린병원:2001-05-31
        If lstItemList.Exists(Trim(txtTestCd.Text)) Then
            lstItemList.KeyChange (Trim(txtTestCd.Text))
            lblTestName.Caption = lstItemList.Fields("testnm")
'        If ObjLISComCode.LisItem.Exists(Trim(txtTestCd.Text)) Then
'            lblTestName.Caption = ObjLISComCode.LisItem.Fields("testnm")
        Else
            lblTestName.Caption = ""
        End If
    
        tmpSql = MySqlStmt.SqlSpecimenRead(txtTestCd.Text)
        Call LabSpecimenLoad(tmpSql)
        If cboSpcCd.ListCount > 0 Then
            cboSpcCd.ListIndex = 0
            cboSpcCd.SetFocus
        End If
        ClearFg = False
    End If
End Sub


'% Sub Routine 3 : LabSpecimenLoad
'%                        지정검체명들을 Tab에 Display

Private Sub LabSpecimenLoad(ByVal SqlStmt As String)

    Dim objRs As Recordset
    Dim i As Integer

    Set objRs = New Recordset   'Sql 실행
    objRs.Open SqlStmt, DBConn

    cboSpcCd.Clear

    While (objRs.EOF = False)
        cboSpcCd.AddItem "" & objRs.Fields("SpcCd").Value & "   " & objRs.Fields("SpcNm").Value  ', Val(objRs.Fields("Seq").Value) - 1
        objRs.MoveNext
    Wend

    Set objRs = Nothing

End Sub


'% Sub Routine 3 : LabSpecimenLoad
'%                        지정검체명들을 Tab에 Display

Private Sub Lab005AppDt(ByVal SqlStmt As String)

    Dim objRs As Recordset       'Oracle DynaSet
    Dim i As Integer
    Dim tmpKey As String
    Dim tmpCaption As String

    Set objRs = New Recordset  'Sql 실행
    objRs.Open SqlStmt, DBConn

    i = 0
    'tabAppDt.Pages.Clear
    tabAppDt.Tabs.Clear
    While (objRs.EOF = False)
        i = i + 1
        tmpKey = objRs.Fields("ApplyDt").Value
        tmpCaption = Format(tmpKey, CS_DateMask)
        tabAppDt.Tabs.Add i, , tmpCaption
        'tabAppDt.Pages.Add tmpKey, tmpCaption, i - 1
        objRs.MoveNext
    Wend

    Set objRs = Nothing
End Sub


'% Sub Routine 4 : Lab005Load
'%                        Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                        클래스 clsReference의 Data Attribute에 저장한다.

Private Sub Lab005Load(ByVal SqlStmt As String)

    Dim i As Integer
    Dim objRs As Recordset

    On Error GoTo Error_Trap


    Call medClearTable(tblReference, False, False)

    Set objRs = New Recordset  'Sql 실행
    objRs.Open SqlStmt, DBConn
    
    If objRs.EOF Then GoTo NoData

    i = 0
    txtAppDt.Value = Format(CStr(objRs.Fields("ApplyDt").Value), CS_DateMask)
    With tblReference
        .MaxRows = 50
        .Row = 0
        While (objRs.EOF = False)
            i = i + 1
            If .Row = .MaxRows Then .MaxRows = .MaxRows + i
            .Row = .Row + 1
            .RowHidden = False
            .TypeHAlign = TypeHAlignCenter
            .Col = 1:
                     Select Case "" & objRs.Fields("ApplySex").Value
                     Case "M":
                        .TypeComboBoxCurSel = 0
                     Case "F":
                        .TypeComboBoxCurSel = 1
                     Case "B":
                        .TypeComboBoxCurSel = 2
                     Case "U":
                        .TypeComboBoxCurSel = 3
                     End Select
            .Col = 2: .Value = "" & objRs.Fields("AgeFrom").Value
            .Col = 3: .Value = "" & objRs.Fields("AgeTo").Value
            .Col = 4: .Value = "" & objRs.Fields("RefValFrom").Value
            .Col = 5: .Value = "" & objRs.Fields("RefValTo").Value
            .Col = 6: .Value = "" & objRs.Fields("RefCd").Value
            .Col = 7: .Value = "1"
            
            '.Col = 8: .Value = "" & objRs.Fields("RefText").Value
            .Col = 9: .Value = "" & objRs.Fields("panicfrval").Value
            .Col = 10: .Value = "" & objRs.Fields("panictoval").Value
            .Col = 11: .Value = "" & objRs.Fields("arefvalfrom").Value
            .Col = 12: .Value = "" & objRs.Fields("arefvalto").Value
            
            txtAppDt.Value = Format(objRs.Fields("ApplyDt").Value, CS_DateMask)
            If Trim(objRs.Fields("ExpDt").Value) = "" Then
                txtExpDt.Value = ""
            Else
                txtExpDt.Value = Format(objRs.Fields("ExpDt").Value, CS_DateMask)
            End If
            objRs.MoveNext
        Wend
        .RowHeight(-1) = 13.3
    End With

NoData:
    Set objRs = Nothing
    Exit Sub

Error_Trap:
    If Err.Number <> 94 Then
        MsgBox Err.Number & "  " & Err.Description
        Set objRs = Nothing
    End If

End Sub


Private Sub Lab005Move(ByRef MyReference As clsReference, ByVal Row As Long)
    With tblReference
        .Row = Row
        MyReference.TestCd = txtTestCd.Text
        MyReference.SpcCd = medGetP(cboSpcCd.Text, 1, " ")
        MyReference.ApplyDt = Format(txtAppDt.Value, CS_DateDbFormat)
        If IsNull(txtExpDt.Value) Then
        MyReference.ExpDt = ""
        Else
        MyReference.ExpDt = Format(txtExpDt.Value, CS_DateDbFormat)
        End If
        
        .Col = 2: MyReference.AgeFrom = Val(.Value)
        .Col = 3: MyReference.AgeTo = Val(.Value)
        
        '.Col = 4: MyReference.AgeDiv = Mid(.Value, 1, 1)
        
        .Col = 8
        If .Value <> "" Then
            .Col = 11: MyReference.RefText = .Value
        Else
            MyReference.RefText = ""
        End If
        '.Col = 8: MyReference.RefText = .Value
        
        .Col = 6: MyReference.RefCd = .Value
        If tabRef.SelectedItem.Index = 1 Then
            .Col = 1: MyReference.ApplySex = Choose(.TypeComboBoxCurSel + 1, "M", "F", "B", "U")
            .Col = 4: MyReference.RefValFrom = Val(.Value)
            .Col = 5: MyReference.RefValTo = Val(.Value)
            .Col = 9: MyReference.ARefValFrom = Val(.Value)
            .Col = 10: MyReference.ARefValTo = Val(.Value)
        ElseIf tabRef.SelectedItem.Index = 2 Then
            .Col = 1: MyReference.ApplySex = Choose(.TypeComboBoxCurSel + 1, "X", "Y", "Z")
            .Col = 4: MyReference.PanicFrVal = Val(.Value)
            .Col = 5: MyReference.PanicToVal = Val(.Value)
            .Col = 9: MyReference.ArletFrVal = Val(.Value)
            .Col = 10: MyReference.ArletToVal = Val(.Value)
        Else
            .Col = 1: MyReference.ApplySex = Choose(.TypeComboBoxCurSel + 1, "X", "Y", "Z")
            .Col = 4: MyReference.AMRFrVal = Val(.Value)
            .Col = 5: MyReference.AMRToVal = Val(.Value)
'            .Col = 9: MyReference.ArletFrVal = 0
'            .Col = 10: MyReference.ArletToVal = 0
        End If
        MyReference.RefDiv = tabRef.SelectedItem.Index
    End With
End Sub


Private Sub ClearRtn()

    'txtAppDt.Text = CS_BlankMask
    InsertFlag = 0
    UpdateFlag = 0
    cmdNew.Caption = "추가"
    cmdNew.Enabled = True
    cmdEdit.Caption = "수정"
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False
    Call ClearTable
    txtComment.Text = ""

End Sub

Private Sub ClearTable()

    txtExpDt.Value = ""
    With tblReference
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = 50
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .RowHeight(-1) = 13.3
        .BlockMode = False
        .ReDraw = True
    End With

End Sub

Private Sub LockRtn(ByVal intPart As Integer, ByVal LockValue As Boolean)

    If LockValue Then
'        EnableValue = False
        tblReference.OperationMode = OperationModeRead
    Else
'        EnableValue = True
        tblReference.OperationMode = OperationModeNormal
    End If

'    If intPart = 1 Then txtAppDt.Enabled = EnableValue
    If intPart = 2 Then
        With tblReference
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 1: .Col2 = 3
            .BlockMode = True
            .Lock = False
            .BlockMode = False
        End With
    End If
'    txtExpDt.Enabled = EnableValue

End Sub

Public Sub Raise_TestCd_Keypress()
    Call txtTestCd_KeyPress(13)
End Sub

Public Sub Raise_cboSpcCd_Click()
    Call cboSpcCd_Click
End Sub

Private Function ConfirmExit() As Boolean

    Dim intResp As VbMsgBoxResult

    ConfirmExit = True
        
    If tblReference.DataRowCnt < 1 Then Exit Function
        
    If InsertFlag = 1 Or UpdateFlag = 1 Then
        intResp = MsgBox("변경된 내용을 취소하시겠습니까 ? ", vbYesNo)
        If intResp = vbNo Then
            ConfirmExit = False
            Exit Function
        End If
        InsertFlag = 0
        UpdateFlag = 0
    End If

End Function


