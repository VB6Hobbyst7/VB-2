VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmSearchPt 
   BackColor       =   &H00DBE6E6&
   Caption         =   "환자 검색"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmSearchPt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdConfig 
      BackColor       =   &H00F4F0F2&
      Caption         =   "설정 저장"
      Height          =   510
      Left            =   3255
      Style           =   1  '그래픽
      TabIndex        =   19
      Top             =   -15
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   4695
      Left            =   45
      TabIndex        =   9
      Top             =   1920
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16643054
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   1725
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   344
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 환자 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   1110
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   344
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 환자명으로 검색"
      Appearance      =   0
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00D7E6E6&
      Height          =   510
      Left            =   53
      TabIndex        =   12
      Tag             =   "136"
      Top             =   1215
      Width           =   4515
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Clear"
         Height          =   300
         Left            =   3750
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   135
         Width           =   570
      End
      Begin VB.OptionButton optKey 
         BackColor       =   &H00D7E6E6&
         Caption         =   "&ID"
         Height          =   240
         Index           =   0
         Left            =   2205
         TabIndex        =   15
         Tag             =   "15304"
         Top             =   195
         Width           =   495
      End
      Begin VB.OptionButton optKey 
         BackColor       =   &H00D7E6E6&
         Caption         =   "&Name"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Tag             =   "15305"
         Top             =   180
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.TextBox txtSearchKey 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   13
         Top             =   150
         Width           =   1830
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   510
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   344
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 검색 조건"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   15
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   344
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 외래/병동 구분(병동환자에서 검색)"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   390
      Left            =   53
      TabIndex        =   1
      Top             =   120
      Width           =   3195
      Begin VB.OptionButton optPtFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "구분없음"
         Height          =   180
         Index           =   2
         Left            =   1860
         TabIndex        =   7
         Top             =   150
         Width           =   1020
      End
      Begin VB.OptionButton optPtFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "병동"
         Height          =   180
         Index           =   1
         Left            =   975
         TabIndex        =   3
         Top             =   150
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton optPtFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "외래"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   150
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   495
      Left            =   53
      TabIndex        =   6
      Top             =   615
      Width           =   4515
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00F4F0F2&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1830
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   135
         Width           =   300
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00D7E6E6&
         Caption         =   "전체"
         Height          =   180
         Left            =   60
         TabIndex        =   10
         Top             =   195
         Width           =   660
      End
      Begin VB.CheckBox chkToday 
         BackColor       =   &H00D7E6E6&
         Caption         =   "금일 보고 대상만"
         ForeColor       =   &H00553755&
         Height          =   180
         Left            =   2430
         TabIndex        =   8
         Top             =   210
         Value           =   1  '확인
         Width           =   1680
      End
      Begin VB.Label lblCodeNm 
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   720
         TabIndex        =   18
         Top             =   150
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSearchPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ListClick(ByVal PtId As String)

Private WithEvents objCode As clsPopUpList
Attribute objCode.VB_VarHelpID = -1

Private blnSearch As Boolean
Private strSearch As String

Private Type tpConfig
    PtFg As Long
    All As Long
    CodeNm As String
    Today As Long
    Key As Long
End Type

Private Config As tpConfig

Private Sub chkALL_Click()
    If blnSearch Then lvwPtList.ListItems.Clear
    
    If chkAll.Value = 1 Then
        lblCodeNm.Caption = ""
    Else
'        If Trim(txtSearchKey.Text) <> "" Then
'            blnSearch = False
'            Call txtSearchKey_KeyDown(vbKeyReturn, 0)
'        End If
    End If
End Sub

Private Sub chkToday_Click()
    If blnSearch Then lvwPtList.ListItems.Clear
    
    If Trim(txtSearchKey.Text) <> "" Then
        blnSearch = False
        Call txtSearchKey_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub cmdClear_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
    blnSearch = False
    strSearch = ""
End Sub

Private Sub cmdConfig_Click()
    Dim strMsg As VbMsgBoxResult
    Dim strMsg2 As VbMsgBoxResult
    Dim User As String
    
    'Yes : 현재설정값을 저장
    'No : 현재설정을 무시하고 디폴트로 사용
    'Cancel : 아무짓도 안하구 나가기
    
    optPtFg(0).FontItalic = True
    optPtFg(1).FontItalic = True
    optPtFg(2).FontItalic = True
    chkAll.FontItalic = True
    lblCodeNm.FontItalic = True
    chkToday.FontItalic = True
    optKey(0).FontItalic = True
    optKey(1).FontItalic = True
    
    strMsg = MsgBox("지정할 수 있는 설정은 이탤릭으로 표시된 것들입니다.." & vbNewLine & _
                    "다시 화면을 열때는 설정된 값으로 변경됩니다." & vbNewLine & vbNewLine & _
                    "현재 설정을 저장하시겠습니까?" & vbNewLine & _
                    "(예:설정 저장,아니오:설정 초기화)", vbYesNoCancel + vbExclamation)
    
    If strMsg = vbCancel Then GoTo NoAction
    
    '설정저장
    
    User = GetSetting("Schweitzer2000 LIS\Config", "frmSearchPt", "User", "")
        
    If strMsg = vbYes Then
        
        If lblCodeNm.Caption <> "" Then
            strMsg2 = MsgBox("시작시 데이터를 조회하는 기능이 있습니다." & vbNewLine & _
                             "이 기능은 다소 프로그램 시작시 느리게 느껴질 수도 있습니다." & vbNewLine & vbNewLine & _
                             "이 기능을 사용하시겠습니까?", vbExclamation + vbYesNo)
            
            If strMsg2 = vbYes Then
                Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Start", "1")
            Else
                Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Start", "0")
            End If
        Else
            Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Start", "0")
        End If
        
        If InStr(User, ObjSysInfo.logonid) = 0 Then '새로지정
            Call SaveSetting("Schweitzer2000 LIS\Config", "frmSearchPt", "User", User & ObjSysInfo.logonid & ",")
        End If
        
        Call SaveSetting("Schweitzer2000 LIS\Config", "frmSearchPt", "Desc", "환자 검색(결과조회에서 사용되는)")
        
        Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "All", chkAll.Value)
        Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "CodeNm", lblCodeNm.Caption)
        Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "PtFg", IIf(optPtFg(0).Value, 0, IIf(optPtFg(1).Value, 1, 2)))
        Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Key", IIf(optKey(0).Value, 0, 1))
        Call SaveSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Today", chkToday.Value)
        
        MsgBox "현재의 설정이 지정되었습니다.", vbExclamation
    Else
        User = Replace(User, ObjSysInfo.logonid, "")
        User = Replace(User, ",,", ",")
        If Len(User) = 1 Then User = ""
        Call SaveSetting("Schweitzer2000 LIS\Config", "frmSearchPt", "User", User)
        
        On Error Resume Next
        Call DeleteSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid)
        
        MsgBox "설정이 초기화되었습니다.", vbExclamation
    End If

NoAction:
    optPtFg(0).FontItalic = False
    optPtFg(1).FontItalic = False
    optPtFg(2).FontItalic = False
    chkAll.FontItalic = False
    lblCodeNm.FontItalic = False
    chkToday.FontItalic = False
    optKey(0).FontItalic = False
    optKey(1).FontItalic = False
End Sub

Private Sub cmdSearch_Click()
'    Dim objData As clsBasisData
    
'    Set objData = New clsBasisData
    Set objCode = New clsPopUpList
    With objCode
        medAlwaysOn frmSearchPt, 0
        .Connection = DBConn
        .FormCaption = IIf(optPtFg(0).Value, "진료과 조회", "병동 조회")
        .Tag = "WardId"
        .ColumnHeaderText = IIf(optPtFg(0).Value, "과코드;진료과", "병동코드;병동명")
        If optPtFg(0).Value Then
            .LoadPopUp GetSQLDeptList
        Else
            .LoadPopUp GetSQLWardList
        End If
        
        lblCodeNm.Caption = .SelectedItems(1)
        
'        Call .loadpopup(, , , IIf(optPtFg(0).Value, ObjLISComCode.DeptCd, ObjLISComCode.WardId))
        medAlwaysOn frmSearchPt, 1
    End With
    
'    Set objData = Nothing
End Sub

Private Sub Form_Load()
    lvwPtList.ColumnHeaders.Clear
    medInitLvwHead lvwPtList, "환자ID,환자성명,병동/병실,주민번호,생년월일,성별/나이,주소,전화번호", _
                       "400,400,1000,1400,500,500,800,700"
    
    blnSearch = False
    strSearch = ""
    
    Call InitForm
    Call ReadConfig
    
    Dim Start As String

    medAlwaysOn frmSearchPt, 1
    
    Start = GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Start", "")
    
    If (Start = "1") And (lblCodeNm.Caption <> "") Then
        DoEvents
        Call LoadData
    End If
End Sub

Private Sub InitForm()
    optPtFg(1).Value = True
    chkAll.Value = 1
    chkToday.Value = 1
    optKey(1).Value = True
    
    With Config
        .PtFg = 1
        .All = 1
        .Today = 1
        .CodeNm = ""
        .Key = 1
    End With
End Sub

Private Sub ReadConfig()
    Dim User As String
    
    User = GetSetting("Schweitzer2000 LIS\Config", "frmSearchPt", "User", "")
    
    If InStr(User, ObjSysInfo.logonid) = 0 Then Exit Sub
    
    optPtFg(Val(GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "PtFg", ""))).Value = 1
    chkAll.Value = Val(GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "All", ""))
    chkToday.Value = Val(GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Today", ""))
    optKey(Val(GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "Key", ""))).Value = 1
    lblCodeNm.Caption = GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "CodeNm", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent ListClick("")
End Sub

Private Sub lblCodeNm_Change()
    If Trim(txtSearchKey.Text) <> "" Then
        blnSearch = False
        Call txtSearchKey_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent ListClick(Item.Text)
End Sub

Private Sub objCode_SendCode(ByVal SelString As String)
    
    lblCodeNm.Caption = medGetP(SelString, 1, ";")
    Set objCode = Nothing
    
    If lblCodeNm.Caption <> "" Then
        chkAll.Value = 0
    End If
End Sub

Private Sub optKey_Click(Index As Integer)
    If optKey(0).Value Then  'ID
        LisLabel4.Caption = "◈ 환자ID로 검색"
    Else '환자명
        LisLabel4.Caption = "◈ 환자명으로 검색"
    End If
    On Error Resume Next
    txtSearchKey.SetFocus
End Sub

Private Sub optPtFg_Click(Index As Integer)
    If optPtFg(0).Value Then '외래
        LisLabel1.Caption = "◈ 외래/병동 구분(외래환자에서 검색)"
        cmdSearch.Enabled = False
        chkAll.Value = 1
        chkAll.Enabled = False
       
        lvwPtList.ColumnHeaders.Clear
        medInitLvwHead lvwPtList, "환자ID,환자성명,진료과,주민번호,생년월일,성별/나이,주소,전화번호", _
                           "400,400,1000,1400,500,500,800,700"
        
        lvwPtList.ColumnHeaders(3).Width = 0
       
    ElseIf optPtFg(1).Value Then '병동
        LisLabel1.Caption = "◈ 외래/병동 구분(병동환자에서 검색)"
        cmdSearch.Enabled = True
        chkAll.Enabled = True
        
        Dim User As String
        
        User = GetSetting("Schweitzer2000 LIS\Config", "frmSearchPt", "User", "")
        
        If InStr(User, ObjSysInfo.logonid) > 0 Then
            lblCodeNm.Caption = GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "CodeNm", "")
            chkAll.Value = Val(GetSetting("Schweitzer2000 LIS\Config\frmSearchPt", ObjSysInfo.logonid, "All", ""))
        Else
            chkAll.Value = 1
        End If
        
        lvwPtList.ColumnHeaders.Clear
        medInitLvwHead lvwPtList, "환자ID,환자성명,병동/병실,주민번호,생년월일,성별/나이,주소,전화번호", _
                           "400,400,1000,1400,500,500,800,700"
    ElseIf optPtFg(2).Value Then '구분없음
       LisLabel1.Caption = "◈ 외래/병동 구분(환자마스터에서 검색)"
       cmdSearch.Enabled = False
       chkAll.Value = 1
       chkAll.Enabled = False
       
        lvwPtList.ColumnHeaders.Clear
        medInitLvwHead lvwPtList, "환자ID,환자성명,진료과,주민번호,생년월일,성별/나이,주소,전화번호", _
                           "400,400,1000,1400,500,500,800,700"
        lvwPtList.ColumnHeaders(3).Width = 0
    End If
    
    If blnSearch Then lvwPtList.ListItems.Clear
    
    If Trim(txtSearchKey.Text) <> "" Then
        blnSearch = False
        Call txtSearchKey_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub txtSearchKey_Change()
    If lvwPtList.ListItems.Count = 0 Then Exit Sub
    
    If strSearch <> Mid(Trim(txtSearchKey.Text), 1, Len(strSearch)) Then
        lvwPtList.ListItems.Clear
        blnSearch = False
    End If
    
    Dim strFindItem As String
    Dim itmFound As ListItem   ' FoundItem 변수입니다.
    Dim itmx As ListItem
    Dim I As Long
        
    strFindItem = Trim(txtSearchKey.Text)
    
    With lvwPtList
        If optKey(0).Value Then
            For I = 1 To .ListItems.Count
                Set itmx = .ListItems(I)
                If UCase(itmx.Text) Like UCase(strFindItem & "*") Then
                    itmx.Selected = True
                    itmx.EnsureVisible
                    Exit For
                End If
            Next
        Else
            For I = 1 To .ListItems.Count
                Set itmx = .ListItems(I)
                If (itmx.SubItems(1) Like (strFindItem & "*")) Then
                    itmx.Selected = True
                    itmx.EnsureVisible
                    Exit For
                End If
            Next
        End If
    End With
    
End Sub

Private Sub txtSearchKey_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
'환자명,환자ID로 검색조건의 환자가 외래/병동 구분으로 환자 마스터에 존재하는지 먼저 검색
    
    Dim RsExPt  As Recordset
    Dim strSQL As String
    Dim itmx As ListItem
   
    If Trim(txtSearchKey.Text) = "" Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If (chkAll.Value = 0) And (lblCodeNm.Caption = "") Then
        MsgBox Mid(LisLabel1.Caption, InStr(LisLabel1.Caption, "(") + 1, InStr(LisLabel1.Caption, "환") - InStr(LisLabel1.Caption, "(") - 1) & " 을(를) 선택하세요.", vbExclamation
        Exit Sub
    End If
    
    If blnSearch Then   '리스트에서 검색된 환자를 보낸다.
        If lvwPtList.ListItems.Count = 0 Then Exit Sub
        If txtSearchKey.Text <> IIf(optKey(0).Value, lvwPtList.SelectedItem.Text, lvwPtList.SelectedItem.SubItems(1)) Then
            MsgBox "검색한 환자가 존재하지 않습니다.", vbExclamation
        Else
            RaiseEvent ListClick(lvwPtList.SelectedItem.Text)
        End If
        
        Exit Sub
    Else
    strSearch = Trim(txtSearchKey.Text)
    
    If optPtFg(1).Value Then
        If (chkAll.Value = 1) And (chkToday.Value = 1) Then '재원마스터에 존재하고 처방내역의 Examdt가 박혀있는넘
            strSQL = " select distinct a.ptid as ptid,b." & F_PTNM & " as ptnm,c." & F_PTWARDID & " as wardid," & _
                   " c." & F_PTROOMID & " as roomid, " & F_SSN2("b") & " as ssn," & F_DOB2("b") & " as dob, " & _
                   " b." & F_SEX & " as sex,b." & F_ADDRESS & " as addr,b." & F_TEL & " as tel " & _
                   " from " & T_LAB102 & " a, " & T_HIS001 & " b," & T_HIS002 & " c " & _
                   " where a.ptId = b." & F_PTID & _
                   " and b." & F_PTID & "=c." & F_PTID & _
                   IIf(optKey(1).Value, " and b." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                        " and b." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                   " and " & DBW("a.examdt=", Format(GetSystemDate, "yyyyMMdd"))
        ElseIf (chkAll.Value = 0) And (chkToday.Value = 1) Then '재원마스터의 특정병동에 존재하고 처방내역의 Examdt가 박혀있는넘
            strSQL = " select distinct a.ptid as ptid,b." & F_PTNM & " as ptnm,c." & F_PTWARDID & " as wardid," & _
                   " c." & F_PTROOMID & " as roomid, " & F_SSN2("b") & " as ssn," & F_DOB2("b") & " as dob, " & _
                   " b." & F_SEX & " as sex,b." & F_ADDRESS & " as addr,b." & F_TEL & " as tel " & _
                   " from " & T_LAB102 & " a, " & T_HIS001 & " b," & T_HIS002 & " c " & _
                   " where a.ptId = b." & F_PTID & _
                   " and b." & F_PTID & "=c." & F_PTID & _
                   IIf(optKey(1).Value, " and b." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                        " and b." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                   " and " & DBW("a.examdt=", Format(GetSystemDate, "yyyyMMdd")) & _
                   " and " & DBW("c." & F_PTWARDID & "=", lblCodeNm.Caption)
        ElseIf (chkAll.Value = 1) And (chkToday.Value = 0) Then '재원마스터에 존재하는 넘
            strSQL = " select distinct a.ptid as ptid,b." & F_PTNM & " as ptnm,c." & F_PTWARDID & " as wardid," & _
                     " c." & F_PTROOMID & " as roomid, " & F_SSN2("b") & " as ssn," & F_DOB2("b") & " as dob, " & _
                     " b." & F_SEX & " as sex,b." & F_ADDRESS & " as addr,b." & F_TEL & " as tel " & _
                     " from " & T_LAB102 & " a, " & T_HIS001 & " b," & T_HIS002 & " c " & _
                     " where a.ptId = b." & F_PTID & _
                     " and b." & F_PTID & "=c." & F_PTID & _
                     IIf(optKey(1).Value, " and b." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                          " and b." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                     " union " & _
                     " select distinct a." & F_PTID & " as ptid, a." & F_PTNM & " as ptnm,b." & F_PTWARDID & " as ward, " & _
                     " b." & F_PTROOMID & " as roomid, " & F_SSN2("a") & " as ssn, " & F_DOB2("a") & " as dob, " & _
                     " a." & F_SEX & " as sex, a." & F_ADDRESS & " as addr, a." & F_TEL & " as tel " & _
                     " from " & T_HIS001 & " a, " & T_HIS002 & " b " & _
                     " where a." & F_PTID & "=b." & F_PTID & _
                     IIf(optKey(1).Value, " and a." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                          " and a." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text)))
        Else '재원마스터의 특정병동에 존재하는 넘
            strSQL = " select distinct a.ptid as ptid,b." & F_PTNM & " as ptnm,c." & F_PTWARDID & " as wardid," & _
                     " c." & F_PTROOMID & " as roomid, " & F_SSN2("b") & " as ssn," & F_DOB2("b") & " as dob, " & _
                     " b." & F_SEX & " as sex,b." & F_ADDRESS & " as addr,b." & F_TEL & " as tel " & _
                     " from " & T_LAB102 & " a, " & T_HIS001 & " b," & T_HIS002 & " c " & _
                     " where a.ptId = b." & F_PTID & _
                     " and b." & F_PTID & "=c." & F_PTID & _
                     IIf(optKey(1).Value, " and b." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                          " and b." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                     " and " & DBW("c." & F_PTWARDID & "=", lblCodeNm.Caption) & _
                     " union " & _
                     " select distinct a." & F_PTID & " as ptid, a." & F_PTNM & " as ptnm,b." & F_PTWARDID & " as wardid, " & _
                     " b." & F_PTROOMID & " as roomid," & F_SSN2("a") & " as ssn, " & F_DOB2("a") & " as dob, " & _
                     " a." & F_SEX & " as sex, a." & F_ADDRESS & " as addr, a." & F_TEL & " as tel " & _
                     " from " & T_HIS001 & " a, " & T_HIS002 & " b " & _
                     " where a." & F_PTID & "=b." & F_PTID & _
                     IIf(optKey(1).Value, " and a." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                          " and a." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                     " and " & DBW("b." & F_PTWARDID & "=", lblCodeNm.Caption) '& _
                     " and rownum < 100 "
        End If
     Else
        If chkToday.Value = 1 Then '환자마스터에 존재하고 처방내역의 Examdt가 박혀있는넘
            strSQL = " select distinct a.ptid as ptid,b." & F_PTNM & " as ptnm,'' as wardid," & _
                     " '' as roomid, " & F_SSN2("b") & " as ssn," & F_DOB2("b") & " as dob, " & _
                     " b." & F_SEX & " as sex,b." & F_ADDRESS & " as addr,b." & F_TEL & " as tel " & _
                     " from " & T_LAB102 & " a, " & T_HIS001 & " b " & _
                     " where a.ptid = b." & F_PTID & _
                     IIf(optKey(1).Value, " and b." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                          " and b." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                     " and " & DBW("a.examdt=", Format(GetSystemDate, "yyyyMMdd"))
        Else '환자마스터에 존재하는 넘
            strSQL = " select distinct a.ptid as ptid,b." & F_PTNM & " as ptnm,'' as wardid," & _
                     " '' as roomid, " & F_SSN2("b") & " as ssn," & F_DOB2("b") & " as dob, " & _
                     " b." & F_SEX & " as sex,b." & F_ADDRESS & " as addr,b." & F_TEL & " as tel " & _
                     " from " & T_LAB102 & " a, " & T_HIS001 & " b " & _
                     " where a.ptId = b." & F_PTID & _
                     IIf(optKey(1).Value, " and b." & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                     " and b." & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text))) & _
                     " union " & _
                     " select distinct " & F_PTID & " as ptid, " & F_PTNM & " as ptnm,'' as wardid," & _
                     " '' as roomid," & F_SSN2 & " as ssn," & F_DOB2 & " as dob, " & _
                     F_SEX & " As sex, " & F_ADDRESS & " As addr, " & F_TEL & " As tel " & _
                     " from " & T_HIS001 & _
                     IIf(optKey(1).Value, " where " & F_PTNM & " like " & DBV(F_PTNM, Trim(txtSearchKey.Text) & "%"), _
                                          " where " & F_PTID & " >= " & DBV(F_PTID, Trim(txtSearchKey.Text)))
        End If
    End If
    
    Me.MousePointer = vbHourglass
    
    Set RsExPt = New Recordset
    RsExPt.Open strSQL, DBConn
    
    If RsExPt.EOF Then
        lvwPtList.ListItems.Clear
        GoTo NoData
    End If
    
    lvwPtList.ListItems.Clear
    Do Until RsExPt.EOF
        Set itmx = lvwPtList.ListItems.Add
        itmx.Text = RsExPt.Fields("ptid").Value & ""
        itmx.SubItems(1) = RsExPt.Fields("ptnm").Value & ""
        itmx.SubItems(2) = RsExPt.Fields("wardid").Value & "" & "-" & _
                           RsExPt.Fields("roomid").Value & ""
        itmx.SubItems(3) = RsExPt.Fields("ssn").Value & ""
        itmx.SubItems(4) = RsExPt.Fields("dob").Value & ""
        itmx.SubItems(5) = RsExPt.Fields("sex").Value & "" & "/" & _
                           DateDiff("YYYY", Format(RsExPt.Fields("dob").Value & "", "####-##-##"), GetSystemDate)
        itmx.SubItems(6) = RsExPt.Fields("addr").Value & ""
        itmx.SubItems(7) = RsExPt.Fields("tel").Value & ""
        
        If lvwPtList.ListItems.Count = 1000 Then Exit Do
        RsExPt.MoveNext
    Loop
    End If
    
NoData:
    If lvwPtList.ListItems.Count = 0 Then
        MsgBox "조건에 맞는 자료가 없습니다.", vbExclamation
        blnSearch = False
    Else
        blnSearch = True
    End If
    
    Me.MousePointer = vbDefault
    Set RsExPt = Nothing
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim RsExPt As Recordset
    Dim itmx As ListItem
    
    If chkToday.Value = 0 Then '선택된 병동 환자 모조리 불러옴
        strSQL = "select a." & F_PTID & " as ptid,a." & F_PTNM & " as ptnm,b." & F_PTWARDID & " as wardid," & _
                 " b." & F_PTROOMID & " as roomid," & F_SSN2("a") & " as ssn," & F_DOB2("a") & " as dob," & _
                 " a." & F_SEX & " as sex,a." & F_ADDRESS & " as addr,a." & F_TEL & " as tel" & _
                 " from " & T_HIS001 & " a, " & T_HIS002 & " b " & _
                 " where " & DBW("b." & F_PTWARDID & "=", lblCodeNm.Caption) & _
                 " and b." & F_PTID & "=a." & F_PTID
    Else    '선택된 병동에 소속되고 보고일자가 오늘인넘들..
        strSQL = " select distinct c." & F_PTID & " as ptid,c." & F_PTNM & " as ptnm,b." & F_PTWARDID & " as wardid," & _
                 " b." & F_PTROOMID & " as roomid, " & F_SSN2("c") & " as ssn," & F_DOB2("c") & " as dob, " & _
                 " c." & F_SEX & " as sex,c." & F_ADDRESS & " as addr,c." & F_TEL & " as tel " & _
                 " from " & T_HIS001 & " c " & T_HIS002 & " b, " & T_LAB102 & " a" & _
                 " where " & DBW("a.examdt=", Format(GetSystemDate, "yyyyMMdd")) & _
                 "  and a.ptid=b." & F_PTID & _
                 " and " & DBW("b." & F_PTWARDID & "=", lblCodeNm.Caption) & _
                 " and (b.out_date is null) " & _
                 " and c.bunho=b.bunho "
    End If

    
    MousePointer = vbHourglass
    
    Set RsExPt = New Recordset
    RsExPt.Open strSQL, DBConn
    
    If RsExPt.EOF Then
        lvwPtList.ListItems.Clear
        GoTo NoData
    End If
    
    lvwPtList.ListItems.Clear
    Do Until RsExPt.EOF
        Set itmx = lvwPtList.ListItems.Add
        itmx.Text = RsExPt.Fields("ptid").Value & ""
        itmx.SubItems(1) = RsExPt.Fields("ptnm").Value & ""
        itmx.SubItems(2) = RsExPt.Fields("wardid").Value & "" & "-" & _
                           RsExPt.Fields("roomid").Value & ""
        itmx.SubItems(3) = RsExPt.Fields("ssn").Value & ""
        itmx.SubItems(4) = RsExPt.Fields("dob").Value & ""
        itmx.SubItems(5) = RsExPt.Fields("sex").Value & "" & "/" & _
                           DateDiff("YYYY", Format(RsExPt.Fields("dob").Value & "", "####-##-##"), GetSystemDate)
        itmx.SubItems(6) = RsExPt.Fields("addr").Value & ""
        itmx.SubItems(7) = RsExPt.Fields("tel").Value & ""
        
        If lvwPtList.ListItems.Count = 1000 Then Exit Do
        RsExPt.MoveNext
    Loop
    
NoData:
    If lvwPtList.ListItems.Count = 0 Then
        MsgBox "조건에 맞는 자료가 없습니다.", vbExclamation
        blnSearch = False
    Else
        blnSearch = True
    End If
    
    MousePointer = vbDefault
    Set RsExPt = Nothing
End Sub
