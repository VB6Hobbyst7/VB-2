VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm410PWardColList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FCEFE9&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2850
      Left            =   90
      TabIndex        =   4
      Top             =   1005
      Width           =   10740
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
         Height          =   345
         Left            =   2355
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   705
         Width           =   285
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
         Height          =   345
         Left            =   1275
         TabIndex        =   7
         Top             =   705
         Width           =   1065
      End
      Begin VB.ComboBox cboBuildings 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1275
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   1740
         Width           =   3495
      End
      Begin VB.ComboBox cboWorkTm 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3045
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   1395
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1275
         TabIndex        =   0
         Top             =   1395
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62390275
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   345
         Left            =   2685
         TabIndex        =   6
         Top             =   705
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   609
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   705
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "병동코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   270
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1365
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "채혈리스트 출력 - 병동 일괄채혈"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   795
      TabIndex        =   5
      Top             =   555
      Width           =   4965
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0088A2B7&
      FillColor       =   &H00DEEEFE&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   105
      Shape           =   4  '둥근 사각형
      Top             =   465
      Width           =   6390
   End
End
Attribute VB_Name = "frm410PWardColList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPWACoHeRow As Integer
Dim fWhich As Object
Dim iPageWidth As Integer
Dim iPageHeight As Integer
Dim iCurY As Integer
Dim DataExist As Boolean
Dim sLastDt As String
Dim sLastTm As String
Dim iRecordCount As Integer

Dim pWardId As String
Dim pWardNm As String
Dim pWorkDt As String
Dim pWorkTm As String
Dim pBuildCd As String
Dim pBuildNm As String
Dim pTitleNm As String
    
Dim SvBuildCd As String
Dim SvBuildNm As String

Dim strBuildCd As String
Dim strBuildNm As String
Dim intBuildNo As Integer

Const iCm = 567
Const iLineHeight = 10

Dim iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%


Public Event FormClose()


Public Property Get WardID() As String
    WardID = pWardId
End Property

Public Property Let WardID(ByVal vNewValue As String)
    pWardId = vNewValue
End Property

Public Property Get WardNm() As String
    WardNm = pWardNm
End Property

Public Property Let WardNm(ByVal vNewValue As String)
    pWardNm = vNewValue
End Property

Public Property Get WorkDt() As String
    WorkDt = pWorkDt
End Property

Public Property Let WorkDt(ByVal vNewValue As String)
    pWorkDt = vNewValue
End Property

Public Property Get WorkTm() As String
    WorkTm = pWorkTm
End Property

Public Property Let WorkTm(ByVal vNewValue As String)
    pWorkTm = vNewValue
End Property

Public Property Get BuildCd() As String
    BuildCd = pBuildCd
End Property

Public Property Let BuildCd(ByVal vNewValue As String)
    pBuildCd = vNewValue
End Property

Public Property Get BuildNm() As String
    BuildNm = pBuildNm
End Property

Public Property Let BuildNm(ByVal vNewValue As String)
    pBuildNm = vNewValue
End Property

Public Property Get TitleNm() As String
    TitleNm = pTitleNm
End Property

Public Property Let TitleNm(ByVal vNewValue As String)
    pTitleNm = vNewValue
End Property



Private Sub cboWorkTm_Click()
    
    Dim tmpRs As Recordset
    Dim SqlStmt As String
    Dim i As Integer
    
    SqlStmt = " Select a.buildcd as BuildCd, b.field1 as BuildNm " & _
              " From   " & T_LAB204 & " a, " & T_LAB032 & " b " & _
              " Where  " & DBW("a.workdt", Format(dtpDate.Value, CS_DateDbFormat), 2) & _
              " and    " & DBW("a.wardid", txtWardId.Text, 2) & _
              " and    " & DBW("a.worktm", Format(cboWorkTm.Text, CS_TimeDbFormat), 2) & _
              " and    " & DBW("b.cdindex", LC3_Buildings, 2) & _
              " and    b.cdval1 = a.buildcd " & _
              " Group by a.buildcd, b.field1 "
    
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    
    cboBuildings.Clear
    For i = 1 To tmpRs.RecordCount
        cboBuildings.AddItem tmpRs.Fields("BuildCd").Value & "   " & Trim("" & tmpRs.Fields("BuildNm").Value)
        tmpRs.MoveNext
    Next
    
    Set tmpRs = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
'   Set frm411PCollectionList = Nothing

    RaiseEvent FormClose
End Sub



Private Sub cmdReport_Click()
        
    '임상,해부,혈액은행의 코딩이 다 완료되지 않아서 기능 막음
    
    MsgBox "현재 프로그램을 보완 적용중 입니다." & vbNewLine & _
        "프로그램의 업그래이드가 완료 된 후 사용하십시요.", vbInformation, "프로그램 개발 안내"
    
    If txtWardId.Text = "" Then
         MsgBox "병동코드를 선택하세요.", vbInformation
         txtWardId.SetFocus
         Exit Sub
    End If
    If cboWorkTm.ListIndex < 0 Then
         MsgBox "작업시간을 선택하세요.", vbInformation
         cboWorkTm.SetFocus
         Exit Sub
    End If
    If cboBuildings.ListIndex < 0 Then
         MsgBox "건물코드를 선택하세요.", vbInformation
         cboBuildings.SetFocus
         Exit Sub
    End If
    
    Dim MyReport As New clsWardColList
    
    With MyReport
        .WardID = txtWardId.Text
        .WardNm = lblWardNm.Caption
        .WorkDt = Format(dtpDate.Value, CS_DateDbFormat)
        .WorkTm = Format(cboWorkTm.Text, CS_TimeDbFormat)
        .BuildCd = medGetP(cboBuildings.Text, 1, " ")
        .BuildNm = medGetP(cboBuildings.Text, 4, " ")
    
        Call .Print_ColList
    End With
    
    Set MyReport = Nothing
    
End Sub
    



Private Sub cmdWardList_Click()

'% 병동코드 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
'    Dim objWard As New clsBasisData

    With objMyList
        .FormCaption = "병동 조회"

        .ColumnHeaderText = "병동코드;병동명"
        .Tag = "WardID"
        Me.ScaleMode = 1
'        Call .ListPop(, 3950, 6300, ObjLISComCode.WardID)
        Call .LoadPopUp(GetSQLWardList) ', 3950, 6300)

        txtWardId.Text = medGetP(.SelectedString, 1, ";")
        lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
    
    End With

    Set objMyList = Nothing
'    Set objWard = Nothing


End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
   
    Dim tmpRs As Recordset
    Dim SqlStmt As String
    Dim i As Integer
    
    SqlStmt = " Select worktm " & _
              " From   " & T_LAB204 & " a " & _
              " Where  " & DBW("workdt", Format(dtpDate.Value, CS_DateDbFormat), 2) & _
              " and    " & DBW("wardid", txtWardId.Text, 2) & _
              " Group by worktm Order by worktm "
    
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    cboWorkTm.Clear
    For i = 1 To tmpRs.RecordCount
        cboWorkTm.AddItem Format(Trim("" & tmpRs.Fields("WorkTm").Value), CS_TimeLongMask)
        tmpRs.MoveNext
    Next
    
    Set tmpRs = Nothing
End Sub

Private Sub Form_Load()
   dtpDate.Value = Now
End Sub

Public Sub LoadColTime()

   Dim tmpRs As Recordset
   Dim SqlStmt As String
   
   SqlStmt = " select field1 as ColDate, field2 as ColTime " & _
             " from   " & T_LAB031 & _
             " where  " & DBW("cdindex", LC2_ColListTm, 2) & _
             " and    " & DBW("cdval1", strBuildCd, 2) & _
             " and    " & DBW("cdval2", txtWardId.Text, 2)
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs.EOF Then
      sLastDt = Format(Now, CS_DateDbFormat)
      sLastTm = Format(DateAdd("h", -2, Now), CS_TimeDbFormat)    '2시간 전
      DataExist = False
   Else
      sLastDt = Trim(tmpRs.Fields("ColDate").Value)
      sLastTm = Trim(tmpRs.Fields("ColTime").Value)
      DataExist = True
   End If
   
   Set tmpRs = Nothing
End Sub


Private Sub LoadBuildingInfo()

    Dim SqlStmt As String
'    Dim MySql As New clsLISSqlStatement
    Dim tmpRs As Recordset
    
   SqlStmt = GetSQLWard(txtWardId.Text) 'MySql.SqlHIS003CodeList(txtWardId.Text)
   
   
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If Not tmpRs.EOF Then
        strBuildCd = tmpRs.Fields("BldGb").Value
   Else
        strBuildCd = "10"
   End If

   
   SqlStmt = " select field1 as BldNm, field2 as BldNo " & _
             " from   " & T_LAB032 & _
             " where  " & DBW("cdindex", LC3_Buildings, 2) & _
             " and    " & DBW("cdval1", strBuildCd, 2)
   Set tmpRs = Nothing
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If Not tmpRs.EOF Then
        strBuildNm = tmpRs.Fields("BldNm").Value
        intBuildNo = Val(tmpRs.Fields("BldNo").Value)
   Else
        strBuildNm = "중앙"
        intBuildNo = 1
   End If
    
   Set tmpRs = Nothing
'   Set MySql = Nothing
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If txtWardId.Text = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtWardId_LostFocus()
   
    Call LoadColTime
    Call LoadBuildingInfo
    Call dtpDate_Validate(False)
    cboWorkTm.Clear
    cboBuildings.Clear
   
End Sub

Private Sub txtWardId_Validate(Cancel As Boolean)
'    Dim objWard As clsBasisData
    Dim strWard As String
    
'    Set objWard = New clsBasisData
    strWard = GetWardNm(txtWardId.Text)
'    Set objWard = Nothing
    
    If strWard <> "" Then
        lblWardNm.Caption = strWard
    Else
        Cancel = True
        lblWardNm.Caption = ""
        MsgBox "병동코드를 다시 입력하세요.", vbInformation, "채혈리스트출력"
    End If
    
'    If ObjLISComCode.WardID.Exists(txtWardId.Text) Then
'        ObjLISComCode.WardID.KeyChange (txtWardId.Text)
'        lblWardNm.Caption = ObjLISComCode.WardID.Fields("wardnm")
'    Else
'        Cancel = True
'        lblWardNm.Caption = ""
'        MsgBox "병동코드를 다시 입력하세요.", vbInformation, "채혈리스트출력"
'    End If

End Sub
