VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSystem_manager 
   BackColor       =   &H00E8EEEE&
   BorderStyle     =   1  '단일 고정
   Caption         =   "화면설정"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   Icon            =   "frmSystem_manager.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12375
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E8EEEE&
      Caption         =   "닫기"
      Height          =   555
      Left            =   11055
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   7395
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E8EEEE&
      Caption         =   "설정"
      Height          =   555
      Left            =   9810
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   7395
      Width           =   1200
   End
   Begin MSComctlLib.TabStrip tabSubMenu 
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   825
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   529
      Style           =   2
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "채취/접수"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "결과등록"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "미생물/기타검사"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "조회/출력"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "QC"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "통계"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "종합검증/판독"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "기타"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrSubTool 
      Height          =   525
      Left            =   105
      TabIndex        =   1
      Top             =   135
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   926
      ButtonWidth     =   609
      ButtonHeight    =   926
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblForm 
      Height          =   5565
      Left            =   75
      TabIndex        =   2
      Top             =   1770
      Width           =   12195
      _Version        =   196608
      _ExtentX        =   21511
      _ExtentY        =   9816
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   50
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmSystem_manager.frx":000C
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   10
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00EEEBED&
      FillStyle       =   0  '단색
      Height          =   1020
      Left            =   90
      Top             =   120
      Width           =   12225
   End
   Begin VB.Label lblSubMenu 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "임상병리 화면 사용 권한 부여"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00794444&
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1335
      Width           =   3975
   End
   Begin VB.Shape shpSubMenu 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00EEEBED&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   105
      Top             =   1230
      Width           =   4065
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00EEEBED&
      FillStyle       =   0  '단색
      Height          =   525
      Left            =   90
      Top             =   1215
      Width           =   4095
   End
End
Attribute VB_Name = "frmSystem_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strCdIndex As String

Private Sub TabSubMenuQuery(ByVal FROMIndex As Integer, ByVal ToIndex As Integer)
    Dim objFrm      As clsDictionary
    Dim Rs          As Recordset
    Dim SSQL        As String
    Dim strTmp      As String
    Dim strKey      As String
    Dim aryTmp()    As String
    Dim ii          As Integer
    Dim jj          As Integer
    Dim kk          As Integer
    
    Set Rs = New Recordset
    Set objFrm = New clsDictionary
    objFrm.Clear
    objFrm.FieldInialize "key", "ii"
    
    Call medClearTable(tblForm)
    
    With tblForm
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellType = CellTypeStaticText
        .Value = ""
        .BlockMode = False
        .ReDraw = False
        For ii = FROMIndex To ToIndex 'mainfrm.tabSubMenu.Tabs.Count
            Call objFrm.DeleteAll
            SSQL = " SELECT * FROM " & T_LAB032 & _
                   " WHERE " & _
                             DBW("cdindex=", strCdIndex) & _
                   " AND " & DBW("cdval1=", ii)
                   ' LC3_HosFrmUseing
            Set Rs = Nothing
            Set Rs = New Recordset
            
            Rs.Open SSQL, DBConn
            If Not Rs.EOF Then
                strTmp = Rs.Fields("text1").Value & ""
                aryTmp = Split(strTmp, ";")
                For kk = LBound(aryTmp()) To UBound(aryTmp())
                    objFrm.AddNew aryTmp(kk), ii
                Next
            End If
            
            If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .DataRowCnt + 1
            .Col = 1: .Value = MainFrm.tabSubMenu.Tabs.Item(ii).Caption
                      .ForeColor = DCM_LightBlue: .FontBold = True: .TypeHAlign = TypeHAlignLeft
            
            If MainFrm.imlSubList(ii - 1).ListImages.Count <> 0 Then
                
                For jj = 1 To MainFrm.imlSubList(ii - 1).ListImages.Count
                    strTmp = MainFrm.imlSubList(ii - 1).ListImages(jj).Tag
                    If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .Col = 1:   .Value = Space(5) & medGetP(strTmp, 1, "(")
                                .TypeHAlign = TypeHAlignLeft
                    .Col = 2:   .Value = MainFrm.imlSubList(ii - 1).ListImages(jj).Key
                                .TypeHAlign = TypeHAlignRight
                                strKey = .Value
                    .Col = 3:   .Value = MainFrm.imlSubList(ii - 1).ListImages(jj).Index
                                .TypeHAlign = TypeHAlignRight
                    .Col = 4:   .Value = Space(5) & Trim(medGetP(medGetP(strTmp, 2, "("), 1, ")"))
                                .TypeHAlign = TypeHAlignLeft
                    .Col = 6:   .CellType = CellTypeCheckBox: .TypeCheckCenter = True
                                If objFrm.Exists(strKey) Then .Value = 1
                    .Col = 7: .Value = ii
                Next
                
            End If
        Next
        .ReDraw = True
    End With
    Set Rs = Nothing
    Set objFrm = Nothing
End Sub

Private Sub cmdClose_Click()
    
    Dim Frm As Form
    
    For Each Frm In Forms
        If (Frm.Name <> Me.Name) And (Frm.Name <> MainFrm.Name) Then
            Unload Frm
        End If
    Next
    
    MainFrm.tabSubMenu.Tabs(1).Selected = True
    Unload Me
End Sub

Private Sub Form_Activate()
    Call IniFormShow
    tabSubMenu.Tabs(1).Selected = True
    If ObjSysInfo.ProjectId = "LIS" Then
        lblSubMenu.Caption = "임상병리 화면 사용 권한부여"
    ElseIf ObjSysInfo.ProjectId = "BBS" Then
        lblSubMenu.Caption = "혈액은행 화면 사용 권한부여"
    End If
    
    If ObjSysInfo.ProjectId = "LIS" Then
        strCdIndex = "C261" '"C257"
    ElseIf ObjSysInfo.ProjectId = "BBS" Then
        strCdIndex = "C262"
    End If
End Sub

Private Sub Form_Load()
'    Call medAlwaysOn(frmSystem_manager, 1)
End Sub

Private Sub tabSubMenu_Click()
    
    Dim i       As Integer
    Dim intIDX  As Integer
    Dim strTag  As String
    
    Dim objFrm      As clsDictionary
    Dim Rs          As Recordset
    Dim SSQL        As String
    Dim strTmp      As String
    Dim strKey      As String
    Dim aryTmp()    As String
    Dim kk          As Integer
    
    ' Job Group 선택....Sub Toolbar의 내용이 바뀐다.
    intIDX = tabSubMenu.SelectedItem.Index
    
    'Tab별 화면 조회
    Call TabSubMenuQuery(intIDX, intIDX)
    
    Set Rs = New Recordset
    Set objFrm = New clsDictionary
    objFrm.Clear
    objFrm.FieldInialize "key", "ii"
    Call objFrm.DeleteAll
    SSQL = " SELECT * FROM " & T_LAB032 & _
           " WHERE " & _
                     DBW("cdindex=", strCdIndex) & _
           " AND " & DBW("cdval1=", intIDX)
    Rs.Open SSQL, DBConn
    If Not Rs.EOF Then
        strTmp = Rs.Fields("text1").Value & ""
        aryTmp = Split(strTmp, ";")
        For kk = LBound(aryTmp()) To UBound(aryTmp())
            objFrm.AddNew aryTmp(kk), intIDX
        Next
    End If
    Set Rs = Nothing
    
    ' 올라있던 버튼을 삭제
    For i = tbrSubTool.Buttons.Count To 1 Step -1
        Call tbrSubTool.Buttons.Remove(i)
    Next i
    
    If MainFrm.imlSubList(intIDX - 1).ListImages.Count = 0 Then
        Set objFrm = Nothing
        Exit Sub
    End If
    
    tbrSubTool.ImageList = MainFrm.imlSubList(intIDX - 1)
    kk = 0
    ' 버튼을 다시 그린다.
    For i = 1 To MainFrm.imlSubList(intIDX - 1).ListImages.Count
        strTag = MainFrm.imlSubList(intIDX - 1).ListImages(i).Tag
        If ObjSysInfo.ProjectId = "LIS" Then
            If strTag <> "-" Then
                strKey = MainFrm.imlSubList(intIDX - 1).ListImages(i).Key
                If Not objFrm.Exists(strKey) Then
                    kk = kk + 1
                    If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
                        Call tbrSubTool.Buttons.Add(kk, MainFrm.imlSubList(intIDX - 1).ListImages(i).Key, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
                    Else
                        Call tbrSubTool.Buttons.Add(kk, MainFrm.imlSubList(intIDX - 1).ListImages(i).Key, , , i)
                    End If
                    tbrSubTool.Buttons(kk).ToolTipText = strTag
                    tbrSubTool.Buttons(kk).Tag = strTag
                End If
            Else
                kk = kk + 1
                Call tbrSubTool.Buttons.Add(kk, , , tbrSeparator, i)
            End If
        ElseIf ObjSysInfo.ProjectId = "BBS" Then
            If strTag <> "-" Then
                strKey = MainFrm.imlSubList(intIDX - 1).ListImages(i).Key
                If Not objFrm.Exists(strKey) Then
                    kk = kk + 1
                    Call tbrSubTool.Buttons.Add(kk, MainFrm.imlSubList(intIDX - 1).ListImages(i).Key, , , i)
                    tbrSubTool.Buttons(kk).ToolTipText = strTag
                    tbrSubTool.Buttons(kk).Tag = strTag
                End If
            Else
                kk = kk + 1
                Call tbrSubTool.Buttons.Add(kk, , , tbrSeparator, i)
            End If
        ElseIf ObjSysInfo.ProjectId = "APS" Then
            If strTag <> "-" Then
                strKey = MainFrm.imlSubList(intIDX - 1).ListImages(i).Key
                If Not objFrm.Exists(strKey) Then
                    kk = kk + 1
                    Call tbrSubTool.Buttons.Add(kk, MainFrm.imlSubList(intIDX - 1).ListImages(i).Key, , , i)
                    tbrSubTool.Buttons(kk).ToolTipText = strTag
                    tbrSubTool.Buttons(kk).Tag = strTag
                End If
            Else
                kk = kk + 1
                Call tbrSubTool.Buttons.Add(kk, , , tbrSeparator, i)
            End If
        End If
    Next i
    Set objFrm = Nothing
End Sub

Private Sub IniFormShow()
    Dim ii As Integer
    
    tabSubMenu.Tabs.Clear

    For ii = 1 To MainFrm.tabSubMenu.Tabs.Count
        tabSubMenu.Tabs.Add ii, , MainFrm.tabSubMenu.Tabs.Item(ii).Caption
    Next
End Sub

Private Sub cmdSave_Click()
    Dim SSQL        As String
    Dim strCdval1   As String
    Dim strText     As String
    Dim strSave     As String
    Dim ii          As Integer
    
    
    strSave = UCase(InputBox("개발자 Passward 입력", "개발자 PassWard 확인"))
    If strSave <> UCase("system_manager") Then
        MsgBox "비밀번호가 일치하지 않습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    On Error GoTo SAVE_ERROR
    
    With tblForm
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 7: strCdval1 = .Value
            .Col = 6
            If .CellType = CellTypeCheckBox Then
                If .Value = 1 Then
                    .Col = 2
                    strText = strText & .Value & ";"
                End If
            End If
        Next
    End With
    DBConn.BeginTrans
    
    If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    
    SSQL = " delete " & T_LAB032 & _
           " WHERE " & _
                     DBW("cdindex=", strCdIndex) & _
           " AND " & DBW("cdval1=", strCdval1)
           
    DBConn.Execute SSQL
    
    If strText <> "" Then
        SSQL = " insert into " & T_LAB032 & "(cdindex,cdval1,text1) " & _
               " values(" & _
                DBV("cdindex", strCdIndex, 1) & DBV("cdval1", strCdval1, 1) & _
                DBV("text1", strText) & _
               ")"
        DBConn.Execute SSQL
    End If
    
    DBConn.CommitTrans
    tabSubMenu.Tabs(CLng(strCdval1)).Selected = True
    Exit Sub
    
SAVE_ERROR:
    DBConn.RollbackTrans
'LC3_HosFrmUseing
End Sub
