VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm419PColList 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11295
   ControlBox      =   0   'False
   Icon            =   "Lis419.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "출   력 (&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   6690
      Left            =   75
      ScaleHeight     =   6630
      ScaleWidth      =   10680
      TabIndex        =   9
      Top             =   1755
      Width           =   10740
      Begin FPSpread.vaSpread tblCollect 
         Height          =   6600
         Left            =   30
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   10650
         _Version        =   196608
         _ExtentX        =   18785
         _ExtentY        =   11642
         _StockProps     =   64
         BackColorStyle  =   3
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   10
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis419.frx":144A
         Appearance      =   1
      End
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   330
      Index           =   0
      Left            =   75
      TabIndex        =   11
      Top             =   1410
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "출력대상목록 조회 리스트"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   330
      Index           =   1
      Left            =   75
      TabIndex        =   12
      Top             =   45
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "채취리스트 조회조건"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1095
      Left            =   75
      TabIndex        =   7
      Top             =   300
      Width           =   10725
      Begin VB.ComboBox cboCol 
         Height          =   300
         Left            =   4290
         TabIndex        =   2
         Top             =   660
         Width           =   4290
      End
      Begin VB.CheckBox chkTestdiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사코드출력"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3105
         TabIndex        =   13
         Top             =   300
         Width           =   1425
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
         Height          =   360
         Left            =   1290
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
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
         Height          =   360
         Left            =   2670
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9315
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   345
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         Height          =   375
         Left            =   1290
         TabIndex        =   1
         Top             =   630
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   20840451
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
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
         Caption         =   "병   동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   105
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
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
         Caption         =   "채취일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   3090
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
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
         Caption         =   "채취일시"
         Appearance      =   0
      End
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frm419PColList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()

Private Enum TblColumn
    tcCHK = 1
    tcNO
    tcWORKNO
    tcPTNM
    tcPTID
    tcSA
    tcHOSILID
    tcCOLDT
    tcTEST
    tcSPCNM
End Enum

Private objMySql As New clsWardColList

Private Sub cboCol_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdClear_Click()
    
    Call FrmInitionalize
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub FrmInitionalize()
    txtWardId.Text = gWardid
    dtpColdt.Value = Format(GetSystemDate, "yyyy-mm-dd")
    tblCollect.MaxRows = 0
    cboCol.Clear
    cmdPrint.Enabled = False
    Call medClearTable(tblCollect)
End Sub

Private Sub cmdPrint_Click()
    Dim strChk     As String
    Dim strNo      As String
    Dim strWorkNo  As String
    Dim strPtNm    As String
    Dim strPtId    As String
    Dim strSa      As String
    Dim strHosilid As String
    Dim strcoldt   As String
    Dim strtest    As String
    Dim strSpcNm   As String
    Dim ii         As Integer
    
    Me.MousePointer = 11
    objMySql.objDictionary.DeleteAll
    With tblCollect
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn.tcCHK: strChk = .Value
            If strChk = "0" Then
                .Col = TblColumn.tcNO:      strNo = .Value
                .Col = TblColumn.tcWORKNO:  strWorkNo = .Value
                .Col = TblColumn.tcPTNM:    strPtNm = .Value
                .Col = TblColumn.tcPTID:    strPtId = .Value
                .Col = TblColumn.tcSA:      strSa = .Value
                .Col = TblColumn.tcHOSILID: strHosilid = .Value
                .Col = TblColumn.tcCOLDT:   strcoldt = .Value
                .Col = TblColumn.tcTEST:    strtest = .Value
                .Col = TblColumn.tcSPCNM:   strSpcNm = .Value
                
                If objMySql.objDictionary.Exists(strNo) Then
                    objMySql.objDictionary.KeyChange strNo
                    objMySql.objDictionary.Fields("workno") = strWorkNo
                    objMySql.objDictionary.Fields("ptnm") = strPtNm
                    objMySql.objDictionary.Fields("ptid") = strPtId
                    objMySql.objDictionary.Fields("sexage") = strSa
                    objMySql.objDictionary.Fields("testlist") = strtest
                    objMySql.objDictionary.Fields("spcnm") = strSpcNm
                    objMySql.objDictionary.Fields("hosilid") = strHosilid
                Else
                    If ICSResultChk = True Then
                        objMySql.objDictionary.AddNew strNo, _
                                                        Join(Array(strWorkNo, strPtNm, strPtId, strSa, "", _
                                                            "", strtest, _
                                                            strSpcNm, strHosilid), COL_DIV)
                    Else
                        objMySql.objDictionary.AddNew strNo, _
                                                        Join(Array(strWorkNo, strPtNm, strPtId, strSa, "", _
                                                            strcoldt, strtest, strSpcNm, strHosilid), COL_DIV)
                    
                    End If
                End If
            End If
        Next
    End With
    objMySql.SetCrpt CReport
    Call objMySql.RePrint_CollectList
    Call FrmInitionalize
    Me.MousePointer = 0
End Sub

Private Sub cmdQuery_Click()
    Dim ii     As Integer
    
    If txtWardId = "" Or cboCol.ListIndex < 0 Then
        MsgBox "조건을 입력하신후 조회작업을 하세요.", vbInformation + vbOKOnly, "채혈리스트"
         Exit Sub
    End If
    
    With objMySql
        .TitleNm = "병동 채혈리스트"
        .TestDiv = chkTestdiv.Value
        .WorkTm = Replace(medGetP(cboCol.Text, 1, Space(3)), ":", "")
        .objDictionary.DeleteAll
    End With
    Me.MousePointer = 11
    
    If objMySql.CollectQueryTF = True Then
        
        Dim objPrdBar As New jProgressBar.clsProgress
        With objPrdBar
            .Container = MainFrm.stsbar
            
'            .SetStsBar MAINFRM.STSBAR
            .Max = objMySql.objDictionary.RecordCount
            
        End With
        ii = 1
        With tblCollect
            .MaxRows = 0
            .MaxRows = objMySql.objDictionary.RecordCount
            
            Call medClearTable(tblCollect)
            objMySql.objDictionary.MoveFirst
            Do Until objMySql.objDictionary.EOF
                
                .Row = ii
                    
                .Col = TblColumn.tcNO:       .Value = objMySql.objDictionary.Fields("seq")
                .Col = TblColumn.tcWORKNO:   .Value = objMySql.objDictionary.Fields("workno")
                
                .Col = TblColumn.tcPTNM:     .Value = objMySql.objDictionary.Fields("ptnm") & _
                                                      ICSPatientString(objMySql.objDictionary.Fields("ptid"), enICSNum.LIS_ALL)
                
                .Col = TblColumn.tcPTID:     .Value = objMySql.objDictionary.Fields("ptid")
                .Col = TblColumn.tcSA:       .Value = objMySql.objDictionary.Fields("sexage")
                .Col = TblColumn.tcHOSILID:  .Value = objMySql.objDictionary.Fields("hosilid")
                .Col = TblColumn.tcCOLDT:    .Value = objMySql.objDictionary.Fields("collectdt")
                .Col = TblColumn.tcTEST:     .Value = objMySql.objDictionary.Fields("testlist")
                .Col = TblColumn.tcSPCNM:    .Value = objMySql.objDictionary.Fields("spcnm")
                objPrdBar.Value = ii
                ii = ii + 1
                objMySql.objDictionary.MoveNext
            Loop
        End With
        cmdPrint.Enabled = True
        Set objPrdBar = Nothing
       
    Else
        MsgBox "해당 조건의 데이타가 없습니다.", vbInformation + vbOKOnly, "채혈리스트조회"
    End If
     Me.MousePointer = 0
End Sub

Private Sub cmdWardList_Click()

'% 병동코드 리스트를 팝업한다.
    
    cboCol.Clear
    Call medClearTable(tblCollect)
    
    Dim objMyList  As New clsPopUpList
'    Dim objWard As clsBasisData
    
    With objMyList
        .FormCaption = "병동 조회"
        .Connection = DBConn
        .ColumnHeaderText = "병동코드;병동명"
        .Tag = "WardID"
        Me.ScaleMode = 1
'        Call .ListPop(, 3950, 6300, ObjLISComCode.WardID)
        Call .LoadPopUp(GetSQLWardList)  ', 3950, 6300)
        
        txtWardId.Text = medGetP(.SelectedString, 1, ";")
        
        If txtWardId.Text <> "" Then dtpColdt.SetFocus
        
    End With
    
    Set objMyList = Nothing
'    Set objWard = Nothing
End Sub

Private Sub Get_ColTmList()
    Dim strTmp     As String
    Dim aryColTm() As String
    Dim ii         As Integer
    
    If txtWardId = "" Then Exit Sub
    Me.MousePointer = 11
    With objMySql
        If P_ApplyBuildingInfo Then
            .BuildCd = ObjSysInfo.BuildingCd
        Else
            .BuildCd = "10"
        End If
        
        .WorkDt = Format(dtpColdt.Value, "yyyymmdd")
        .WardID = txtWardId
        strTmp = .Get_Coltm
    End With

    cboCol.Clear
    
    If strTmp <> "" Then
        aryColTm = Split(strTmp, COL_DIV)
        For ii = LBound(aryColTm) To UBound(aryColTm)
            cboCol.AddItem Mid(aryColTm(ii), 1, 2) & ":" & _
                           Mid(aryColTm(ii), 3, 2) & ":" & _
                           Mid(aryColTm(ii), 5) & Space(3) '& _
                           medGetP(aryColTm(ii), 2, Space(3))
        Next
        tblCollect.MaxRows = 0
        cboCol.ListIndex = 0
    Else
        MsgBox "해당일의 채혈리스트가 없습니다.", vbInformation + vbOKOnly, "채혈리스트"
    End If
    Me.MousePointer = 0
End Sub

Private Sub dtpColdt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub dtpColdt_LostFocus()

    Call Get_ColTmList
End Sub

Private Sub Form_Activate()
    txtWardId.SetFocus
End Sub

Private Sub Form_Load()
    Call FrmInitionalize
    cmdPrint.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objMySql = Nothing
End Sub

Private Sub txtWardId_GotFocus()
    txtWardId.SelStart = 0
    txtWardId.SelLength = Len(txtWardId)
    txtWardId.Tag = txtWardId
End Sub

Private Sub txtWardId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtWardId_LostFocus()
    If txtWardId = "" Then Exit Sub
    If txtWardId.Tag = txtWardId Then Exit Sub
    
    cboCol.Clear
    Call medClearTable(tblCollect)
    
'    With ObjLISComCode.WardID
'        If .Exists(txtWardId) Then
'            Call .KeyChange(txtWardId)
'            objsysinfo.BuildingCd = .Tags("bldgb")
'            objsysinfo.BuildingNm = .Tags("bldnm")
'            objsysinfo.BuildingNo = .Tags("bldno")
'            txtWardId.Tag = txtWardId
'        Else
'            MsgBox "병동 코드를 확인하세요..", vbInformation, "코드입력오류"
'            txtWardId.Text = ""
'        End If
'    End With
    
'    Dim objWard As clsBasisData
    Dim Rs As Recordset
    Dim strWard As String
    
'    Set objWard = New clsBasisData
    Set Rs = New Recordset
    
    strWard = GetSQLWard(txtWardId.Text)
    
    Rs.Open strWard, DBConn
    
    If Rs.EOF = False Then
        ObjSysInfo.BuildingCd = Rs.Fields("bldgb").Value & ""
        ObjSysInfo.BuildingNm = Rs.Fields("bldnm").Value & ""
        ObjSysInfo.BuildingNo = Rs.Fields("bldno").Value & ""
        txtWardId.Tag = txtWardId.Text
    Else
        MsgBox "병동 코드를 확인하세요.", vbInformation
        txtWardId.Text = ""
    End If
    Set Rs = Nothing
'    Set objWard = Nothing
        
End Sub
