VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm386OtPtCnt 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "환자별의뢰항목조회"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   8190
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   5145
      Top             =   3825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblExcel 
      Height          =   750
      Left            =   4680
      TabIndex        =   6
      Top             =   4635
      Visible         =   0   'False
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm386OtPtCnt.frx":0000
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   4665
      Top             =   3870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "엑셀받기(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8190
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   285
      Left            =   210
      TabIndex        =   2
      Top             =   255
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "조회조건"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   285
      Left            =   210
      TabIndex        =   3
      Top             =   1395
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "조회 리스트"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblData 
      Height          =   6390
      Left            =   210
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1695
      Width           =   10620
      _Version        =   196608
      _ExtentX        =   18732
      _ExtentY        =   11271
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
      MaxCols         =   11
      MaxRows         =   50
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   15463405
      ShadowDark      =   14737632
      SpreadDesigner  =   "frm386OtPtCnt.frx":01A9
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   930
      Left            =   210
      TabIndex        =   8
      Top             =   465
      Width           =   10650
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   6705
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   285
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
         Height          =   315
         Left            =   4365
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frm386OtPtCnt.frx":07D7
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtDeptCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1230
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   360
         Left            =   210
         TabIndex        =   12
         Top             =   330
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         BackColor       =   14411494
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
         Caption         =   "접수일"
      End
      Begin MSComCtl2.DTPicker dtpApplyDate 
         Height          =   330
         Left            =   1125
         TabIndex        =   13
         Top             =   345
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM"
         Format          =   59113475
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   345
         Left            =   4680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.Label lblCap 
         BackColor       =   &H00DBE6E6&
         Caption         =   "진료과"
         Height          =   225
         Index           =   0
         Left            =   2340
         TabIndex        =   15
         Top             =   405
         Width           =   3765
      End
   End
End
Attribute VB_Name = "frm386OtPtCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private objDic As New clsDictionary



Private Sub cmdExcel_Click()

    Dim strTmp As String
    
    If tblData.DataRowCnt = 0 Then Exit Sub
    
    With tblData
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblExcel.MaxRows = .MaxRows + 1
        tblExcel.MaxCols = .MaxCols
        tblExcel.Row = 1: tblExcel.Row2 = tblExcel.MaxRows
        tblExcel.Col = 1: tblExcel.Col2 = tblExcel.MaxCols
        tblExcel.BlockMode = True
        tblExcel.Clip = strTmp
        tblExcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "외부수탁의뢰건수"
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)

End Sub

Private Sub cmdExit_Click()
    
    Set objCodeList = Nothing
    Set objDic = Nothing
    
    Unload Me
    
End Sub

Private Sub cmdPopupList_Click()
    Dim tmpSql As String
    Dim lngTop As Long, lngLeft As Long
'    Dim objDept As clsBasisData

    
    Set objCodeList = New clsPopUpList
'    Set objDept = New clsBasisData
    
    With objCodeList
        lngTop = txtDeptCd.Top + 2350
        lngLeft = Me.Left + txtDeptCd.Left + 50
        .Connection = DBConn
        .FormCaption = "진료과 리스트"
        .ColumnHeaderText = "진료과;진료과명"
'        .ListPop , lngTop, lngLeft, ObjLISComCode.DeptCd
        .LoadPopUp GetSQLDept ', lngTop, lngLeft
        txtDeptCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
        lblDeptNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
    End With
    
'    Set objDept = Nothing
End Sub



Private Sub cmdPrint_Click()

    Dim strTmp As String
    Dim intFNum As Integer
    Dim strRfile As String
    Dim strRptPath As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    Me.MousePointer = 11
    With tblData
        For ii = 1 To .DataRowCnt
            .Row = ii
            For jj = 1 To .MaxCols
                .Col = jj
                 
                If jj >= 4 Then
                    If Len(.Value) <> 0 Then
                        .Row = 0
                        strTmp = strTmp & Trim(.Value) & ", "
                    End If
                    
                    .Row = ii
                Else
                    strTmp = strTmp & Trim(.Value) & vbTab
                End If
                
            Next jj
'            Debug.Print strTmp
            strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
'            Debug.Print strTmp
            
            strTmp = strTmp & vbCr

        Next ii
        
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        
    End With
    
    strRfile = InstallDir & "LIS\Rpt\CrystalReport.txt"
    strRptPath = InstallDir & "LIS\Rpt\LabOtPtCnt.rpt"
    
    
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    
    With CReport
        .ParameterFields(0) = "HosptalNm1;" & lblDeptNm.Caption & "수탁검사 의뢰 건수" & ";true"
        .ParameterFields(1) = "Date;" & Format(dtpApplyDate.Value, "YYYY-MM") & "월" & ";true"
        .ParameterFields(2) = "HosptalNm2;" & P_HOSPITALNAME & ";true"

        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With
    
    Me.MousePointer = 0
    
End Sub

Private Sub Form_Load()
    Dim SSQL As String
    Dim RS   As Recordset
    
    
    SSQL = "select cdval1,field1 from " & T_LAB032 & " where " & DBW("cdindex=", "C250")
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    objDic.Clear
    objDic.FieldInialize "Cd", "Nm"
    
    If Not RS.EOF Then
        Do Until RS.EOF
            If objDic.Exists(RS.Fields("cdval1").Value & "") Then
                objDic.KeyChange RS.Fields("cdval1").Value & ""
                objDic.Fields("nm") = RS.Fields("field1").Value & ""
            Else
                objDic.AddNew RS.Fields("cdval1").Value & "", RS.Fields("field1").Value & ""
            End If
            RS.MoveNext
        Loop
    End If
    
    dtpApplyDate.Value = Format(GetSystemDate, "yyyy-mm")
    
    Set RS = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim objDic As New clsDictionary
    Dim objPro As clsProgress
    Dim RS       As Recordset
    Dim SSQL     As String
    Dim strFdt   As String
    Dim strTdt   As String
    Dim kk       As Integer
    Dim strPtChk As String
    Dim ii       As Integer
    Dim jj As Integer
    
    Call medClearTable(tblData, False, False)
    
    strFdt = Format(dtpApplyDate.Value, "yyyymm") & "01"
    strTdt = Format(dtpApplyDate.Value, "yyyymm") & "31"
    Me.MousePointer = 11
    SSQL = GetSQLString(strFdt, strTdt, txtDeptCd.Text)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
        
    kk = 4
    objDic.Clear
    objDic.FieldInialize "testcd", "seq,testnm"
    
    objDic.Sort = False
    If Not RS.EOF Then
        Do Until RS.EOF
            If objDic.Exists(RS.Fields("testcd").Value & "") = False Then
                objDic.AddNew RS.Fields("testcd").Value & "", _
                              kk & COL_DIV & RS.Fields("abbrnm5").Value & ""
                kk = kk + 1
            End If
            RS.MoveNext
        Loop
        objDic.Sort = True
        objDic.MoveFirst
        RS.MoveFirst
        
        With tblData
            .ReDraw = True
            Set objPro = New clsProgress
'            Set objPro.StatusBar = mainfrm.stsbar
            objPro.Container = MainFrm.stsbar
            objPro.Max = RS.RecordCount
            
            .MaxCols = objDic.RecordCount + 3
            .Row = 0
            Do Until objDic.EOF
                .Col = objDic.Fields("seq"): .Value = objDic.Fields("testnm")
                objDic.MoveNext
            Loop
            RS.MoveFirst
            Do Until RS.EOF
                If strPtChk <> RS.Fields("ptid").Value & "" & RS.Fields("orddt").Value & "" Then
                    ii = ii + 1
                    If .DataRowCnt >= .MaxRows Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    .Col = 3: .Value = GetPtNm(RS.Fields("ptid").Value & "")
                End If
                .Col = 1: .Value = Format(Mid(RS.Fields("orddt").Value & "", 5), "0#-##"): .TypeHAlign = TypeHAlignRight
                .Col = 2: .Value = Format(Mid(RS.Fields("ptid").Value & "", 2), "########")
                
                objDic.KeyChange RS.Fields("testcd").Value & ""
                .Col = objDic.Fields("seq"): .Value = RS.Fields("cnt").Value & "": .TypeHAlign = TypeHAlignRight
                strPtChk = RS.Fields("ptid").Value & "" & RS.Fields("orddt").Value & ""
                RS.MoveNext
                
                jj = jj + 1
                objPro.Value = jj
            Loop
            
            Set objPro = Nothing
        End With
    End If
    Me.MousePointer = 0
End Sub


Private Function GetSQLString(ByVal FRcvDt As String, ByVal TRcvDt As String, _
                              ByVal DeptCd As String) As String

  Dim SSQL As String
    
    '일반검사
    SSQL = " SELECT" & _
           " b.orddt,a.ptid,b.testcd,d.abbrnm5, count(*) as Cnt  " & _
           " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
           " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
           " AND    " & DBW("a.vfydt <=", TRcvDt) & _
           " AND    " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
           " AND    " & DBW("a.deptcd=", DeptCd) & _
           " AND b.workarea = a.workarea" & _
           " AND b.accdt    = a.accdt" & _
           " AND b.accseq   = a.accseq" & _
           " AND d.testcd = b.testcd" & _
           " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
           " AND ( d.itemseq<>0)  " & _
           " GROUP BY  b.orddt,a.ptid,b.testcd,d.abbrnm5"
    
    '기타검사
    SSQL = SSQL & " UNION ALL" & _
           " SELECT " & _
           " b.orddt,a.ptid,b.testcd,d.abbrnm5, count(*) as Cnt  " & _
           " FROM " & T_LAB001 & " d," & T_LAB351 & " b," & T_LAB201 & " a" & _
           " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
           " AND " & DBW("a.vfydt <=", TRcvDt) & _
           " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
           " AND " & DBW("a.deptcd =", DeptCd) & _
           " AND b.workarea = a.workarea" & _
           " AND b.accdt    = a.accdt" & _
           " AND b.accseq   = a.accseq" & _
           " AND d.testcd = b.testcd" & _
           " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
           " AND ( d.itemseq<>0)  " & _
           " GROUP BY  b.orddt,a.ptid,b.testcd,d.abbrnm5"
    
    '미생물검사
    SSQL = SSQL & " UNION ALL" & _
           " SELECT " & _
           " b.orddt, a.ptid,b.testcd,d.abbrnm5, count(*) as Cnt  " & _
           " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
           " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
           " AND " & DBW("a.vfydt <=", TRcvDt) & _
           " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
           " AND " & DBW("a.deptcd =", DeptCd) & _
           " AND b.workarea = a.workarea" & _
           " AND b.accdt    = a.accdt" & _
           " AND b.accseq   = a.accseq" & _
           " AND d.testcd = b.testcd" & _
           " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
           " AND ( d.itemseq<>0)  " & _
           " GROUP BY  b.orddt,a.ptid,b.testcd,d.abbrnm5"

    SSQL = SSQL & " ORDER BY  orddt,ptid"

    GetSQLString = SSQL
End Function
'Private Function GetPtNm(ByVal Ptid As String) As String
'    Dim SSQL As String
'    Dim RS   As Recordset
'
'    SSQL = " SELECT " & F_PTNM & " as ptnm from " & T_HIS001 & " where " & DBW(F_PTID, Ptid, 2)
'
'    Set RS = New Recordset
'    RS.Open SSQL, DBConn
'
'    If Not RS.EOF Then
'        GetPtNm = RS.Fields("ptnm").Value & ""
'    End If
'    Set RS = Nothing
'End Function

Private Sub Form_Unload(Cancel As Integer)
    
    Set objCodeList = Nothing
    Set objDic = Nothing

End Sub
