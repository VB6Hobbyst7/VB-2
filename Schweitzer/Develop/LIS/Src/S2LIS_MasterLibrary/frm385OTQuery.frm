VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm385OTQuery 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "검사별 단가조회"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   8190
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00DBE6E6&
      Caption         =   "엑셀받기(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   8190
      Width           =   1320
   End
   Begin VB.OptionButton OptDiv 
      BackColor       =   &H00800000&
      Caption         =   "주치의"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   3765
      TabIndex        =   8
      Top             =   300
      Width           =   1215
   End
   Begin VB.OptionButton OptDiv 
      BackColor       =   &H00800000&
      Caption         =   "처방의"
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   2565
      TabIndex        =   7
      Top             =   285
      Width           =   1215
   End
   Begin VB.OptionButton OptDiv 
      BackColor       =   &H00800000&
      Caption         =   "진료과"
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   0
      Left            =   1335
      TabIndex        =   6
      Top             =   285
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8190
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   5490
      Top             =   7890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblExcel 
      Height          =   540
      Left            =   2490
      TabIndex        =   4
      Top             =   4995
      Visible         =   0   'False
      Width           =   2475
      _Version        =   196608
      _ExtentX        =   4366
      _ExtentY        =   952
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
      SpreadDesigner  =   "frm385OTQuery.frx":0000
   End
   Begin Crystal.CrystalReport crtReport 
      Left            =   5025
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9525
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8190
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   285
      Left            =   195
      TabIndex        =   0
      Top             =   285
      Width           =   10500
      _ExtentX        =   18521
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
      Left            =   195
      TabIndex        =   1
      Top             =   1425
      Width           =   10500
      _ExtentX        =   18521
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
      Height          =   6375
      Left            =   195
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1725
      Width           =   10515
      _Version        =   196608
      _ExtentX        =   18547
      _ExtentY        =   11245
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
      MaxCols         =   7
      MaxRows         =   50
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   15463405
      ShadowDark      =   14737632
      SpreadDesigner  =   "frm385OTQuery.frx":01A9
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   930
      Left            =   195
      TabIndex        =   11
      Top             =   495
      Width           =   10500
      Begin VB.TextBox txtDeptCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   1230
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
         Picture         =   "frm385OTQuery.frx":07BE
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   6705
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   285
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   360
         Left            =   210
         TabIndex        =   15
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
         TabIndex        =   16
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
         Format          =   62586883
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   345
         Left            =   4680
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   405
         Width           =   3765
      End
   End
End
Attribute VB_Name = "frm385OTQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private objDic As New clsDictionary
Private objCnt As New clsDictionary

Private Sub cmdClear_Click()
    dtpApplyDate.Value = Format(GetSystemDate, "yyyy-mm")
    txtDeptCd.Text = "": lblDeptNm.Caption = ""
    tblData.MaxRows = 0
    
End Sub

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
    
    DlgSave.InitDir = "C:\My Documents"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "외부수탁청구조회"
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)

End Sub

Private Sub cmdExit_Click()

    Set objCodeList = Nothing
    Set objDic = Nothing
    Set objCnt = Nothing
    Unload Me

End Sub

Private Sub cmdPopupList_Click()
    Dim tmpSql As String
    Dim lngTop As Long, lngLeft As Long
'    Dim objData As clsBasisData

    
    Set objCodeList = New clsPopUpList
'    Set objData = New clsBasisData
    
    objCodeList.Connection = DBConn
    
    If OptDiv(2).Value = True Then
        With objCodeList
            lngTop = txtDeptCd.Top + 2350
            lngLeft = Me.Left + txtDeptCd.Left + 50
            .FormCaption = "주치의 리스트"
            .ColumnHeaderText = "주치의;주치의명"
'            Call .ListPop(GetDoctListSQL, lngTop, lngLeft)
            Call .LoadPopUp(GetSQLDoctList) ', lngTop, lngLeft)
            txtDeptCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
            lblDeptNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
        End With
    ElseIf OptDiv(0).Value = True Then
        With objCodeList
            lngTop = txtDeptCd.Top + 2350
            lngLeft = Me.Left + txtDeptCd.Left + 50
            .FormCaption = "진료과 리스트"
            .ColumnHeaderText = "진료과;진료과명"
'            .ListPop , lngTop, lngLeft, ObjLISComCode.DeptCd
            .LoadPopUp GetSQLDeptList ', lngTop, lngLeft
            txtDeptCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
            lblDeptNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
            
        End With
    Else
        With objCodeList
            lngTop = txtDeptCd.Top + 2350
            lngLeft = Me.Left + txtDeptCd.Left + 50
            .FormCaption = "처방의 리스트"
            .ColumnHeaderText = "처방의;처방의명"
'            Call .ListPop(GetDoctListSQL, lngTop, lngLeft)
            Call .LoadPopUp(GetSQLDoctList) ', lngTop, lngLeft)
            txtDeptCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
            lblDeptNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
        End With
    End If
End Sub

Private Sub cmdPrint_Click()
    
    Dim RS As Recordset
    Dim strTmp As String
    Dim strRfile As String
    Dim strRptPath As String
    Dim strBank As String
    Dim intFNum As Integer
    Dim SQL As String
    Dim i As Integer
    Dim j As Integer
    Dim dblTotal As Double
    
    If tblData.DataRowCnt = 0 Then
        MsgBox "출력할 내역이 없습니다.확인하여 주세요.", vbCritical, "출력오류"
        Exit Sub
    End If
    
    SQL = "SELECT text2 FROM " & T_LAB032 & " WHERE cdval1='" & txtDeptCd.Text & "'"
    Set RS = New Recordset
    RS.Open SQL, DBConn
    
    If Not RS.EOF Then
        strBank = RS.Fields("text2").Value & ""
    End If
    
    Set RS = Nothing
    
    With tblData
        For i = 1 To .DataRowCnt
            .Row = i
            For j = 1 To 7
                .Col = j
                strTmp = strTmp & .Value & vbTab
                
                If j = 6 Then
                    dblTotal = dblTotal + Val(Replace(.Value, ",", ""))
                End If
                
           Next j
            
            strTmp = Mid(strTmp, 1, Len(strTmp) - 1) & vbCr
        Next i
        
        strTmp = strTmp & vbTab & "합 계" & vbTab & vbTab & vbTab & vbTab & _
                     Format(dblTotal, "#,###") & vbTab & vbTab
    End With
    
    strRfile = InstallDir & "Lis\Rpt\CrystalReport.txt"
    strRptPath = InstallDir & "Lis\Rpt\LABOTCnt.rpt"
    intFNum = FreeFile
    
On Error GoTo ErrPrint
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum

    With crtReport
        .ParameterFields(0) = "HosptalNm1;" & lblDeptNm.Caption & " 귀하" & ";true"
        .ParameterFields(1) = "Date;" & "(" & Left(dtpApplyDate.Value, 7) & ")" & ";true"
        .ParameterFields(2) = "HosptalNm2;" & P_HOSPITALNAME & ";true"
        .ParameterFields(3) = "Bank;" & strBank & ";true"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
        
    End With
    
    Exit Sub
        
ErrPrint:
    MsgBox "출력이 되지 않았습니다.", vbCritical

End Sub

Private Sub cmdQuery_Click()
    Dim RS       As Recordset
    Dim pRs      As Recordset
    Dim objPro   As clsProgress
    Dim objRs    As clsProgress
    
    Dim SSQL     As String
    Dim strFdt   As String
    Dim strTdt   As String
    Dim strTmp   As String
    Dim strtmp1  As String
    Dim strOptNm As String
    Dim strNm    As String
    Dim strOptCd As String
    Dim ii       As Long
    
'    If txtDeptCd.Text = "" Then Exit Sub
    
    'Call medClearTable(tblData, False, False)
    tblData.MaxRows = 0
    Me.MousePointer = 11
    objCnt.Clear
    objCnt.Sort = False
    If Trim(txtDeptCd.Text) = "" Then
        objCnt.FieldInialize "optcd,testcd", "optnm,testnm,cnt,pay,text1"
    Else
        objCnt.FieldInialize "testcd", "testnm,cnt,pay,text1"
    End If
    
    strFdt = Format(dtpApplyDate.Value, "yyyymm") & "01"
    strTdt = Format(dtpApplyDate.Value, "yyyymm") & "31"
    
    Set RS = New Recordset
    RS.Open GetSQLString(strFdt, strTdt, txtDeptCd.Text), DBConn
    If Not RS.EOF Then
        Set objRs = New clsProgress
        objRs.Container = MainFrm.stsbar
'        Set objRs.StatusBar = mainfrm.stsbar
        objRs.Max = RS.RecordCount
        
        If Trim(txtDeptCd.Text) = "" Then
            Do Until RS.EOF
                If OptDiv(1).Value = True Then
                    strOptCd = RS.Fields("orddoct").Value & ""
                ElseIf OptDiv(0).Value = True Then
                    strOptCd = RS.Fields("deptcd").Value & ""
                Else
                    strOptCd = RS.Fields("orddoct").Value & ""
                End If
                
                If objCnt.Exists(strOptCd & COL_DIV & RS.Fields("testcd").Value & "") Then
                    objCnt.KeyChange strOptCd & COL_DIV & RS.Fields("testcd").Value & ""
                    objCnt.Fields("cnt") = Val(objCnt.Fields("cnt")) + Val(RS.Fields("cnt").Value & "")
                Else
                    If objDic.Exists(RS.Fields("testcd").Value & "") Then
                        objDic.KeyChange RS.Fields("testcd").Value & ""
                        strTmp = objDic.Fields("val")
                        strtmp1 = objDic.Fields("rmk")
                    Else
                        strTmp = "0": strtmp1 = ""
                    End If

                    If OptDiv(0).Value = True Then
'                        Dim objDept As clsBasisData
                        
'                        Set objDept = New clsBasisData
                        
                        strOptNm = GetDeptNm(RS.Fields("deptcd").Value & "")
                        
                        If strOptNm = "" Then
                            strOptNm = RS.Fields("deptcd").Value & ""
                        End If
'                        Set objDept = Nothing
                        
'                        If ObjLISComCode.DeptCd.Exists(Rs.Fields("deptcd").Value & "") Then
'                            ObjLISComCode.DeptCd.KeyChange Rs.Fields("deptcd").Value & ""
'                            strOptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                        Else
'                            strOptNm = Rs.Fields("deptcd").Value & ""
'                        End If
                    Else
'                        Dim objDoct As clsBasisData
                        
'                        Set objDoct = New clsBasisData
                        
                        strOptNm = GetDoctNm(RS.Fields("orddoct").Value & "")
                        
'                        strOptNm = getempname(Rs.Fields("orddoct").Value & "")
                        If strOptNm = "" Then
                            strOptNm = RS.Fields("orddoct").Value & ""
                        End If
                        
'                        Set objDoct = Nothing
                    End If
                    
                    objCnt.AddNew strOptCd & COL_DIV & RS.Fields("testcd").Value & "", strOptNm & COL_DIV & RS.Fields("testnm").Value & "" & COL_DIV & _
                                                                  RS.Fields("cnt").Value & "" & COL_DIV & _
                                                                  strTmp & COL_DIV & strtmp1
                    ii = ii + 1
                    objRs.Value = ii
    
    '                objRS.Msg = "데이터를 집계하고 있습니다. (" & ii & " 건)"
                End If
                RS.MoveNext
            Loop
        Else
            
            Do Until RS.EOF
                If objCnt.Exists(RS.Fields("testcd").Value & "") Then
                    objCnt.KeyChange RS.Fields("testcd").Value & ""
                    objCnt.Fields("cnt") = Val(objCnt.Fields("cnt")) + Val(RS.Fields("cnt").Value & "")
                Else
                
                    If objDic.Exists(RS.Fields("testcd").Value & "") Then
                        objDic.KeyChange RS.Fields("testcd").Value & ""
                        strTmp = objDic.Fields("val")
                        strtmp1 = objDic.Fields("rmk")
                    Else
                        strTmp = "0": strtmp1 = ""
                    End If
               
'                    SSQL = " select field1,text2 from " & T_LAB032 & _
'                         " where " & _
'                         DBW("cdindex=", "C249") & " and " & DBW("cdval1=", RS.Fields("testcd").Value)
'                    Set pRs = OpenRecordSet(SSQL)
'                    If Not pRs.EOF Then
'                        strTmp = pRs.Fields("field1").Value & ""
'                        strtmp1 = pRs.Fields("text2").Value & ""
'                    Else
'                        strTmp = ""
'                        strtmp1 = ""
'                    End If
                    
                    objCnt.AddNew RS.Fields("testcd").Value & "", RS.Fields("testnm").Value & "" & COL_DIV & _
                                                                  RS.Fields("cnt").Value & "" & COL_DIV & _
                                                                  strTmp & COL_DIV & strtmp1
                    ii = ii + 1
                    objRs.Value = ii
                
    '                objRS.Msg = "데이터를 집계하고 있습니다. (" & ii & " 건)"
                End If
                RS.MoveNext
            Loop
           
        End If
        Set objRs = Nothing
    Else
        MsgBox "해당 데이타가 없습니다.", vbInformation + vbOKOnly, "검사건수조회"
        Me.MousePointer = 0
        Set RS = Nothing
        Exit Sub
    End If
    objCnt.Sort = True
    
    If objCnt.RecordCount < 1 Then
        Me.MousePointer = 0
        Set RS = Nothing
        Exit Sub
    End If
    
    Dim objTot As New clsDictionary
    
    objTot.Clear
    objTot.FieldInialize "cd,nm", "tot"
    
    objCnt.MoveFirst
    With tblData
        Set objPro = New clsProgress
'        Set objPro.StatusBar = mainfrm.stsbar
        objPro.Container = MainFrm.stsbar
        objPro.Max = objCnt.RecordCount

        ii = 1
        .MaxRows = objCnt.RecordCount
        strNm = ""
        If Trim(txtDeptCd.Text) = "" Then
            Do Until objCnt.EOF
                .Row = ii
                If strNm <> objCnt.Fields("optnm") Then
                    .Col = 1: .Value = objCnt.Fields("optnm")
                    objTot.AddNew objCnt.Fields("optcd") & COL_DIV & objCnt.Fields("optnm"), "0"
                End If

                strNm = objCnt.Fields("optnm")
                .Col = 2: .Value = objCnt.Fields("testcd")
                .Col = 3: .Value = objCnt.Fields("testnm"): .ForeColor = DCM_Blue
                .Col = 4: .Value = Format(objCnt.Fields("cnt"), "#,###")
                
                objTot.KeyChange objCnt.Fields("optcd") & COL_DIV & objCnt.Fields("optnm")
                
                If objCnt.Fields("pay") = "" Then
                    .Col = 5: .Value = "0"
                    .Col = 6: .Value = "0"
                Else
                    .Col = 5: .Value = Format(Val(objCnt.Fields("pay")), "#,###")
                    .Col = 6: .Value = Format(Val(objCnt.Fields("pay")) * Val(objCnt.Fields("cnt")), "#,###")
                End If
                
                objTot.Fields("tot") = Val(objTot.Fields("tot")) + Val(Replace(.Value, ",", ""))
                
'                Debug.Print .Value & "        ,   " & objTot.Fields("tot") & "    " & objTot.Fields("nm")
                
                .Col = 7: .Value = objCnt.Fields("text1")
                
                ii = ii + 1
                objPro.Value = ii
                objCnt.MoveNext
            Loop
        
        Else
            Do Until objCnt.EOF
                If objTot.Exists(txtDeptCd.Text & COL_DIV & lblDeptNm.Caption) = False Then
                    objTot.AddNew txtDeptCd.Text & COL_DIV & lblDeptNm.Caption, "0"
                End If
                objTot.KeyChange txtDeptCd.Text & COL_DIV & lblDeptNm.Caption
                
                .Row = ii
                .Col = 1: .Value = ii
                .Col = 2: .Value = objCnt.Fields("testcd")
                .Col = 3: .Value = objCnt.Fields("testnm"): .ForeColor = DCM_Blue
                .Col = 4: .Value = Format(objCnt.Fields("cnt"), "#,###")
'                .Col = 5: .Value = Format(objCnt.Fields("pay"), "#,###")
'                .Col = 6: .Value = CDbl(objCnt.Fields("pay")) * objCnt.Fields("cnt")
                
                If objCnt.Fields("pay") = "" Then
                    .Col = 5: .Value = "0"
                    .Col = 6: .Value = "0"
                Else
                    .Col = 5: .Value = Format(objCnt.Fields("pay"), "#,###")
                    .Col = 6: .Value = Format(Val(objCnt.Fields("pay")) * Val(objCnt.Fields("cnt")), "#,###")
                End If
                
                
                objTot.Fields("tot") = Val(objTot.Fields("tot")) + Val(Replace(.Value, ",", ""))
                
                
                .Col = 7: .Value = objCnt.Fields("text1")
                
                ii = ii + 1
                objPro.Value = ii
                objCnt.MoveNext
            Loop
        End If
        
        .MaxRows = .MaxRows + objTot.RecordCount + 1
        objTot.MoveFirst
        Dim dblTot As Double
        
        .Row = .DataRowCnt + 2
        Do Until objTot.EOF
            
            .Col = 1: .Value = objTot.Fields("nm")
            .Col = 6: .Value = Format(objTot.Fields("tot"), "#,###")
            dblTot = dblTot + Val(Replace(.Value, ",", ""))
            .Row = .DataRowCnt + 1
            
            objTot.MoveNext
        Loop
        
        .MaxRows = .MaxRows + 2
        .Row = .MaxRows
        .Col = 1: .Value = "합 계": .FontBold = True
        .Col = 6: .Value = Format(dblTot, "#,###"): .FontBold = True
        
        Set objTot = Nothing
        
                
'        For ii = 1 To .MaxRows
'            .Row = ii: .Col = 6
'            dblTot = dblTot + .Value
'            .Value = Format(.Value, "#,###")
'        Next
'        .MaxRows = .MaxRows + 2
'        .Row = .MaxRows
'        .Col = 1: .Value = "합 계"
'        .Col = 6: .Value = Format(dblTot, "#,###")
    
        Set objPro = Nothing
    End With
    
    Set RS = Nothing
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim SSQL As String
    Dim RS   As Recordset

    objDic.Clear
    objDic.FieldInialize "testcd", "val,rmk"
    objDic.Sort = False
    
    txtDeptCd.Locked = False
    SSQL = " select cdval1,field1,text2 from " & T_LAB032 & _
           " where " & DBW("cdindex=", "C249")
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            If objDic.Exists(RS.Fields("cdval1").Value & "") Then
                objDic.KeyChange RS.Fields("cdval1").Value & ""
                objDic.Fields("val") = RS.Fields("field1").Value & ""
                objDic.Fields("rmk") = RS.Fields("text2").Value & ""
            Else
                objDic.AddNew RS.Fields("cdval1").Value & "", RS.Fields("field1").Value & "" & COL_DIV & RS.Fields("text2").Value & ""
            End If
            RS.MoveNext
        Loop
    End If
    
    
'    SSQL = "select cdval1,field1 from " & T_LAB032 & " where " & DBW("cdindex=", "C250")
'    Set RS = OpenRecordSet(SSQL)
'
'    objDic.Clear
'    objDic.FieldInialize "Cd", "Nm"
'
'    If Not RS.EOF Then
'        Do Until RS.EOF
'            If objDic.Exists(RS.Fields("cdval1").Value & "") Then
'                objDic.KeyChange RS.Fields("cdval1").Value & ""
'                objDic.Fields("nm") = RS.Fields("field1").Value & ""
'            Else
'                objDic.AddNew RS.Fields("cdval1").Value & "", RS.Fields("field1").Value & ""
'            End If
'            RS.MoveNext
'        Loop
'    End If
'    RS.RsClose
    
    dtpApplyDate.Value = Format(GetSystemDate, "yyyy-mm")
    
    Set RS = Nothing
End Sub

Private Function GetSQLString(ByVal FRcvDt As String, ByVal TRcvDt As String, _
                              ByVal DeptCd As String) As String
    Dim SSQL As String
    
    If OptDiv(1).Value = True Then
        If Trim(DeptCd) = "" Then
            '처방의 전체
            SSQL = " SELECT" & _
               " a.orddoct,b.testcd, d.testnm, count(*) as Cnt  " & _
               " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
               " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
               " AND    " & DBW("a.vfydt <=", TRcvDt) & _
               " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               " AND b.workarea = a.workarea" & _
               " AND b.accdt    = a.accdt" & _
               " AND b.accseq   = a.accseq" & _
               " AND d.testcd = b.testcd" & _
               " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
               " AND (d.itemseq<>0) " & _
               " GROUP BY a.orddoct,b.testcd, d.testnm"
        
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " a.orddoct,b.testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB351 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.testcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  a.orddoct,b.testcd, d.testnm "
            
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " a.orddoct,b.ordcd as testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB102 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_MidRst) & _
                   " AND " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.ordcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  a.orddoct,b.ordcd, d.testnm "
            
            
'            SSQL = SSQL & " UNION ALL" & _
'                   " SELECT " & _
'                   " a.majdoct,b.testcd, d.testnm, count(*) as Cnt" & _
'                   " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
'                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
'                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
'                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
'                   " AND b.workarea = a.workarea" & _
'                   " AND b.accdt    = a.accdt" & _
'                   " AND b.accseq   = a.accseq" & _
'                   " AND d.testcd = b.testcd" & _
'                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
'                   " AND (d.itemseq<>0) " & _
'                   " GROUP BY  a.majdoct,b.testcd, d.testnm "
    
    
            SSQL = SSQL & " ORDER BY  orddoct,testcd, testnm"
            
            
        
        Else
            '처방의 별
            SSQL = " SELECT" & _
               " b.testcd, d.testnm, count(*) as Cnt  " & _
               " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
               " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
               " AND    " & DBW("a.vfydt <=", TRcvDt) & _
               " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               " AND    " & DBW("a.orddoct=", DeptCd) & _
               " AND b.workarea = a.workarea" & _
               " AND b.accdt    = a.accdt" & _
               " AND b.accseq   = a.accseq" & _
               " AND d.testcd = b.testcd" & _
               " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
               " AND (d.itemseq<>0) " & _
               " GROUP BY b.testcd, d.testnm"
        
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " b.testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB351 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                   " AND " & DBW("a.orddoct =", DeptCd) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.testcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  b.testcd, d.testnm "
                   
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " b.ordcd as testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB102 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_MidRst) & _
                   " AND " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                   " AND " & DBW("a.orddoct =", DeptCd) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.ordcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  b.ordcd, d.testnm "
                   
            
'            SSQL = SSQL & " UNION ALL" & _
'                   " SELECT " & _
'                   " b.testcd, d.testnm, count(*) as Cnt" & _
'                   " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
'                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
'                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
'                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
'                   " AND " & DBW("a.majdoct =", DeptCd) & _
'                   " AND b.workarea = a.workarea" & _
'                   " AND b.accdt    = a.accdt" & _
'                   " AND b.accseq   = a.accseq" & _
'                   " AND d.testcd = b.testcd" & _
'                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
'                   " AND (d.itemseq<>0) " & _
'                   " GROUP BY  b.testcd, d.testnm "
    
            SSQL = SSQL & " ORDER BY  testcd, testnm"
        
        End If
    ElseIf OptDiv(0).Value = True Then
        If Trim(DeptCd) = "" Then
            '부서 전체
            SSQL = " SELECT" & _
               " a.deptcd,b.testcd, d.testnm, count(*) as Cnt  " & _
               " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
               " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
               " AND    " & DBW("a.vfydt <=", TRcvDt) & _
               " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               " AND b.workarea = a.workarea" & _
               " AND b.accdt    = a.accdt" & _
               " AND b.accseq   = a.accseq" & _
               " AND d.testcd = b.testcd" & _
               " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
               " AND (d.itemseq<>0) " & _
               " GROUP BY a.deptcd,b.testcd, d.testnm"
        
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " a.deptcd,b.testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB351 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.testcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  a.deptcd,b.testcd, d.testnm "
            
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " a.deptcd,b.ordcd as testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB102 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_MidRst) & _
                   " AND " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.ordcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  a.deptcd,b.ordcd, d.testnm "
            
            
'            SSQL = SSQL & " UNION ALL" & _
'                   " SELECT " & _
'                   " a.deptcd,b.testcd, d.testnm, count(*) as Cnt" & _
'                   " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
'                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
'                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
'                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
'                   " AND b.workarea = a.workarea" & _
'                   " AND b.accdt    = a.accdt" & _
'                   " AND b.accseq   = a.accseq" & _
'                   " AND d.testcd = b.testcd" & _
'                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
'                   " AND (d.itemseq<>0) " & _
'                   " GROUP BY  a.deptcd,b.testcd, d.testnm "
    
            SSQL = SSQL & " ORDER BY  deptcd, testcd, testnm"

        Else
            '부서별
            SSQL = " SELECT" & _
                   " b.testcd, d.testnm, count(*) as Cnt  " & _
                   " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
                   " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND    " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                   " AND    " & DBW("a.deptcd=", DeptCd) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.testcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY b.testcd, d.testnm"
            
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " b.testcd, d.testnm, count(*) as Cnt" & _
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
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  b.testcd, d.testnm "


            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " b.ordcd as testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB102 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_MidRst) & _
                   " AND " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                   " AND " & DBW("a.deptcd =", DeptCd) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.ordcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  b.ordcd, d.testnm "


'            SSQL = SSQL & " UNION ALL" & _
'                   " SELECT " & _
'                   " b.testcd, d.testnm, count(*) as Cnt" & _
'                   " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
'                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
'                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
'                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
'                   " AND " & DBW("a.deptcd =", DeptCd) & _
'                   " AND b.workarea = a.workarea" & _
'                   " AND b.accdt    = a.accdt" & _
'                   " AND b.accseq   = a.accseq" & _
'                   " AND d.testcd = b.testcd" & _
'                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
'                   " AND (d.itemseq<>0) " & _
'                   " GROUP BY  b.testcd, d.testnm "
        
        
            SSQL = SSQL & " ORDER BY  testcd, testnm"
        End If
    Else
        If Trim(DeptCd) = "" Then
            '주치의 전체
            SSQL = " SELECT" & _
               " a.majdoct as orddoct,b.testcd, d.testnm, count(*) as Cnt  " & _
               " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
               " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
               " AND    " & DBW("a.vfydt <=", TRcvDt) & _
               " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               " AND b.workarea = a.workarea" & _
               " AND b.accdt    = a.accdt" & _
               " AND b.accseq   = a.accseq" & _
               " AND d.testcd = b.testcd" & _
               " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
               " AND (d.itemseq<>0) " & _
               " GROUP BY a.majdoct,b.testcd, d.testnm"
            
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " a.majdoct as orddoct,b.testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB351 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.testcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  a.majdoct,b.testcd, d.testnm "
                
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " a.majdoct as orddoct,b.ordcd as testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB102 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_MidRst) & _
                   " AND " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.ordcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  a.majdoct,b.ordcd, d.testnm "
            
                
'            SSQL = SSQL & " UNION ALL" & _
'                   " SELECT " & _
'                   " a.majdoct,b.testcd, d.testnm, count(*) as Cnt" & _
'                   " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
'                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
'                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
'                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
'                   " AND b.workarea = a.workarea" & _
'                   " AND b.accdt    = a.accdt" & _
'                   " AND b.accseq   = a.accseq" & _
'                   " AND d.testcd = b.testcd" & _
'                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
'                   " AND (d.itemseq<>0) " & _
'                   " GROUP BY  a.majdoct,b.testcd, d.testnm "
        
        
            SSQL = SSQL & " ORDER BY  orddoct,testcd, testnm"
        
        Else
            '주치의 별
            SSQL = " SELECT" & _
               " b.testcd, d.testnm, count(*) as Cnt  " & _
               " FROM " & T_LAB001 & " d," & T_LAB302 & " b," & T_LAB201 & " a" & _
               " WHERE  " & DBW("a.vfydt >=", FRcvDt) & _
               " AND    " & DBW("a.vfydt <=", TRcvDt) & _
               " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               " AND    " & DBW("a.majdoct=", DeptCd) & _
               " AND b.workarea = a.workarea" & _
               " AND b.accdt    = a.accdt" & _
               " AND b.accseq   = a.accseq" & _
               " AND d.testcd = b.testcd" & _
               " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
               " AND (d.itemseq<>0) " & _
               " GROUP BY b.testcd, d.testnm"
            
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " b.testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB351 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
                   " AND " & DBW("a.majdoct =", DeptCd) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.testcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  b.testcd, d.testnm "
                       
            SSQL = SSQL & " UNION ALL" & _
                   " SELECT " & _
                   " b.ordcd as testcd, d.testnm, count(*) as Cnt" & _
                   " FROM " & T_LAB001 & " d," & T_LAB102 & " b," & T_LAB201 & " a" & _
                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_MidRst) & _
                   " AND " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                   " AND " & DBW("a.majdoct =", DeptCd) & _
                   " AND b.workarea = a.workarea" & _
                   " AND b.accdt    = a.accdt" & _
                   " AND b.accseq   = a.accseq" & _
                   " AND d.testcd = b.ordcd" & _
                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
                   " AND (d.itemseq<>0) " & _
                   " GROUP BY  b.ordcd, d.testnm "
                       
                
    '            SSQL = SSQL & " UNION ALL" & _
    '                   " SELECT " & _
    '                   " b.testcd, d.testnm, count(*) as Cnt" & _
    '                   " FROM " & T_LAB001 & " d," & T_LAB404 & " b," & T_LAB201 & " a" & _
    '                   " WHERE " & DBW("a.vfydt >=", FRcvDt) & _
    '                   " AND " & DBW("a.vfydt <=", TRcvDt) & _
    '                   " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
    '                   " AND " & DBW("a.majdoct =", DeptCd) & _
    '                   " AND b.workarea = a.workarea" & _
    '                   " AND b.accdt    = a.accdt" & _
    '                   " AND b.accseq   = a.accseq" & _
    '                   " AND d.testcd = b.testcd" & _
    '                   " AND d.applydt = (select max(applydt) from " & T_LAB001 & " where  testcd = d.testcd)" & _
    '                   " AND (d.itemseq<>0) " & _
    '                   " GROUP BY  b.testcd, d.testnm "
        
            SSQL = SSQL & " ORDER BY  testcd, testnm"
        
        
        End If
    End If

    GetSQLString = SSQL

End Function

Private Sub Form_Unload(Cancel As Integer)
    
    Set objCodeList = Nothing
    Set objDic = Nothing
    Set objCnt = Nothing

End Sub

Private Sub optDiv_Click(Index As Integer)
    tblData.MaxRows = 0: tblData.MaxRows = 25
    txtDeptCd.Text = "": lblDeptNm.Caption = ""
End Sub


Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
'    Dim objDept As clsBasisData
    Dim strTmp As String
    
    tblData.MaxRows = 0
    lblDeptNm.Caption = ""
    If KeyAscii = vbKeyReturn Then
        If OptDiv(0).Value = True Then
'            Set objDept = New clsBasisData
            
            strTmp = GetDeptNm(Trim(txtDeptCd.Text))
            
            If strTmp = "" Then
                txtDeptCd.Text = ""
                lblDeptNm.Caption = ""
            Else
                lblDeptNm.Caption = strTmp
            End If
'            Set objDept = Nothing
            
'            If ObjLISComCode.DeptCd.Exists(UCase(Trim(txtDeptCd.Text))) Then
'                ObjLISComCode.DeptCd.KeyChange Trim(txtDeptCd.Text)
'                txtDeptCd.Text = UCase(txtDeptCd.Text)
'                lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'            Else
'                txtDeptCd.Text = "": lblDeptNm.Caption = ""
'            End If
        Else
'            Set objDept = New clsBasisData
            
            lblDeptNm.Caption = GetDoctNm(txtDeptCd.Text)
'            lblDeptNm.Caption = getempname(txtDeptCd.Text)
            If lblDeptNm.Caption = "" Then txtDeptCd.Text = ""
'            Set objDept = Nothing
        End If
    End If
    If lblDeptNm.Caption <> "" Then cmdQuery.SetFocus
End Sub
