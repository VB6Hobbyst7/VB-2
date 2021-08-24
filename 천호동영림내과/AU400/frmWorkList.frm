VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "워크리스트 조회"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14490
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      TabIndex        =   11
      Top             =   810
      Width           =   225
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   9
      Text            =   "0001"
      Top             =   210
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   8
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownClose 
      Caption         =   "Down >> Close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "Down"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   150
      Width           =   1395
   End
   Begin FPSpread.vaSpread vasWorkList 
      Height          =   6195
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   14055
      _Version        =   393216
      _ExtentX        =   24791
      _ExtentY        =   10927
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   20
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":000C
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   180
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   21430273
      CurrentDate     =   40457
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   315
      Left            =   3030
      TabIndex        =   4
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   21430273
      CurrentDate     =   40457
   End
   Begin VB.Label Label2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6360
      TabIndex        =   10
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   6
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label7 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2820
      TabIndex        =   5
      Top             =   270
      Width           =   105
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 0
        Next iRow
    End If
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDownClose_Click()
    
    Call cmdDownLoad_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdDownLoad_Click()
    Dim intVasRow As Integer
    Dim intRow As Integer
    Dim j  As Integer
    
    j = 0
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
                intVasRow = frmInterface.vasID.MaxRows
                
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colSpecNo), intVasRow, colSpecNo)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colCheckBox), intVasRow, colCheckBox)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colHospDate), intVasRow, colHospDate)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colGubun), intVasRow, colGubun)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colBarcode), intVasRow, colBarcode)
                'Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colRack), intVasRow, colRack)
                'Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colPos), intVasRow, colPos)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colPID - 2), intVasRow, colPID)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colPName - 2), intVasRow, colPName)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colSex - 2), intVasRow, colSex)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colAge - 2), intVasRow, colAge)
                
                frmInterface.txtNum = frmInterface.txtNum + 1
                
                .Col = 1
                .Value = "0"
            End If
        Next
        frmInterface.vasID.RowHeight(-1) = 12
    End With



'    Dim i As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        For i = 1 To vasWorkList.MaxRows
'            vasWorkList.Row = i
'            vasWorkList.Col = 1
'            If vasWorkList.Value = "1" Then
'                If Trim(txtPos.Text) = "" Then
'                    txtPos.Text = "1"
'                End If
'                Call SetText(frmInterface.vasID, Format(txtPos.Text, "0000"), i, 0)
'                txtPos.Text = Format(txtPos.Text + 1, "0000")
'            End If
'        Next
'    End If
End Sub

Private Sub cmdSearch_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim i, X As Long
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim FilNum
    Dim TxtString As String
    Dim TxtRece As String
    Dim PChartNum As String
    Dim PNAME As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAGE As String
    Dim PSEX As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    Dim PExamname As String
    Dim PEquipCode As String
    Dim pEqipType  As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim TxtPat As String
    Dim TestNum, IOGubun As String
    Dim FindFile As String
    Dim StartDate As String
    Dim EndDate As String
    Dim varXML      As Variant
    Dim varTmp      As Variant
    Dim strBarNo As String
    Dim intCnt As Integer
    Dim pGrid_Point As Integer
    Dim sList As Integer
    Dim strBarNum As String
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim RSX  As ADODB.Recordset
    Dim strItems As String
    
    Screen.MousePointer = 11
    DoEvents
    
    ClearSpread vasWorkList
    strItems = ""
    
    vasWorkList.ReDraw = False
    
    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    If blnSameRecord = False Then
        'MsgBox "검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        
              SQL = "select distinct commdate,chartno,patname,patsex,patage from pat_res "
        SQL = SQL & " where commdate between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStopDt.Value, "yyyymmdd") & "'"
        SQL = SQL & "   and (result = '' or result is null)"
        
        Set RSX = cn.Execute(SQL)
        Do Until RSX.EOF
            With vasWorkList
                pGrid_Point = SeqSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
    
                If pGrid_Point = 0 Then
                    pGrid_Point = SeqNullSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
                    If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                    .RowHeight(-1) = 12
                End If
                
                If strItems = "" Then
                    strItems = Trim(gReadBuf(0))
                Else
                    strItems = strItems & "/" & Trim(gReadBuf(0))
                End If
                
                .SetText 0, pGrid_Point, txtSeq.Text
                .SetText 1, pGrid_Point, "1"
                .SetText 2, pGrid_Point, Format(Trim(RSX.Fields("COMMDATE")), "####-##-##")
                .SetText 3, pGrid_Point, gPart
                strBarNum = Mid(Format(Trim(RSX.Fields("COMMDATE")), "########"), 5, 4) & Format(Trim(RSX.Fields("CHARTNO")), "0000000000")
                .SetText 4, pGrid_Point, strBarNum
                .SetText 5, pGrid_Point, Trim(RSX.Fields("CHARTNO"))
                .SetText 6, pGrid_Point, Trim(RSX.Fields("PATNAME"))
                .SetText 7, pGrid_Point, Trim(RSX.Fields("PATSEX"))
                .SetText 8, pGrid_Point, Trim(RSX.Fields("PATAGE"))
                .SetText 9, pGrid_Point, "Order"
                
                .SetText 10, pGrid_Point, strItems

            End With
            txtSeq.Text = Format(txtSeq.Text + 1, "0000")
            RSX.MoveNext
        Loop
        RSX.Close
        Exit Sub
    End If
    
    If UBound(varXML) < 1 Then
        'MsgBox "검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
              SQL = "select * from pat_res "
        SQL = SQL & " where commdate between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStartDt.Value, "yyyymmdd") & "'"
        
        Res = db_select_Col(gLocal, SQL)
        If Res > 0 Then
            PEquipno = gReadBuf(0)
            PEquipCode = gReadBuf(1)
            PExamname = gReadBuf(2)
        End If
        vasWorkList.ReDraw = True
        Exit Sub
    Else
        strBarNo = ""

        With vasWorkList
            '.Visible = False
            For intCnt = 0 To UBound(varXML) - 1
                varTmp = Split(varXML(intCnt), ",")
                                
                '-- 장비채널값찾기
                SQL = ""
                SQL = SQL & " SELECT EQUIPCODE,EXAMNAME "
                SQL = SQL & "   FROM EQUIPEXAM"
                SQL = SQL & "  WHERE EXAMCODE = '" & Trim(varTmp(8)) & "' "
                
                Res = GetDBSelectColumn(gLocal, SQL)
                XMLInData.ComExamID = ""
                
                '-- 오더 있을 경우
                If Res > 0 Then
                    XMLInData.ComExamID = Trim(gReadBuf(0))
                    If strItems = "" Then
                        strItems = Trim(gReadBuf(1))
                    Else
                        strItems = strItems & "/" & Trim(gReadBuf(1))
                    End If
                    
                    XMLInData.Company = varTmp(0)
                    XMLInData.HospCode = varTmp(1)
                    XMLInData.ChartNo = varTmp(2)
                    XMLInData.PatName = varTmp(3)
                    XMLInData.PatJumin = varTmp(4)
                    XMLInData.PatNo = varTmp(5)
                    XMLInData.CommDate = varTmp(6)
                    XMLInData.ExamNo = varTmp(7)
                    XMLInData.ExamID = varTmp(8)
                    'XMLInData.ComExamID = varTmp(9)
                    XMLInData.Specimen = varTmp(10)
                    XMLInData.Result = varTmp(11)
                    XMLInData.Reference = varTmp(12)
                    XMLInData.Remark = varTmp(13)
                    XMLInData.RsltDate = varTmp(14)
                    XMLInData.IOFlag = varTmp(15)
                    
                    SQL = "select equipno, equipcode, examname, examtype from equipexam where examcode = '" & XMLInData.ExamID & "' "
                    Res = db_select_Col(gLocal, SQL)
    '                Debug.Print XMLInData.ExamID
                    If Res > 0 Then
                        PEquipno = gReadBuf(0)
                        PEquipCode = gReadBuf(1)
                        PExamname = gReadBuf(2)
                                        
                        If strBarNo <> XMLInData.ChartNo Or pEqipType <> gReadBuf(3) Then
                            pEqipType = gReadBuf(3)
                            
                            pGrid_Point = SeqSearch_New(vasWorkList, XMLInData.ChartNo, pEqipType, 5)
        
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(vasWorkList, XMLInData.ChartNo, 5)
                                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                                .RowHeight(-1) = 12
                            End If
                            
                            .SetText 0, pGrid_Point, txtSeq.Text
                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, Format(XMLInData.CommDate, "####-##-##")
                            .SetText 3, pGrid_Point, pEqipType
                            strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")
                            'strBarNum = Format$(XMLInData.ChartNo, String$(SPCLEN, "#"))
                            
                            .SetText 4, pGrid_Point, strBarNum
                            .SetText 5, pGrid_Point, XMLInData.ChartNo
                            .SetText 6, pGrid_Point, XMLInData.PatName
                                        PJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                                        Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
                            .SetText 7, pGrid_Point, gPatGen.Sex
                            .SetText 8, pGrid_Point, gPatGen.Age
                            .SetText 9, pGrid_Point, "Order"
                            .SetText 10, pGrid_Point, strItems
                            
                            txtSeq.Text = Format(txtSeq.Text + 1, "0000")
                            strItems = ""
                        Else
                            .SetText 10, pGrid_Point, strItems
                        End If
                              SQL = "Select ChartNo from pat_res "
                        SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                        SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                        SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                        SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                        SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                        Res = db_select_Col(gLocal, SQL)
                        
                        If Res = 0 Then
                                  SQL = " insert into pat_res("
                            SQL = SQL & "Company,HospCode,ChartNo, "
                            SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
                            SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
                            SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
                            SQL = SQL & " values ("
                            SQL = SQL & "'" & XMLInData.Company & "',"
                            SQL = SQL & "'" & XMLInData.HospCode & "',"
                            SQL = SQL & "'" & XMLInData.ChartNo & "',"
                            SQL = SQL & "'" & XMLInData.PatName & "',"
                            SQL = SQL & "'" & gPatGen.Sex & "',"
                            SQL = SQL & "'" & gPatGen.Age & "',"
                            SQL = SQL & "'" & XMLInData.PatJumin & "',"
                            SQL = SQL & "'" & XMLInData.PatNo & "',"
                            SQL = SQL & "'" & XMLInData.CommDate & "',"
                            SQL = SQL & "'" & XMLInData.ExamNo & "',"
                            SQL = SQL & "'" & XMLInData.ExamID & "',"
                            SQL = SQL & "'" & XMLInData.ComExamID & "',"
                            SQL = SQL & "'" & XMLInData.Specimen & "',"
                            SQL = SQL & "'" & XMLInData.Result & "',"
                            SQL = SQL & "'" & XMLInData.Reference & "',"
                            SQL = SQL & "'" & XMLInData.Remark & "',"
                            SQL = SQL & "'" & XMLInData.RsltDate & "',"
                            SQL = SQL & "'" & XMLInData.IOFlag & "',"
                            SQL = SQL & "'" & strBarNum & "',"
                            SQL = SQL & "'" & pEqipType & "')"
                            
                            Res = SendQuery(gLocal, SQL)
                            
                            If Res = -1 Then
                                SaveQuery SQL
                            End If
                        
                        '-- 속도향상을 위해 쿼리문 지우기
                        Else
                                  SQL = " Update pat_res Set "
                            SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
                            SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
                            'SQL = SQL & " ExamNo = '" & XMLInData.ExamNo & "', "
                            'SQL = SQL & " PatNo = '" & XMLInData.PatNo & "',"
                            SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                            SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                            SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                            SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                            SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                            
                            Res = SendQuery(gLocal, SQL)
                        End If
                        
                        strBarNo = XMLInData.ChartNo
                    End If
                Else
                    'XMLInData.ComExamID = ""
                End If
                XMLInData.ComExamID = ""
            Next
            
        End With
    End If
    
    If vasWorkList.MaxRows > 0 Then
        cmdDownLoad.Enabled = True
        cmdDownClose.Enabled = True
    Else
        cmdDownLoad.Enabled = False
        cmdDownClose.Enabled = False
    End If
        
    vasWorkList.ReDraw = True
    
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    dtpStartDt.Value = Now
    dtpStopDt.Value = Now
    txtSeq.Text = "0001"
    
    vasWorkList.MaxRows = 0
    
End Sub



Private Sub txtPos_KeyPress(KeyAscii As Integer)
'    Dim i As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        For i = 1 To vasWorkList.MaxRows
'            vasWorkList.Row = i
'            vasWorkList.Col = 1
'            If vasWorkList.Value = "1" Then
'                If Trim(txtPos.Text) = "" Then
'                    txtPos.Text = "1"
'                End If
'                Call SetText(frmInterface.vasID, Format(txtPos.Text, "0000"), i, 0)
'                txtPos.Text = Format(txtPos.Text + 1, "0000")
'            End If
'        Next
'    End If
    
End Sub


Private Sub txtSeq_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtSeq.Text = Format(txtSeq.Text, "0000")
    End If

End Sub

Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim pGrid_Point As Integer
    Dim sBarcode As String
    Dim sChartNo As String
    
    If Row = 0 Then Exit Sub
    
    With vasWorkList
        '.Col = Col
        '.Row = Row
        '.Col = colBarcode
        pGrid_Point = SeqSearch(frmInterface.vasID, GetText(vasWorkList, Row, colBarcode), colBarcode)

        If pGrid_Point = 0 Then
            pGrid_Point = SeqNullSearch(frmInterface.vasID, Trim(.Text), colBarcode)
            If pGrid_Point = 0 Then
                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
                pGrid_Point = frmInterface.vasID.MaxRows
            End If
            .RowHeight(-1) = 12
        End If
        
'        .Row = Row: .Col = colBarcode
'        sBarcode = Trim(.Text)
        
        
'        Call frmInterface.vasID.SetText(colSpecNo, pGrid_Point, "1")
'        Call frmInterface.vasID.SetText(1, pGrid_Point, "1")
'        Call frmInterface.vasID.SetText(4, pGrid_Point, .Text)

        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSpecNo), pGrid_Point, colSpecNo)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCheckBox), pGrid_Point, colCheckBox)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colHospDate), pGrid_Point, colHospDate)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colGubun), pGrid_Point, colGubun)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colBarcode), pGrid_Point, colBarcode)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPID - 2), pGrid_Point, colPID)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPName - 2), pGrid_Point, colPName)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSex - 2), pGrid_Point, colSex)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colAge - 2), pGrid_Point, colAge)

'        .Row = Row: .Col = 5
'        Call vasID.SetText(5, pGrid_Point, .Text)
'        .Row = Row: .Col = 6
'        Call vasID.SetText(6, pGrid_Point, .Text)
'        .Row = Row: .Col = 7
'        Call vasID.SetText(7, pGrid_Point, .Text)
'        .Row = Row: .Col = 8
'        Call vasID.SetText(8, pGrid_Point, .Text)
        frmInterface.vasID.RowHeight(-1) = 12
    
'''        '바코드번호로 환자정보 불러오기
'''              SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,COMMDATE,EXAMNO,EXAMID,IOFLAG  "
'''        SQL = SQL & vbCrLf & "  FROM PAT_RES "
'''        SQL = SQL & vbCrLf & " WHERE EXAMTYPE = '" & gPart & "' "
'''        SQL = SQL & vbCrLf & "   AND BARCODE = '" & sBarcode & "'"
'''
'''
'''        Res = GetDBSelectColumn(gLocal, SQL)
'''
'''        If Res = 1 Then
'''            SetText frmInterface.vasID, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
'''            SetText frmInterface.vasID, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
'''            SetText frmInterface.vasID, Trim(gReadBuf(1)), pGrid_Point, colPName  '6
'''            SetText frmInterface.vasID, Trim(gReadBuf(2)), pGrid_Point, colSex    '7
'''            SetText frmInterface.vasID, Trim(gReadBuf(3)), pGrid_Point, colAge    '8
'''            SetText frmInterface.vasID, Format(Trim(gReadBuf(8)), "####-##-##"), pGrid_Point, 2
'''
'''            SetText frmInterface.vasID, Trim(gReadBuf(4)), pGrid_Point, 12
'''            SetText frmInterface.vasID, Trim(gReadBuf(5)), pGrid_Point, 13
'''            SetText frmInterface.vasID, Trim(gReadBuf(6)), pGrid_Point, 14
'''            SetText frmInterface.vasID, Trim(gReadBuf(7)), pGrid_Point, 15
'''            SetText frmInterface.vasID, Trim(gReadBuf(8)), pGrid_Point, 16
'''            SetText frmInterface.vasID, Trim(gReadBuf(9)), pGrid_Point, 17
'''            SetText frmInterface.vasID, Trim(gReadBuf(10)), pGrid_Point, 18
'''            SetText frmInterface.vasID, Trim(gReadBuf(11)), pGrid_Point, 19
'''            frmInterface.vasID.RowHeight(-1) = 12
'''        End If
    
    End With
End Sub
