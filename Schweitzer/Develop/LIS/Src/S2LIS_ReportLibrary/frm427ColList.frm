VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm427ColList 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11820
   Begin FPSpread.vaSpread tblExcel 
      Height          =   750
      Left            =   1200
      TabIndex        =   16
      Top             =   7845
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
      SpreadDesigner  =   "frm427ColList.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgExcel 
      Left            =   4815
      Top             =   7965
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "출   력 (&P)"
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CheckBox chkPreView 
      BackColor       =   &H00800000&
      Caption         =   "출력물 미리보기"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8805
      TabIndex        =   12
      Top             =   1140
      Value           =   1  '확인
      Width           =   1620
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   45
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   529
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
      Caption         =   "◈  검색 조건"
      LeftGab         =   100
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00EBF3ED&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   7
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
      TabIndex        =   6
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin Crystal.CrystalReport crtRpt 
      Left            =   4380
      Top             =   7995
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7005
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1410
      Width           =   10740
      _Version        =   196608
      _ExtentX        =   18944
      _ExtentY        =   12356
      _StockProps     =   64
      BackColorStyle  =   3
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
      MaxCols         =   8
      MaxRows         =   29
      ScrollBars      =   2
      ShadowColor     =   15463405
      ShadowDark      =   14737632
      SpreadDesigner  =   "frm427ColList.frx":01A9
      Appearance      =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   780
      Left            =   75
      TabIndex        =   0
      Top             =   285
      Width           =   10755
      Begin VB.ComboBox cboDept 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5130
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   285
         Width           =   1440
      End
      Begin VB.ComboBox cboWorkarea 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   285
         Width           =   2835
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "리스트 조회(&Q)"
         Height          =   480
         Left            =   8940
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   180
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         Height          =   315
         Left            =   7515
         TabIndex        =   2
         Top             =   285
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
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
         Format          =   62259203
         CurrentDate     =   36328
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "진료과 :"
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
         Left            =   4305
         TabIndex        =   14
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Work Area :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "채취일 :"
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
         Index           =   0
         Left            =   6675
         TabIndex        =   3
         Top             =   345
         Width           =   750
      End
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   300
      Left            =   75
      TabIndex        =   11
      Top             =   1080
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   529
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
      Caption         =   "◈  조회 리스트"
      LeftGab         =   100
   End
End
Attribute VB_Name = "frm427ColList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()

Private Sub cmdClear_Click()
    Call InitForm
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim i As Long
    Dim blnHeader As Boolean
    
    If tblList.DataRowCnt = 0 Then Exit Sub
    
    Call medClearTable(tblExcel)
    
    With tblList
        For i = 1 To .DataRowCnt
            .Col = 1
            .Row = i
            If .Value <> "1" Then
                .Col = 2: .Col2 = .MaxCols
                .Row = i: .Row2 = i
                .BlockMode = True
                strTmp = .Clip
                .BlockMode = False
                
                If blnHeader = False Then
                    Call tblExcel.SetText(1, 1, "접수 번호")
                    Call tblExcel.SetText(2, 1, "환 자 명")
                    Call tblExcel.SetText(3, 1, "환자 ID")
                    Call tblExcel.SetText(4, 1, "진 료 과")
                    Call tblExcel.SetText(5, 1, "채취 시간")
                    Call tblExcel.SetText(6, 1, "검사 항목")
                    Call tblExcel.SetText(7, 1, "검 체")
                    blnHeader = True
                End If
                
                tblExcel.Col = 1: tblExcel.Col2 = tblExcel.MaxCols
                tblExcel.Row = tblExcel.DataRowCnt + 1: tblExcel.Row2 = tblExcel.DataRowCnt + 1
                tblExcel.BlockMode = True
                tblExcel.Clip = strTmp
                tblExcel.BlockMode = False
            End If
        Next
    End With
    
    If tblExcel.DataRowCnt = 0 Then Exit Sub
    
    dlgExcel.InitDir = "C:\"
    dlgExcel.Filter = "ExCelFile(*.XLS)|*.XLS"
    dlgExcel.FileName = "ColList" & "_" & Format(dtpColdt.Value, "MMdd")
    dlgExcel.ShowSave

    Call tblExcel.SaveTabFile(dlgExcel.FileName)
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub cmdPrint_Click()
    Dim lngFNo As Long
    Dim strFNm As String
    Dim strRptNm As String
    Dim strTmp As String
    Dim i As Long
    Dim j As Long
    Dim varChk As Variant
    
    If tblList.DataRowCnt = 0 Then Exit Sub
    
    If Dir(InstallDir & "LIS\Rpt\OutColList.rpt") = "" Then
        MsgBox "출력도중 오류가 발생하였습니다. OutColList.rpt 파일이 없습니다.", vbExclamation
        Exit Sub
    Else
        strRptNm = InstallDir & "LIS\Rpt\OutColList.rpt"
    End If
    
    If Dir(InstallDir & "LIS\Rpt\CrystalReport.txt") = "" Then
        MsgBox "출력도중 오류가 발생하였습니다. CrystalReport.txt 파일이 없습니다.", vbExclamation
        Exit Sub
    Else
        strFNm = InstallDir & "LIS\Rpt\CrystalReport.txt"
    End If
    
    strTmp = ""
    With tblList
        For i = 1 To .DataRowCnt
            
            Call .GetText(1, i, varChk)
            If varChk <> "1" Then
                .Row = i
                .Col = 2: strTmp = strTmp & .Value & vbTab
                .Col = 3: strTmp = strTmp & .Value & vbTab
                .Col = 4: strTmp = strTmp & .Value & vbTab
                .Col = 5: strTmp = strTmp & .Value & vbTab
                .Col = 6: strTmp = strTmp & .Value & vbTab
                .Col = 7: strTmp = strTmp & .Value & vbTab
                .Col = 8: strTmp = strTmp & .Value & vbTab & vbNewLine
            End If
        Next
    End With
    
    If strTmp = "" Then
        MsgBox "출력할 리스트가 없습니다.", vbExclamation
        Exit Sub
    End If
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    
    lngFNo = FreeFile
    
    Open strFNm For Output As #lngFNo
    Print #lngFNo, strTmp
    Close #lngFNo
    With crtRpt
        .ReportFileName = strRptNm

        .ParameterFields(0) = "Workarea;" & Trim(medGetP(cboWorkarea.Text, 1, COL_DIV)) & ";true"
        .ParameterFields(1) = "ColDt;" & Format(dtpColdt.Value, "yyyy/MM/dd") & ";true"
        .ParameterFields(2) = "HospNm;" & ObjSysInfo.Hospital & ";true"
        .ParameterFields(3) = "Dept;" & Trim(medGetP(cboDept.Text, 1, COL_DIV)) & ";true"

        .RetrieveDataFiles
        .WindowState = 2 ' crptMaximized
        .Destination = IIf(chkPreView.Value = 1, crptToWindow, crptToPrinter)
        .Action = 1
        .Reset
    End With
End Sub

Private Sub cmdQuery_Click()
    Dim objAge As clsPatient
    Dim objPro As jProgressBar.clsProgress
    Dim Rs As Recordset
    Dim strSQL As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim strTmp As String
    Dim strAccNo As String
    Dim PreAccNo As String
    Dim CurAccNo As String
    Dim strTestNM As String
    Dim Row As Long
'rstdiv ='*' and detailfg<>''
'rstdiv <>'' and detailfg='' 만 검색.. 엄마랑, 단일항목만 보여주기 위해서
    
    If Trim(medGetP(cboWorkarea, 2, COL_DIV)) = "-1" Then
        strTmp = ""
    Else
        strTmp = " and " & DBW("a.workarea=", Trim(medGetP(cboWorkarea, 2, COL_DIV)))
    End If
    
    If Trim(medGetP(cboDept, 2, COL_DIV)) = "-1" Then
        strTmp = strTmp
    Else
        strTmp = strTmp & " and " & DBW("a.deptcd=", Trim(medGetP(cboDept, 2, COL_DIV)))
    End If
    
    Set objPro = New jProgressBar.clsProgress
    
    With objPro
        .Container = Me
        .Left = lblPrgBar.Left
        .Top = lblPrgBar.Top
        .Width = lblPrgBar.Width
        .Height = lblPrgBar.Height
        .Message = "자료를 검색하고 있습니다..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = lblPrgBar.Width
'        .XPos = lblPrgBar.Left
'        .YPos = lblPrgBar.Top
'        .YHeight = lblPrgBar.Height
'        .ForeColor = &H864B24
'        .Msg = " 자료를 검색하고 있습니다..."
'        .Value = 1
    End With

    strSQL1 = " select a.workarea,a.accdt,a.accseq,a.ptid,e." & F_PTNM & " as ptnm, e." & F_SSN & "  as ssn, " & F_DOB2("e") & " as dob, " & _
            " e." & F_SEX & " as " & F_SEX & ", e." & F_SSN & " as ssn, f." & F_DEPTCD & " as deptcd, f." & F_DEPTNM & " as deptnm, " & _
            " a.coldt,a.coltm, b.testcd,d.abbrnm5,a.spccd,c.field3 as spcnm, b.rstdiv as rstdiv, b.detailfg as detailfg" & _
            " from  " & T_HIS003 & " f, " & T_HIS001 & " e, " & T_LAB001 & " d, " & T_LAB032 & " c, " & T_LAB302 & " b, " & T_LAB201 & " a " & _
            " where " & DBW("a.coldt=", Format(dtpColdt.Value, "yyyyMMdd")) & _
            strTmp & _
            " and (a.wardid is null or a.wardid='') " & _
            " and a.workarea=b.workarea " & _
            " and a.accdt=b.accdt " & _
            " and a.accseq=b.accseq " & _
            " and " & DBW("c.cdindex=", LC3_Specimen) & _
            " and a.spccd=c.cdval1 " & _
            " and a.ptid=e." & F_PTID & _
            " and b.testcd= d.testcd " & _
            " and d.applydt=(select max(applydt) from " & T_LAB001 & " where testcd=d.testcd) " & _
            " and a.deptcd=f." & F_DEPTCD
    strSQL2 = " select a.workarea,a.accdt,a.accseq,a.ptid,e." & F_PTNM & " as ptnm, e." & F_SSN & "  as ssn, " & F_DOB2("e") & " as dob, " & _
            " e." & F_SEX & " as " & F_SEX & ", e." & F_SSN & " as ssn, f." & F_DEPTCD & " as deptcd, f." & F_DEPTNM & " as deptnm, " & _
            " a.coldt,a.coltm, b.testcd,d.abbrnm5,a.spccd,c.field3 as spcnm, b.rstdiv as rstdiv, b.detailfg as detailfg" & _
            " from  " & T_HIS003 & " f, " & T_HIS001 & " e, " & T_LAB001 & " d, " & T_LAB032 & " c, " & T_LAB404 & " b, " & T_LAB201 & " a " & _
            " where " & DBW("a.coldt=", Format(dtpColdt.Value, "yyyyMMdd")) & _
            strTmp & _
            " and (a.wardid is null or a.wardid='') " & _
            " and a.workarea=b.workarea " & _
            " and a.accdt=b.accdt " & _
            " and a.accseq=b.accseq " & _
            " and " & DBW("c.cdindex=", LC3_Specimen) & _
            " and a.spccd=c.cdval1 " & _
            " and a.ptid=e." & F_PTID & _
            " and b.testcd= d.testcd " & _
            " and d.applydt=(select max(applydt) from " & T_LAB001 & " where testcd=d.testcd) " & _
            " and a.deptcd=f." & F_DEPTCD
    strSQL3 = " select a.workarea,a.accdt,a.accseq,a.ptid,e." & F_PTNM & " as ptnm, e." & F_SSN & "  as ssn, " & F_DOB2("e") & " as dob, " & _
            " e." & F_SEX & " as " & F_SEX & ", e." & F_SSN & " as ssn, f." & F_DEPTCD & " as deptcd, f." & F_DEPTNM & " as deptnm, " & _
            " a.coldt,a.coltm, b.testcd,d.abbrnm5,a.spccd,c.field3 as spcnm, 'R' as rstdiv, '' as detailfg" & _
            " from  " & T_HIS003 & " f, " & T_HIS001 & " e, " & T_LAB001 & " d, " & T_LAB032 & " c, " & T_LAB351 & " b, " & T_LAB201 & " a " & _
            " where " & DBW("a.coldt=", Format(dtpColdt.Value, "yyyyMMdd")) & _
            strTmp & _
            " and (a.wardid is null or a.wardid='') " & _
            " and a.workarea=b.workarea " & _
            " and a.accdt=b.accdt " & _
            " and a.accseq=b.accseq " & _
            " and " & DBW("c.cdindex=", LC3_Specimen) & _
            " and a.spccd=c.cdval1 " & _
            " and a.ptid=e." & F_PTID & _
            " and b.testcd= d.testcd " & _
            " and d.applydt=(select max(applydt) from " & T_LAB001 & " where testcd=d.testcd) " & _
            " and a.deptcd=f." & F_DEPTCD
    
    strSQL = strSQL1 & " union " & strSQL2 & " union " & strSQL3 & " order by  workarea, accdt, accseq, coldt, coltm "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    Call medClearTable(tblList)
    tblList.MaxRows = 25
    
    If Rs.EOF Then
        MsgBox "검색된 자료가 존재하지 않습니다.", vbExclamation
        GoTo Nodata
    End If
    
    objPro.Max = Rs.RecordCount
    
    tblList.ReDraw = False
    
    Do Until Rs.EOF
        Dim Cnt As Long
        Cnt = Cnt + 1
        objPro.Value = Cnt
        
        CurAccNo = Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & "" & Rs.Fields("accseq").Value & ""
        
        If PreAccNo = CurAccNo Then ' 같은 접수번호일때
            If Rs.Fields("rstdiv").Value & "" = "*" And Rs.Fields("detailfg").Value & "" <> "" Then
                strTestNM = strTestNM & "," & Rs.Fields("abbrnm5").Value & ""
                tblList.Row = Row: tblList.Col = 7
                tblList.Value = strTestNM
            ElseIf Rs.Fields("rstdiv").Value & "" <> "" And Rs.Fields("detailfg").Value & "" = "" Then
                strTestNM = strTestNM & "," & Rs.Fields("abbrnm5").Value & ""
                tblList.Row = Row: tblList.Col = 7
                tblList.Value = strTestNM
            End If
        Else    '다른 접수번호
            strTestNM = Rs.Fields("abbrnm5").Value & ""
            Row = Row + 1
            
            With tblList
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
                Call .SetText(2, Row, Rs.Fields("workarea").Value & "" & "-" & Mid(Rs.Fields("accdt").Value & "", 3) & "-" & Rs.Fields("accseq").Value & "")
                Call .SetText(3, Row, Rs.Fields("ptnm").Value & "")
                Call .SetText(4, Row, Rs.Fields("ptid").Value & "")
                Call .SetText(5, Row, Rs.Fields("deptcd").Value & "")
                Call .SetText(6, Row, Format(Rs.Fields("coltm").Value & "", "@@:@@:@@"))
                Call .SetText(7, Row, strTestNM)
                Call .SetText(8, Row, Rs.Fields("spcnm").Value & "")
                
                'S2Con_hos\clsPatient\HIS001READ에 있는 내용 그대로 퍼왔음
                'SEX/AGE를 보여줘야 할때 사용하면 됨. 지금은 사용을 안하지만 나중을 위해 지우지 말것.
'                Dim SSN As String
'                Dim DOB As String
'                Dim SEX As String
'                Dim AGE As Integer
'                Dim SEXAGE  As String
'
'                SSN = Trim(rs.Fields("ssn").Value & "")
'                DOB = Trim(rs.Fields("dob").Value & "")
'                If Not IsDate(Format(DOB, "####-##-##")) Then
'                    If IsDate(Format(SSN, "####-##-##")) Then
'                        DOB = SSN
'                    Else
'                        If Len(Mid(DOB, 1, 4)) = 4 Then
'                            DOB = Mid(DOB, 1, 4) & "0101"
'                        ElseIf Len(Mid(SSN, 1, 4)) = 4 Then
'                            DOB = Mid(SSN, 1, 4) & "0101"
'                        Else
'                            DOB = Format(GetSystemDate, CS_DateDbFormat)
'                        End If
'                    End If
'                End If
'                SEX = Trim(rs.Fields("sex").Value & "")
'                If IsNumeric(SEX) Then
'                    SEX = Choose((Val(SEX) Mod 2) + 1, "F", "M")
'                End If
'                Set objAge = New clsPatient
'                Call objAge.GetAge(DOB, AGE, "Y")
'                Set objAge = Nothing
'                SEXAGE = SEX & "/" & AGE
'
'                Call .SetText(5, Row, SEXAGE)
            End With
        End If
        PreAccNo = Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & "" & Rs.Fields("accseq").Value & ""
        
        Rs.MoveNext
    Loop
    
    tblList.ReDraw = True
    
Nodata:
    Set objPro = Nothing
    Set Rs = Nothing
End Sub

Private Sub Form_Activate()
    Call LoadWorkarea
    Call LoadDept
End Sub

Private Sub LoadWorkarea()
    Dim objList As New clsLISHospital05
    Dim Rs As New Recordset
    
    Set Rs = objList.LoadWorkarea
    
    cboWorkarea.Clear
    
    cboWorkarea.AddItem Format("전 체", "!" & String(50, "@")) & COL_DIV & "-1"
    
    Do Until Rs.EOF
        cboWorkarea.AddItem Format(Rs.Fields("field1").Value & "", "!" & String(50, "@")) & Rs.Fields("cdval1").Value & ""
        
        Rs.MoveNext
    Loop
    
    cboWorkarea.ListIndex = 0
    
    Set Rs = Nothing
    Set objList = Nothing
End Sub

Private Sub LoadDept()
'    Dim objDept As New clsBasisData
    Dim Rs As New Recordset
    
    Rs.Open GetSQLDeptList, DBConn
    
    cboDept.Clear
    cboDept.AddItem Format("전 체", "!" & String(50, "@")) & COL_DIV & "-1"

    Do Until Rs.EOF
        cboDept.AddItem Format(Rs.Fields("deptnm").Value & "", "!" & String(50, "@")) & COL_DIV & Rs.Fields("deptcd").Value & ""

        Rs.MoveNext
    Loop

    cboDept.ListIndex = 0
    
    Set Rs = Nothing
'    Set objDept = Nothing
    
'    Dim objDept As clsDictionary
'
'    Set objDept = New clsDictionary
'
'    Set objDept = ObjLISComCode.DeptCd
'
'    cboDept.Clear
'
'    cboDept.AddItem Format("전 체", "!" & String(50, "@")) & COL_DIV & "-1"
'
'    objDept.MoveFirst
'    Do Until objDept.EOF
'        cboDept.AddItem Format(objDept.Fields("deptnm"), "!" & String(50, "@")) & COL_DIV & objDept.Fields("deptcd")
'
'        objDept.MoveNext
'    Loop
'
'    cboDept.ListIndex = 0
'
'    Set objDept = Nothing
End Sub

Private Sub Form_Load()
    dtpColdt.Value = GetSystemDate
    Call InitForm
End Sub

Private Sub InitForm()
    tblList.MaxRows = 25
    Call medClearTable(tblList)
End Sub
