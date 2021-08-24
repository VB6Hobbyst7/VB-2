VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm420RIAColList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   4
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
      TabIndex        =   3
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   330
      Index           =   1
      Left            =   75
      TabIndex        =   9
      Top             =   45
      Width           =   10710
      _ExtentX        =   18891
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
      Caption         =   "리스트 조회조건"
      LeftGab         =   100
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   10650
      Top             =   690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   330
      Index           =   0
      Left            =   75
      TabIndex        =   8
      Top             =   1515
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   6555
      Left            =   75
      ScaleHeight     =   6495
      ScaleWidth      =   10695
      TabIndex        =   6
      Top             =   1860
      Width           =   10755
      Begin FPSpread.vaSpread tblCollect 
         Height          =   6480
         Left            =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   10665
         _Version        =   196608
         _ExtentX        =   18812
         _ExtentY        =   11430
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
         MaxCols         =   9
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis420.frx":0000
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1200
      Left            =   75
      TabIndex        =   10
      Top             =   315
      Width           =   10725
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9270
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   405
         Width           =   1320
      End
      Begin VB.TextBox txtWork 
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
         Left            =   1410
         TabIndex        =   12
         Top             =   270
         Width           =   1665
      End
      Begin VB.CheckBox chkTestdiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사코드출력"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   390
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         Height          =   375
         Left            =   1410
         TabIndex        =   0
         Top             =   660
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
         Format          =   59310083
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpcoltm 
         Height          =   375
         Left            =   3105
         TabIndex        =   1
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
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
         CustomFormat    =   "HH:mm:ss"
         Format          =   59310083
         CurrentDate     =   36328.8820023148
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   225
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
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
         Caption         =   "Workarea"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   225
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   660
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
         Caption         =   "채 취 일"
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "이후"
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
         Index           =   1
         Left            =   4275
         TabIndex        =   13
         Top             =   765
         Width           =   405
      End
   End
End
Attribute VB_Name = "frm420RIAColList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()
Private Enum TblColumn
    tcSel = 1
    tcNO
    tcWORKNO
    tcPTNM
    tcPTID
    tcSA
    tcCOLDT
    tcTEST
    tcSPCNO
End Enum

Private Sub Clear()
    tblCollect.MaxRows = 0
    txtWork = RI_WORKAREA
    dtpColdt.Value = GetSystemDate
    dtpcoltm.Value = dtpColdt.Value
    chkTestdiv.Value = 0
End Sub

Private Sub cmdClear_Click()
    Call Clear
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Me.MousePointer = 11
    Call CollectList_Print
    Call Clear
    Me.MousePointer = 0
End Sub

Private Sub dtpColdt_Change()
    dtpcoltm.Value = dtpColdt.Value
End Sub

Private Sub Form_Load()
    Call Clear
End Sub

Private Sub cmdQuery_Click()
    Dim objProBar   As New jProgressBar.clsProgress
    Dim objSql      As clsWardColList
    Dim Rs          As Recordset
    Dim strPtId     As String
    Dim strcoldt    As String
    Dim strColtm    As String
    Dim bussdiv     As String
    Dim ii          As Integer
    
    
    strcoldt = Format(dtpColdt.Value, "yyyymmdd")
    strColtm = Format(dtpcoltm.Value, "hhmmss")
    
    Set objSql = New clsWardColList
    
    tblCollect.MaxRows = 0
    Set Rs = objSql.RI_CollectList(RI_WORKAREA, strcoldt, strColtm)
    
    If Rs.EOF Then
        MsgBox "해당조건의 리스트가 없습니다.", vbInformation + vbOKOnly, "리스트조회"
        GoTo Skip
    End If
    
    'Set objProBar.StatusBar = medMain.stsBar
'    objProBar.SetStsBar MAINFRM.STSBAR
    objProBar.Container = MainFrm.STSBAR
    objProBar.Min = 1
    objProBar.Max = Rs.RecordCount
    
    With tblCollect
        objSql.TestDiv = chkTestdiv.Value
        .ReDraw = False
        .MaxRows = Rs.RecordCount
        Do Until Rs.EOF
            ii = ii + 1
            .Row = ii
            .Col = TblColumn.tcNO:     .Value = ii
            .Col = TblColumn.tcWORKNO: .Value = Rs.Fields("workarea").Value & "" & "-" & _
                                                Rs.Fields("accdt").Value & "" & "-" & _
                                                Rs.Fields("accseq").Value & ""
            .Col = TblColumn.tcPTNM:   .Value = Rs.Fields("ptnm").Value & ""
            .Col = TblColumn.tcPTID:   .Value = Rs.Fields("ptid").Value & ""
            '주민번호 에러 발생난다.(sybase) 로드 후에...ptid 를 가지고 주민번호를 구해서
            
            
            .Col = TblColumn.tcSA:     .Value = Rs.Fields("sex").Value & "" & "/" & PtAge(Rs.Fields("ptid").Value & "")
            
            .Col = TblColumn.tcCOLDT:  .Value = Mid(Rs.Fields("coldt").Value & "", 3) & " " & _
                                                Mid(Rs.Fields("coltm").Value & "", 1, 4)
            '검사종목
            .Col = TblColumn.tcTEST:   .Value = objSql.RI_TESTLIST(Rs.Fields("workarea").Value & "", _
                                                                  Rs.Fields("accdt").Value & "", _
                                                                  Rs.Fields("accseq").Value & "")
            '검체
            '.Col = TblColumn.tcSPCNO:  .Value = Rs.Fields("spcnm").Value & ""
            bussdiv = Rs.Fields("bussdiv").Value
            If bussdiv = "2" Or bussdiv = "3" Then
                .Col = TblColumn.tcSPCNO
                .Value = Rs.Fields("wardid").Value & ""
                If Rs.Fields("hosilid").Value & "" <> "" Then
                    .Value = .Value & "-" & Rs.Fields("hosilid").Value & ""
                End If
            Else
                .Col = TblColumn.tcSPCNO: .Value = Rs.Fields("deptcd").Value & "" & "(외래)"
            End If
            objProBar.Value = ii
            Rs.MoveNext
        Loop
        .ReDraw = True
    End With
Skip:
    Set Rs = Nothing
    Set objProBar = Nothing
    Set objSql = Nothing
End Sub
Private Function PtAge(ByVal ssn As String) As String
    Dim strTmp As String
    Dim strSex As String
    Dim strAGE As String
    Dim strDOB As String
    
    Dim strYY  As String
    Dim strMM  As String
    Dim strDD  As String
    
    Dim objSql      As clsWardColList
    Dim strSSN       As String
    
    
    Set objSql = New clsWardColList
    
    strSSN = objSql.GetPtid_SSN(ssn)
    
    GoTo Nodata
    
    If strSSN = "" Then Exit Function
    
    strYY = Mid(ssn, 1, 2)
    strMM = Mid(ssn, 3, 2)
    strDD = Mid(ssn, 5, 2)
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    If IsDate(strYY & "-" & strMM & "-" & strDD) = False Then
        strDD = "01"
    End If
    
    strSex = "기타": strAGE = "": strDOB = ""
    
    If ssn <> "" Then
        strTmp = Mid(ssn, 7, 1)
        Select Case strTmp
            Case "0":  strSex = "여": strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1":  strSex = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2":  strSex = "여": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3":  strSex = "남": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4":  strSex = "여": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strSex = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        If Len(ssn) = 13 Then
            PtAge = medFindAge(Replace(strDOB, "-", ""), "Y")
        Else
            PtAge = ""
        End If
        'PtAGE = strSEX & COL_DIV & strDOB & COL_DIV & strAGE
    Else
        PtAge = ""
       ' PtAGE = "" & COL_DIV & "" & COL_DIV & ""
    End If
Nodata:
    Set objSql = Nothing
    
End Function

Private Sub CollectList_Print()
    Dim strTmp As String
    Dim intFNum As Integer
    Dim strRfile As String
    Dim strRptPath As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    Dim strWork  As String
    Dim strPtId  As String
    Dim strPtNm  As String
    Dim strSa    As String
    Dim strcoldt As String
    Dim strtest  As String
    Dim strspc   As String
   
    With tblCollect
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSel
            If .Value = 0 Then
                .Col = TblColumn.tcWORKNO: strWork = Mid(.Value, 4)
                .Col = TblColumn.tcPTNM:   strPtNm = .Value
                .Col = TblColumn.tcPTID:   strPtId = .Value
                .Col = TblColumn.tcSA:     strSa = .Value
                .Col = TblColumn.tcCOLDT:  strcoldt = Format(medGetP(.Value, 1, " "), "00-00-00") & " " & _
                                                      Format(medGetP(.Value, 2, " "), "0#:##")
                .Col = TblColumn.tcTEST:   strtest = .Value
                .Col = TblColumn.tcSPCNO:  strspc = .Value
                
                strTmp = strTmp & strWork & vbTab & strPtNm & vbTab & strPtId & vbTab & _
                                  strSa & vbTab & strcoldt & vbTab & strtest & vbTab & strspc & vbCr
                jj = jj + 1
            End If
        Next
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End With
    
    strRfile = InstallDir & "LIS\Rpt\CrystalReport.txt"
    strRptPath = InstallDir & "LIS\Rpt\RICollectionList.rpt"
    
    
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    '
    With CReport
        .ParameterFields(0) = "ActionDate;" & Format(dtpColdt.Value, "yyyy-mm-dd") & ";TRUE"
        .ParameterFields(1) = "ActionNm;" & ObjSysInfo.EmpNm & ";TRUE"
        .ParameterFields(2) = "SpcNm;" & "검체수 :  " & jj & ";TRUE"
        .ParameterFields(3) = "title;" & "RIA 외래채혈리스트" & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        
        .WindowState = 0
        .WindowTitle = "RI 채혈리스트"
        
        .Action = 1
        .Reset
    End With
End Sub
