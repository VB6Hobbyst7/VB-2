VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm515CVR 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '����
   Caption         =   "���ġ ������"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14955
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���(&P)"
      Height          =   510
      Left            =   11640
      Style           =   1  '�׷���
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   570
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '���
      BackColor       =   &H00DBE6E6&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   75
      ScaleHeight     =   930
      ScaleWidth      =   14355
      TabIndex        =   3
      Top             =   345
      Width           =   14385
      Begin VB.CommandButton cmdNewQuary 
         BackColor       =   &H00DBE6E6&
         Caption         =   "NEW �� ȸ(&Q)"
         Height          =   510
         Left            =   8880
         Style           =   1  '�׷���
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   210
         Width           =   1320
      End
      Begin VB.Frame fraDt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ȸ �Ⱓ"
         Height          =   705
         Left            =   90
         TabIndex        =   12
         Top             =   120
         Width           =   1485
         Begin MSComCtl2.DTPicker dtpFromDt 
            Height          =   315
            Left            =   180
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   86507520
            CurrentDate     =   36342.5951388889
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   2460
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   86507520
            CurrentDate     =   36342.5951388889
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "����"
            Height          =   180
            Left            =   1830
            TabIndex        =   16
            Tag             =   "15104"
            Top             =   300
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "����"
            Height          =   180
            Left            =   4110
            TabIndex        =   15
            Tag             =   "15104"
            Top             =   330
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame fraWa 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Work Area"
         Height          =   705
         Left            =   1590
         TabIndex        =   10
         Top             =   120
         Width           =   2415
         Begin VB.ComboBox cboWA 
            Height          =   300
            Left            =   120
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�˻��׸�"
         Height          =   705
         Left            =   4020
         TabIndex        =   6
         Top             =   120
         Width           =   4725
         Begin VB.CommandButton cmdHelpList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2025
            MaskColor       =   &H00F4F0F2&
            MousePointer    =   14  'ȭ��ǥ�� ����ǥ
            Style           =   1  '�׷���
            TabIndex        =   8
            Tag             =   "DeptCd"
            Top             =   270
            Width           =   285
         End
         Begin VB.TextBox txtTestCd 
            Height          =   315
            Left            =   135
            TabIndex        =   7
            Top             =   285
            Width           =   1875
         End
         Begin MedControls1.LisLabel lblTestNm 
            Height          =   330
            Left            =   2370
            TabIndex        =   9
            Top             =   285
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            BackColor       =   15988984
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
         End
      End
      Begin VB.CommandButton cmdQuary 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ȸ(&Q)"
         Height          =   510
         Left            =   10200
         Style           =   1  '�׷���
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Excel(&E)"
         Height          =   510
         Left            =   12900
         Style           =   1  '�׷���
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   210
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   285
      Left            =   75
      TabIndex        =   17
      Top             =   45
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "���ġ ������"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   7110
      Left            =   60
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1350
      Width           =   14415
      _Version        =   196608
      _ExtentX        =   25426
      _ExtentY        =   12541
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   8
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15463405
      MaxCols         =   17
      MaxRows         =   1
      ScrollBars      =   0
      ShadowColor     =   14411494
      SpreadDesigner  =   "Lis515.frx":0000
      UserResize      =   0
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   6090
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   750
      _Version        =   196608
      _ExtentX        =   1323
      _ExtentY        =   1191
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
      SpreadDesigner  =   "Lis515.frx":07DF
   End
End
Attribute VB_Name = "frm515CVR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' ���� �Ӽ��� ������ �����ؾ� �մϴ�.
'
' BorderStyle : 0 - ����
' MdiChild    : False
' WindowState : 0 - ǥ��
' Top         : 0
' Left        : 0
'
Public Event FormClose()
Public Event LastFormUnload()

Private Const FAddCol = 1


'����Ʈ �˾�
Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
Private WithEvents fL401 As S2LIS_ReviewLib.clsLisReviewForm
Attribute fL401.VB_VarHelpID = -1

Private objSQL  As New clsLISSqlStatistic
Private objIcdList  As clsDictionary
Private objRstCd    As clsDictionary

Private aryResultText() As String

Private blnCHkLoad As Boolean

Dim CaseStudy_TestCd As String

Private strWork01 As String
Private strWork02 As String
Private strWork03 As String
Private strWork04 As String
Private strWork05 As String
Private strWork06 As String
Private strWork07 As String
Private strWork08 As String
Private strWork15 As String
Private strWorkOT As String

Private Sub chkIndex_Click()
    
    txtTblClear
End Sub

Private Sub chkShow_Click()
    txtTblClear
End Sub

Private Function PrintOut() As Boolean
'    Dim strTmp      As String
'    Dim strFileNm   As String
'    Dim strRptNm    As String
'    Dim strMyFile   As String
'    Dim strTemp     As String
'    Dim strOption   As String
'    Dim lngFNum     As Long
'    Dim lngCnt      As Long
'    Dim i           As Long
'    Dim j           As Long
'
'
'    strMyFile = Dir(APSAppPath & "\..\rpt\CrystalReport.txt")
'
'    If strMyFile = "" Then
'        PrintOut = True
'        MsgBox "CrystalReport.txt ������ �����ϴ�.", vbCritical, "����Ȯ��"
'        Exit Function
'    End If
'    strMyFile = ""
'
'    strFileNm = APSAppPath & "\..\rpt\CrystalReport.txt"
'
'    strMyFile = Dir(APSAppPath & "\..\rpt\rptAPS021.rpt")
'
'    If strMyFile = "" Then
'        PrintOut = True
'        MsgBox "rptAPS021.rpt ������ �����ϴ�.", vbCritical, "����Ȯ��"
'        Exit Function
'    End If
'
'    strRptNm = APSAppPath & "\..\rpt\rptAPS021.rpt"
'
'    With tblIndex
'        For i = 1 To .DataRowCnt '.MaxRows
'            .Row = i
'            For j = 1 To 8
'                .Col = j
'                strTmp = strTmp & .Value & vbTab
'                lngCnt = lngCnt + 1
'            Next
'
'            If (lngCnt Mod 8) = 0 Then
'                strTmp = strTmp & vbCr
'            End If
'        Next
'    End With
'
'    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
'
'    Debug.Print strTmp
'
'    lngFNum = FreeFile
'
'On Error GoTo ErrPrint
'
'    Open strFileNm For Output As #lngFNum
'    Print #lngFNum, strTmp
'    Close #lngFNum
'    With crtReport
'        .ReportFileName = strRptNm
'        .ParameterFields(0) = "hostnm;" & AC5_HOSPITAL_DEPT_NAME & ";true"
''        .ParameterFields(0) = "HostNm;" & objSysInfo.Hospital & ";true"
'        .RetrieveDataFiles
'        .WindowState = 2 ' crptMaximized
'        .Destination = crptToWindow
'        .Action = 1
'        .Reset
'    End With
'    PrintOut = True
'    Exit Function
'
'ErrPrint:
'    PrintOut = False
End Function

Private Sub cboWA_Click()
    Call TxtClear
    Call txtTblClear
    If cboWA.ListIndex <> -1 Then
        If cboWA.Text <> CaseStudy_TestCd Then
            CaseStudy_TestCd = cboWA.Text
            txtTestCd.Text = ""
            lblTestNm.Caption = ""
        End If
    End If
End Sub


Private Sub cmdExcel_Click()

    Dim strTmp  As String
    
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    With tblResult
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
    DlgSave.FileName = "CVR ������"
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdHelpList_Click()
    Dim objTestDiv As New clsDictionary
    Dim objRs As Recordset
    
    If cboWA.ListIndex = -1 Then Exit Sub
    
    Set objListPop = New clsPopUpList
    
    Call TxtClear
    Call txtTblClear
    
    With objTestDiv
        .Clear
        .FieldInialize "�˻��׸��ڵ�", "�˻��,����"
        Set objRs = New Recordset
        objRs.Open objSQL.GetWAvsTest(medGetP(cboWA.Text, 1, " ")), DBConn
        While Not objRs.EOF
            .AddNew objRs.Fields("testcd").Value & "", objRs.Fields("abbrnm10").Value & COL_DIV & objRs.Fields("testdiv").Value
            objRs.MoveNext
        Wend
    End With
    Set objRs = Nothing
    
    With objListPop
        .Connection = DBConn
        .FormCaption = "�˻��׸� ��ȸ"
        .ColumnHeaderText = "�˻��׸��ڵ�;�˻��;����"
        .ColumnHeaderWidth = "1440;1260.284;750.0473"
        .FormWidth = 3900
        .LoadPopUp objSQL.GetWAvsTest(medGetP(cboWA.Text, 1, " "))
        txtTestCd.Text = medGetP(.SelectedString, 1, ";")
        lblTestNm.Caption = medGetP(.SelectedString, 2, ";")
        Call GetRstCdList
    End With
    Set objListPop = Nothing
End Sub

Private Sub cmdNewQuary_Click()
    Dim objProgress  As jProgressBar.clsProgress
    Dim RS           As New Recordset
    Dim RS1          As New Recordset
    Dim objPatient   As New clsPatient      'ȯ�� Ŭ����
    Dim SSQL         As String
    Dim strRstCdSql  As String
    Dim strDeptCd    As String
    Dim I            As Long
    Dim lngMaxHeight As Long
    Dim iCnt         As Integer
    Dim strDate      As String
    Dim strTmp       As Double
    Dim strWardTm    As String
    Dim strEmTm      As String
    Dim strOutTm     As String
    Dim strTotTm     As String
    Dim strEm1Tm     As String
    Dim strWorkArea  As String
    Dim strTestNm    As String
    Dim strTestcd    As String
    Dim varTestNm    As Variant
    Dim iRowCnt      As Integer
    Dim varTmp
    Dim intCnt       As Integer
    Dim J, k, l, m, n, o, p As Integer
    Dim strFrDate, strToDate As String
    Dim strCvrCnt As String
    Dim tmpCvrCnt As String
    
    On Error Resume Next
    
    If cboWA.ListIndex < 0 Then
        MsgBox "WA(�˻�μ�)�� �Է��Ͽ� �ּ���", vbCritical, "��ȸ����"
        cboWA.ListIndex = 0
        Exit Sub
    End If
        
     '��������
    Call txtTblClear
    
    strWorkArea = Mid(cboWA.Text, 1, 2)
    
    strRstCdSql = RstCdSql
       
    '���α׷����� ����..
    Set objProgress = New jProgressBar.clsProgress

    With objProgress
        .Container = Me
        .Width = tblResult.Width
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Height = 530
        .Message = "��������� �˻��ϰ� �ֽ��ϴ�..."
    End With

    strTmp = 0
    
    For iCnt = 1 To 12
        strDate = Format(dtpFromDt.Value, "yyyy") & "-" & Format(iCnt, "0#")
        
        I = 0: J = 0: k = 0: l = 0: m = 0: n = 0: o = 0: p = 0

        strFrDate = strDate & "-01"
        
        Select Case Mid(strDate, 6, 2)
            Case "01", "03", "05", "07", "08", "10", "12"
                strToDate = strDate & "-31"
            Case "02"
                strToDate = strDate & "-28"
            Case Else
                strToDate = strDate & "-30"
        End Select
        
'        SSQL = ""
'        SSQL = SSQL & vbLf & "SELECT * FROM S2COM102 "
'        SSQL = SSQL & vbLf & " WHERE substr(TRANSDT,1,7) = '" & strDate & "' "
'        SSQL = SSQL & vbLf & "   AND SUBSTR(REMARK,1,2) = '" & strWorkArea & "'"

        SSQL = ""
        SSQL = SSQL & "  SELECT /*+ INDEX (a MDNOTIFT_IDX1) +*/ "
        SSQL = SSQL & "       b.REMARK, b.TRANSMSG, b.TESTCD "
        SSQL = SSQL & "  FROM MDNOTIFT a,"
        SSQL = SSQL & "       S2COM102 b"
        SSQL = SSQL & " WHERE a.notidate BETWEEN to_date('" & strFrDate & "','yyyy-mm-dd') AND to_date('" & strToDate & "','yyyy-mm-dd')"
        SSQL = SSQL & "       AND a.notitype = '7'"
        SSQL = SSQL & "       AND SUBSTR(a.workarea,1,2) = '" & strWorkArea & "' "
'        SSQL = SSQL & "       AND a.RECVDATE IS NOT NULL"
        SSQL = SSQL & "       AND a.workarea = b.REMARK"
        
        Select Case strWorkArea
            Case "01": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork01 & ""
            Case "02": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork02 & ""
            Case "03": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork03 & ""
            'Case "04": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork04 & ""
            Case "05": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork05 & ""
            Case "07": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork07 & ""
'            Case "08": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork08 & ""
            Case "OT": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWorkOT & ""
            Case "15": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork15 & ""
        End Select
        
        SSQL = SSQL & " GROUP BY b.REMARK, b.TRANSMSG, b.TESTCD "
        
'        SSQL = ""
'        SSQL = SSQL & " SELECT * FROM MDNOTIFT "
'        SSQL = SSQL & "   WHERE  notidate between to_date('" & strDate & "','yyyymm') and to_date('" & strDate & "','yyyymm') "
'        SSQL = SSQL & "     AND notitype = '7' "
'        SSQL = SSQL & "     AND SUBSTR(workarea,1,2) = '" & strWorkArea & "' "
        
        RS.Open SSQL, DBConn
            
        If RS.RecordCount > 0 Then
            For intCnt = 1 To RS.RecordCount
                varTestNm = Split(RS.Fields("TRANSMSG").Value & "", vbCrLf)
                strTestNm = Trim(medGetP(varTestNm(1), 1, ":"))
                strTestNm = Replace(strTestNm, " ", "")
                strTestNm = Replace(strTestNm, vbCr, "")
                strTestNm = Replace(strTestNm, vbLf, "")
                strTestNm = Replace(strTestNm, vbCrLf, "")
                strTestcd = RS.Fields("TESTCD").Value & ""
                With tblResult
    '                For iRowCnt = 1 To .MaxRows
    '                    .GetText 2, iRowCnt, varTmp
    '                    If UCase(strTestNm) = UCase(varTmp) Then
    '
    '                    End If
    '                Next
                    Select Case Mid(cboWA.Text, 1, 2)
                        Case "02"
'                            Select Case Trim(UCase(strTestNm))
'                                Case "GLUCOSE": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "CA": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "PI": k = k + 1:  .SetText iCnt + 5, 3, k
'                                Case "MG": l = l + 1:  .SetText iCnt + 5, 4, l
'                                Case "NA": m = m + 1:  .SetText iCnt + 5, 5, m
'                                Case "K": n = n + 1:  .SetText iCnt + 5, 6, n
'                                Case "CL": o = o + 1:  .SetText iCnt + 5, 7, o
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select
'                            Select Case Trim(UCase(strTestNm))
'                                Case "GLUCOSE": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "CA": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "NA": k = k + 1:  .SetText iCnt + 5, 3, k
'                                Case "K": l = l + 1:  .SetText iCnt + 5, 4, l
''                                Case Else
''                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select
                            Select Case Trim(RS.Fields("TESTCD"))
                                Case "C3711": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "C3795": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "C3791": k = k + 1:  .SetText iCnt + 5, 3, k
                                Case "C3792": l = l + 1:  .SetText iCnt + 5, 4, l
                                Case "C4602A": m = m + 1:  .SetText iCnt + 5, 5, m
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "01"
'                            Select Case Trim(UCase(strTestNm))
'                                Case "HEMOGLOBIN": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "PLTCOUNT": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "BLAST": k = k + 1:  .SetText iCnt + 5, 3, k
'                                Case "PBS": l = l + 1:  .SetText iCnt + 5, 4, l
'                                Case "MALARIA": m = m + 1:  .SetText iCnt + 5, 5, m
'                                Case "PT": n = n + 1:  .SetText iCnt + 5, 6, n
'                                Case "APTT": o = o + 1:  .SetText iCnt + 5, 7, o
'
''                                Case Else
''                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select
                            Select Case Trim(RS.Fields("TESTCD"))
                                Case "B1010": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "B106001": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "B1060": k = k + 1:  .SetText iCnt + 5, 3, k
                                Case "B109108": l = l + 1:  .SetText iCnt + 5, 4, l
                                Case "B1150": m = m + 1:  .SetText iCnt + 5, 5, m
                                Case "B1540C": n = n + 1:  .SetText iCnt + 5, 6, n
                                Case "B1530": o = o + 1:  .SetText iCnt + 5, 7, o

'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "05"
'                            Select Case Trim(strTestNm)
'                                Case "RHTYPING": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "AntibodyScreeningtest": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "Rho(D)typing": k = k + 1:  .SetText iCnt + 5, 3, k
''                                Case Else
''                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select
                            Select Case Trim(RS.Fields("TESTCD"))
                                Case "B2021": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "B2061": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "B1060": k = k + 1:  .SetText iCnt + 5, 3, k
'                                Case "B109108": l = l + 1:  .SetText iCnt + 5, 4, l
'                                Case "B1150": m = m + 1:  .SetText iCnt + 5, 5, m
'                                Case "B1540C": n = n + 1:  .SetText iCnt + 5, 6, n
'                                Case "B1530": o = o + 1:  .SetText iCnt + 5, 7, o

'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "03"
'                            Select Case Trim(strTestNm)
'                                Case "�ŵ���ü(����)": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "H-Widal": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "HIV AG/AB": k = k + 1:  .SetText iCnt + 5, 3, k
'                                Case "HAVAbIgM.": l = l + 1:  .SetText iCnt + 5, 4, l
''                                Case Else
''                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select

                            Select Case Trim(RS.Fields("TESTCD"))
                                Case "C4612A": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "C4712A": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "C46401": k = k + 1: .SetText iCnt + 5, 3, k
                                Case "C46402": l = l + 1: .SetText iCnt + 5, 4, l
                                Case "C4862A": m = m + 1:  .SetText iCnt + 5, 5, m
'                                Case "C4601": m = m + 1: .SetText iCnt + 5, 5, m
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "15"
                            Select Case Trim(RS.Fields("TESTCD"))
                                Case "CPEPCR5", "CPEPCR4", "CPEPCR3", "CPEPCR2", "CPEPCR1": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "PNBPCR3": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "07"
                            Select Case Trim(strTestNm)
                                Case "KetoneBody(UA)": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "AMM.URATE": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "OT"
'                            Select Case Trim(strTestNm)
'                                Case "Triplemarker(DOWN)": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "17-��-OHP�缺(��õ���ν������������������˻�)": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "Methionine": k = k + 1: .SetText iCnt + 5, 3, k
'                                Case "NST": l = l + 1: .SetText iCnt + 5, 4, l
'                                Case "FISH,X/Y(Blood)": m = m + 1: .SetText iCnt + 5, 5, m
'                                Case "MS/MS�����˻�缺": n = n + 1: .SetText iCnt + 5, 6, n
''                                Case Else
''                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select
                            Select Case Trim(RS.Fields("TESTCD"))
                                Case "27BM": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "17-��-OHP�缺(��õ���ν������������������˻�)": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "Methionine": k = k + 1: .SetText iCnt + 5, 3, k
                                Case "NST": l = l + 1: .SetText iCnt + 5, 4, l
                                Case "FISH,X/Y(Blood)": m = m + 1: .SetText iCnt + 5, 5, m
                                Case "Y995": n = n + 1: .SetText iCnt + 5, 6, n
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "04"
                            Select Case Trim(strTestNm)
                                Case "���׹���߰�����": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "Indiaink": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "08"
                            Select Case Trim(strTestNm)
                                Case "AFBStain(����������)": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                    End Select
                End With
                
                RS.MoveNext
            Next
            tblResult.SetText iCnt + 5, 9, I + J + k + l + m + n + o + p
            
            strCvrCnt = I + J + k + l + m + n + o + p
            tmpCvrCnt = Fix(strCvrCnt + (strCvrCnt * 0.4))
            
'            strFrDate = Replace(strDate, "-", "") & "01"
'            strToDate = Replace(strDate, "-", "") & "31"
            
'            SSQL = ""
'            SSQL = SSQL & " SELECT COUNT(*) AS SEQ FROM S2LAB302 "
'            SSQL = SSQL & "  WHERE ACCDT between '" & strFrDate & "' and '" & strToDate & "' "
'            SSQL = SSQL & "    AND WORKAREA = '" & strWorkArea & "' "
'            SSQL = SSQL & "    AND dpdiv like '%C' "
'            Select Case strWorkArea
'                Case "01": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork01 & ""
'                Case "02": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork02 & ""
'                Case "03": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork03 & ""
'                Case "04": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork04 & ""
'                Case "05": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork05 & ""
'                Case "07": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork07 & ""
'                Case "08": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork08 & ""
'            End Select

'            RS1.Open SSQL, DBConn
            
            strFrDate = strDate & "-01"
            
            Select Case Mid(strToDate, 6, 2)
                Case "01", "03", "05", "07", "08", "10", "12"
                    strToDate = strDate & "-31"
                Case "02"
                    strToDate = strDate & "-28"
                Case Else
                    strToDate = strDate & "-30"
            End Select
            
            SSQL = ""
            SSQL = SSQL & "  SELECT /*+ INDEX (a MDNOTIFT_IDX1) +*/ "
            SSQL = SSQL & "       COUNT(*) AS SEQ, b.REMARK, b.TRANSMSG"
            SSQL = SSQL & "  FROM MDNOTIFT a,"
            SSQL = SSQL & "       S2COM102 b"
            SSQL = SSQL & " WHERE a.notidate BETWEEN to_date('" & strFrDate & "','yyyy-mm-dd') AND to_date('" & strToDate & "','yyyy-mm-dd')"
            SSQL = SSQL & "       AND a.notitype = '7'"
            SSQL = SSQL & "       AND SUBSTR(a.workarea,1,2) = '" & strWorkArea & "' "
            SSQL = SSQL & "       AND a.RECVDATE IS NOT NULL"
            SSQL = SSQL & "       AND a.workarea = b.REMARK"
            
            Select Case strWorkArea
                Case "01": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork01 & ""
                Case "02": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork02 & ""
                Case "03": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork03 & ""
                'Case "04": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork04 & ""
                Case "05": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork05 & ""
                Case "07": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork07 & ""
                'Case "08": SSQL = SSQL & vbLf & "   AND b.TESTCD IN " & strWork08 & ""
                Case "OT": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWorkOT & ""
                Case "15": SSQL = SSQL & vbLf & "   AND TESTCD IN " & strWork15 & ""
            End Select
            
            SSQL = SSQL & " GROUP BY b.REMARK, b.TRANSMSG     "
            
            RS1.Open SSQL, DBConn
            
            If RS1.RecordCount > 0 Then
                Select Case Mid(cboWA.Text, 1, 2)
                    Case "01", "02", "03", "05", "07", "OT"
                        tmpCvrCnt = RS1.RecordCount ' RS1.Fields("SEQ").Value
                    Case Else
                        tmpCvrCnt = tmpCvrCnt
                End Select
                tblResult.SetText iCnt + 5, 10, tmpCvrCnt
'                tblResult.SetText iCnt + 5, 11, Fix(Val(strCvrCnt) / Val(tmpCvrCnt) * 100) & "%"
                tblResult.SetText iCnt + 5, 11, Fix(Val(tmpCvrCnt) / Val(strCvrCnt) * 100) & "%"
            End If
            
        End If
        RS.Close
        RS1.Close
    Next
    
    With tblResult
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .BlockMode = False
    End With
    
    Set RS = Nothing
    Set RS1 = Nothing
    Set objPatient = Nothing
End Sub

Private Sub tblResult_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'    If Col = 15 Then
'        If Trim(aryResultText(Row)) <> "" Then
'            txtRst.TextRTF = aryResultText(Row)
'            txtRst.Visible = True
'            txtRst.ZOrder 0
'            DoEvents
'        End If
'    End If
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
'    Static iSortOrder As Integer
'    Dim I As Double
'
'    '-- �߰� Colum�� Sort By M.G.Choi 2002.10.09
'    With tblResult
'        If Row = 0 Then
'            .SortBy = SortByRow
'            .SortKey(1) = Col
'            If iSortOrder = SortKeyOrderAscending Then
'                .SortKeyOrder(1) = SortKeyOrderDescending
'                iSortOrder = SortKeyOrderDescending
'            Else
'                .SortKeyOrder(1) = SortKeyOrderAscending
'                iSortOrder = SortKeyOrderAscending
'            End If
'            .Col = 1
'            .Col2 = .MaxCols
'            .Row = 0
'            .Row2 = .MaxRows
'            .Action = ActionSort
'        End If
''    End With
'
'    If Col > 1 And Col < 5 Then
'' 2008.12.17. �缺�� �۾����Դϴ�.
'' 2009.01.09 �缺�� ȯ��ID �Ķ���� �߰�
'        Dim pFrmName As String
'        Dim strPtId  As String
'        .Col = 3
'        .Row = Row
'        strPtId = .Value
'        If Len(strPtId) < 2 Then GoTo End2Stop
'
'        pFrmName = "frm401ResultView"
'
'        If ObjMyUser(pFrmName) Is Nothing Then GoTo End2Stop
'        If Not ObjMyUser(pFrmName).CanRead Then GoTo End2Stop
'
''        medMain.lblSubMenu.Caption = "ó������ȸ"
'
''        frmLisReviewInStatisticLib.ButtonKey = "LIS155B" 'Button.Key
''        frmLisReviewInStatisticLib.PTid = strPtId
''        frmLisReviewInStatisticLib.show
''        frmLisReview.show
''        frmLisReviewInStatisticLib.ShowThisForm
''        frmLisReviewInStatisticLib.ZOrder 0
'End2Stop:
'    Exit Sub
'
'
'    End If
'    If Col = 15 Then
'' 2009.04.13 �缺�� ary����� �����ϱ����� i�� �����ϰ� ��ư�� ���ڸ� Row�� ������.
''    With tblResult
'        .Row = Row: .Col = Col: I = Val(.TypeButtonText)
''    End With
'
'    End If
'
'    End With

End Sub

'���콺�� ���� ��Ŀ���� ���̺�� �ű��� Tooltip �����ֱ�����..
Private Sub tblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tblResult.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call TxtClear
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    ' �̰����� �̺�Ʈ�� �߻����Ѿ� �մϴ�.
    If IsLastForm Then RaiseEvent LastFormUnload
    RaiseEvent FormClose
End Sub

Private Sub cmdQuary_Click()
    Dim objProgress  As jProgressBar.clsProgress
    Dim RS           As New Recordset
    Dim RS1          As New Recordset
    Dim objPatient   As New clsPatient      'ȯ�� Ŭ����
    Dim SSQL         As String
    Dim strRstCdSql  As String
    Dim strDeptCd    As String
    Dim I            As Long
    Dim lngMaxHeight As Long
    Dim iCnt         As Integer
    Dim strDate      As String
    Dim strTmp       As Double
    Dim strWardTm    As String
    Dim strEmTm      As String
    Dim strOutTm     As String
    Dim strTotTm     As String
    Dim strEm1Tm     As String
    Dim strWorkArea  As String
    Dim strTestNm    As String
    Dim varTestNm    As Variant
    Dim iRowCnt      As Integer
    Dim varTmp
    Dim intCnt       As Integer
    Dim J, k, l, m, n, o, p As Integer
    
    On Error Resume Next
    
    If cboWA.ListIndex < 0 Then
        MsgBox "WA(�˻�μ�)�� �Է��Ͽ� �ּ���", vbCritical, "��ȸ����"
        cboWA.ListIndex = 0
        Exit Sub
    End If
        
     '��������
    Call txtTblClear
    
    strWorkArea = Mid(cboWA.Text, 1, 2)
    
    strRstCdSql = RstCdSql
       
    '���α׷����� ����..
    Set objProgress = New jProgressBar.clsProgress

    With objProgress
        .Container = Me
        .Width = tblResult.Width
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Height = 530
        .Message = "��������� �˻��ϰ� �ֽ��ϴ�..."
    End With

    strTmp = 0
    
    For iCnt = 1 To 12
        strDate = Format(dtpFromDt.Value, "yyyy") & "-" & Format(iCnt, "0#")
        
        I = 0: J = 0: k = 0: l = 0: m = 0: n = 0: o = 0: p = 0
        
        SSQL = ""
        SSQL = SSQL & vbLf & "SELECT * FROM S2COM102 "
        SSQL = SSQL & vbLf & " WHERE substr(TRANSDT,1,7) = '" & strDate & "' "
        SSQL = SSQL & vbLf & "   AND SUBSTR(REMARK,1,2) = '" & strWorkArea & "'"
    
'        SSQL = ""
'        SSQL = SSQL & " SELECT * FROM MDNOTIFT "
'        SSQL = SSQL & "   WHERE  notidate between to_date('" & strDate & "','yyyymm') and to_date('" & strDate & "','yyyymm') "
'        SSQL = SSQL & "     AND notitype = '7' "
'        SSQL = SSQL & "     AND SUBSTR(workarea,1,2) = '" & strWorkArea & "' "
        
        RS.Open SSQL, DBConn
            
        If RS.RecordCount > 0 Then
            For intCnt = 1 To RS.RecordCount
                varTestNm = Split(RS.Fields("TRANSMSG").Value & "", vbCrLf)
                strTestNm = Trim(medGetP(varTestNm(1), 1, ":"))
                strTestNm = Replace(strTestNm, " ", "")
                strTestNm = Replace(strTestNm, vbCr, "")
                strTestNm = Replace(strTestNm, vbLf, "")
                strTestNm = Replace(strTestNm, vbCrLf, "")
                With tblResult
    '                For iRowCnt = 1 To .MaxRows
    '                    .GetText 2, iRowCnt, varTmp
    '                    If UCase(strTestNm) = UCase(varTmp) Then
    '
    '                    End If
    '                Next
                    Select Case Mid(cboWA.Text, 1, 2)
                        Case "02"
'                            Select Case Trim(UCase(strTestNm))
'                                Case "GLUCOSE": I = I + 1: .SetText iCnt + 5, 1, I
'                                Case "CA": J = J + 1: .SetText iCnt + 5, 2, J
'                                Case "PI": k = k + 1:  .SetText iCnt + 5, 3, k
'                                Case "MG": l = l + 1:  .SetText iCnt + 5, 4, l
'                                Case "NA": m = m + 1:  .SetText iCnt + 5, 5, m
'                                Case "K": n = n + 1:  .SetText iCnt + 5, 6, n
'                                Case "CL": o = o + 1:  .SetText iCnt + 5, 7, o
'                                Case Else
'                                    p = p + 1:  .SetText iCnt + 5, 8, p
'                            End Select
                            Select Case Trim(UCase(strTestNm))
                                Case "GLUCOSE": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "CA": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "NA": k = k + 1:  .SetText iCnt + 5, 3, k
                                Case "K": l = l + 1:  .SetText iCnt + 5, 4, l
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "01"
                            Select Case Trim(UCase(strTestNm))
                                Case "HEMOGLOBIN": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "PLTCOUNT": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "BLAST": k = k + 1:  .SetText iCnt + 5, 3, k
                                Case "PBS": l = l + 1:  .SetText iCnt + 5, 4, l
                                Case "MALARIA": m = m + 1:  .SetText iCnt + 5, 5, m
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "05"
                            Select Case Trim(strTestNm)
                                Case "RHTYPING": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "AntibodyScreeningtest": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "Rho(D)typing": k = k + 1:  .SetText iCnt + 5, 3, k
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "03"
                            Select Case Trim(strTestNm)
                                Case "�ŵ���ü(����)": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "H-Widal": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "HIV AG/AB": k = k + 1:  .SetText iCnt + 5, 3, k
                                Case "HAVAbIgM.": l = l + 1:  .SetText iCnt + 5, 4, l
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "07"
                            Select Case Trim(strTestNm)
                                Case "KetoneBody(UA)": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "AMM.URATE": J = J + 1: .SetText iCnt + 5, 2, J
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "OT"
                            Select Case Trim(strTestNm)
                                Case "Triplemarker(DOWN)": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "17-��-OHP�缺(��õ���ν������������������˻�)": J = J + 1: .SetText iCnt + 5, 2, J
                                Case "Methionine": k = k + 1: .SetText iCnt + 5, 3, k
                                Case "NST": l = l + 1: .SetText iCnt + 5, 4, l
                                Case "FISH,X/Y(Blood)": m = m + 1: .SetText iCnt + 5, 5, m
                                Case "MS/MS�����˻�缺": n = n + 1: .SetText iCnt + 5, 6, n
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "04"
                            Select Case Trim(strTestNm)
                                Case "���׹���߰�����": I = I + 1: .SetText iCnt + 5, 1, I
                                Case "Indiaink": J = J + 1: .SetText iCnt + 5, 2, J
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                        Case "08"
                            Select Case Trim(strTestNm)
                                Case "AFBStain(����������)": I = I + 1: .SetText iCnt + 5, 1, I
                                Case Else
                                    p = p + 1:  .SetText iCnt + 5, 8, p
                            End Select
                    End Select
                End With
                
                RS.MoveNext
            Next
            tblResult.SetText iCnt + 5, 9, I + J + k + l + m + n + o + p
            
            Dim strFrDate, strToDate As String
            Dim strCvrCnt As String
            Dim tmpCvrCnt As String
            
            strCvrCnt = I + J + k + l + m + n + o + p
            tmpCvrCnt = Fix(strCvrCnt + (strCvrCnt * 0.4))
            
            strFrDate = Replace(strDate, "-", "") & "01"
            strToDate = Replace(strDate, "-", "") & "31"
            
            SSQL = ""
            SSQL = SSQL & " SELECT COUNT(*) AS SEQ FROM S2LAB302 "
            SSQL = SSQL & "  WHERE ACCDT between '" & strFrDate & "' and '" & strToDate & "' "
            SSQL = SSQL & "    AND WORKAREA = '" & strWorkArea & "' "
            SSQL = SSQL & "    AND dpdiv like '%C' "

            RS1.Open SSQL, DBConn
            
            If RS1.RecordCount > 0 Then
                Select Case Mid(cboWA.Text, 1, 2)
                    Case "01", "02"
                        tmpCvrCnt = RS1.Fields("SEQ").Value
                    Case Else
                        tmpCvrCnt = tmpCvrCnt
                End Select
                tblResult.SetText iCnt + 5, 10, tmpCvrCnt
                tblResult.SetText iCnt + 5, 11, Fix(Val(strCvrCnt) / Val(tmpCvrCnt) * 100) & "%"
            End If
            
        End If
        RS.Close
        RS1.Close
    Next
    
    With tblResult
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .BlockMode = False
    End With
    
    Set RS = Nothing
    Set RS1 = Nothing
    Set objPatient = Nothing
End Sub

'Private Function IcdSql() As String
'
'    If Trim(txtICd(0).Text) <> "" Then
'        IcdSql = "'" & Trim(txtICd(0).Text) & "'"
'    Else
'        IcdSql = ""
'    End If
'
'    If Trim(txtICd(1).Text) <> "" Then
'        If IcdSql <> "" Then
'            IcdSql = IcdSql & "," & "'" & Trim(txtICd(1).Text) & "'"
'        Else
'            IcdSql = "'" & Trim(txtICd(1).Text) & "'"
'        End If
'    End If
'
'    If Trim(txtICd(2).Text) <> "" Then
'        If IcdSql <> "" Then
'            IcdSql = IcdSql & "," & "'" & Trim(txtICd(2).Text) & "'"
'        Else
'            IcdSql = "'" & Trim(txtICd(2).Text) & "'"
'        End If
'    End If
'
'End Function

Private Function RstCdSql() As String
    
'    If Trim(txtRstCd(0).Text) <> "" Then
'        RstCdSql = "'" & Trim(txtRstCd(0).Text) & "'"
'    Else
'        RstCdSql = ""
'    End If
'
'    If Trim(txtRstCd(1).Text) <> "" Then
'        If RstCdSql <> "" Then
'            RstCdSql = RstCdSql & "," & "'" & Trim(txtRstCd(1).Text) & "'"
'        Else
'            RstCdSql = "'" & Trim(txtRstCd(1).Text) & "'"
'        End If
'    Else
'        If RstCdSql = "" Then RstCdSql = ""
'    End If
'
'    If Trim(txtRstCd(2).Text) <> "" Then
'        If RstCdSql <> "" Then
'            RstCdSql = RstCdSql & "," & "'" & Trim(txtRstCd(2).Text) & "'"
'        Else
'            RstCdSql = "'" & Trim(txtRstCd(2).Text) & "'"
'        End If
'    Else
'        If RstCdSql = "" Then RstCdSql = ""
'    End If

End Function

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
    If blnCHkLoad = False Then
        DoEvents
        blnCHkLoad = True
        Call GetWorkAreaCombo
        'GetIcdList
    End If
End Sub

Private Sub Form_Load()
    blnCHkLoad = False
    TxtClear
    chkIndex_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSQL = Nothing
    Set objListPop = Nothing
''    Set objTMCd = Nothing
End Sub

Private Sub GetWorkAreaCombo()
    
    Dim sSqlGetWA As String
    Dim rsGetWA As Recordset
    Dim I%
    
    Set rsGetWA = New Recordset
    rsGetWA.Open objSQL.GetWACd, DBConn
    
    cboWA.Clear
    For I = 1 To rsGetWA.RecordCount
        cboWA.AddItem "" & rsGetWA.Fields("WACd").Value & "   " & _
                            "" & rsGetWA.Fields("WANm").Value
        rsGetWA.MoveNext
    Next I

    Set rsGetWA = Nothing

End Sub

Private Sub cmdListPop_Click(Index As Integer)
'    Dim objData As clsBasisData
    
    '����Ʈ �˾��� �ҷ�����...
    Set objListPop = New clsPopUpList
'    Set objData = New clsBasisData
    
    With objListPop
        .Connection = DBConn
'        .BackColor = Me.BackColor
        Select Case Index
            '��ü�ڵ� �ҷ�����
            Case 0:
'                .Caption = "��ü�ڵ� ��ȸ"
'                .HeadName = "��ü�ڵ�, ��ü��"
'                .Width = .Width + 700
'                Call .ListPop(objSql.GetSpcList, 2950, 4700)
'                txtSpcCd.Text = medGetP(.SelectedString, 1, ";")
'                lblTNm.Caption = medGetP(.SelectedString, 2, ";")
                
            '���ڵ� �ҷ�����
            Case 1:
'                If objIcdList Is Nothing Then
'                    Call GetIcdList
'                End If
'                .Caption = "���ڵ� ��ȸ"
'                .HeadName = "���ڵ�, �󺴸�"
'                .Width = .Width + 700
'                Call .ListPop(, 3350, 4700, objIcdList)
'                If Trim(txtICd(0).Text) = "" Then
'                    txtICd(0).Text = medGetP(.SelectedString, 1, ";")
'                ElseIf Trim(txtICd(1).Text) = "" Then
'                    If Trim(txtICd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtICd(1).Text = ""
'                    Else
'                        txtICd(1).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                Else
'                    If Trim(txtICd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Or _
'                       Trim(txtICd(1).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtICd(2).Text = ""
'                    Else
'                        txtICd(2).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                End If
            '����ڵ� �ҷ�����
            Case 2:
                Dim objRstSQL As New clsLISSqlETest
                .FormCaption = "����ڵ� ��ȸ"
                .ColumnHeaderText = "����ڵ�;�����"
'                .Width = .Width + 700
                Call .LoadPopUp(objRstSQL.SqlGetSpeRstCode(txtTestCd.Text))  ', 3750, 4700, objRstCd)
'                If Trim(txtRstCd(0).Text) = "" Then
'                    txtRstCd(0).Text = medGetP(.SelectedString, 1, ";")
'                ElseIf Trim(txtRstCd(1).Text) = "" Then
'                    If Trim(txtRstCd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtRstCd(1).Text = ""
'                    Else
'                        txtRstCd(1).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                Else
'                    If Trim(txtRstCd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Or _
'                       Trim(txtRstCd(1).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtRstCd(2).Text = ""
'                    Else
'                        txtRstCd(2).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                End If
                Set objRstSQL = Nothing
            '����� �ҷ�����
            Case 3:
                .FormCaption = "����� ��ȸ"
                .ColumnHeaderText = "������ڵ�;�������"
'                .Width = .Width + 300
'                .ColSize(0) = 1000
                Call .LoadPopUp(GetSQLDeptList) ', 3950, 9300) ', ObjLISComCode.DeptCd)
'                txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
'                lblDeptNm.Caption = medGetP(.SelectedString, 2, ";")
'
            Case 4:
'                .Caption = "��ü�ڵ� ��ȸ"
'                .HeadName = "��ü�ڵ�, ��ü��"
'                .Width = .Width + 700
'                Call .ListPop(objSql.GetSpcListByTest(txtTestCd.Text), 2950, 4700)
'                txtSpcCd.Text = medGetP(.SelectedString, 1, ";")
'                lblTNm.Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
'    Set objData = Nothing
    Set objListPop = Nothing
    
End Sub

Private Sub TxtClear()
    
   
    '��ȸ�Ⱓ
    dtpFromDt.Value = GetSystemDate
    dtpToDt.Value = GetSystemDate
       
    '��������
    Call txtTblClear
End Sub

Private Sub txtTblClear()
    medClearTable tblResult
    tblResult.MaxRows = 0
    tblResult.RowHeight(-1) = 15
    
    strWork01 = "('B1010','B1060','B109108','B1100','B1150','B1540','B1530','B1540C','B106001','B109108')"
    strWork02 = "('C3711','C3795','C3791','C3792','C4602A')"
    strWork03 = "('C4612A','C4712A','C46401','C46402','C4862A')"
    strWork04 = "('B40561','B4111')"
    strWork05 = "('B2021','B2061','B2047')"
    strWork07 = "('B00306','B004123')"
    strWork08 = "('27BM','B4021AE','Y995')"
    strWork15 = "('CPEPCR5','CPEPCR4','CPEPCR3','CPEPCR2','CPEPCR1','PNBPCR3')"
    strWorkOT = "('27BM','B4021AE','Y995')"
    
    With tblResult
        Select Case Mid(cboWA.Text, 1, 2)
            Case "01"
                .MaxRows = 11
                .SetText 1, 1, "��������": .SetText 2, 1, "Hemoglobin": .SetText 3, 1, "B1010": .SetText 4, 1, "5.0": .SetText 5, 1, "":
                .SetText 1, 2, "��������": .SetText 2, 2, "Platelet": .SetText 3, 2, "B106001": .SetText 4, 2, "30,000": .SetText 5, 2, "":
                .SetText 1, 3, "��������": .SetText 2, 3, "Blast": .SetText 3, 3, "B109108": .SetText 4, 3, "": .SetText 5, 3, "":
                .SetText 1, 4, "��������": .SetText 2, 4, "PBS": .SetText 3, 4, "B1100": .SetText 4, 4, "": .SetText 5, 4, "":
                .SetText 1, 5, "��������": .SetText 2, 5, "Malaria": .SetText 3, 5, "CZ397": .SetText 4, 5, "": .SetText 5, 5, "�缺": 'B1150
                .SetText 1, 6, "��������": .SetText 2, 6, "PT": .SetText 3, 6, "B1540C": .SetText 4, 6, "": .SetText 5, 6, "I.N.R 4.0":
                .SetText 1, 7, "��������": .SetText 2, 7, "aPTT": .SetText 3, 7, "B1530": .SetText 4, 7, "": .SetText 5, 7, "180sec":
                .SetText 1, 8, "��������": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "02"
'                .MaxRows = 11
'                .SetText 1, 1, "�ӻ�ȭ��": .SetText 2, 1, "Glucose": .SetText 3, 1, "C3711": .SetText 4, 1, "25": .SetText 5, 1, "1000":
'                .SetText 1, 2, "�ӻ�ȭ��": .SetText 2, 2, "Ca": .SetText 3, 2, "C3795": .SetText 4, 2, "5.0": .SetText 5, 2, "14.0":
'                .SetText 1, 3, "�ӻ�ȭ��": .SetText 2, 3, "Pi": .SetText 3, 3, "C3794": .SetText 4, 3, "1.0": .SetText 5, 3, "8.0":
'                .SetText 1, 4, "�ӻ�ȭ��": .SetText 2, 4, "Mg": .SetText 3, 4, "C3797": .SetText 4, 4, "1.0": .SetText 5, 4, "":
'                .SetText 1, 5, "�ӻ�ȭ��": .SetText 2, 5, "Na": .SetText 3, 5, "C3791": .SetText 4, 5, "110": .SetText 5, 5, "155":
'                .SetText 1, 6, "�ӻ�ȭ��": .SetText 2, 6, "K": .SetText 3, 6, "C3792": .SetText 4, 6, "2.2": .SetText 5, 6, "7.0":
'                .SetText 1, 7, "�ӻ�ȭ��": .SetText 2, 7, "Cl": .SetText 3, 7, "C3793": .SetText 4, 7, "70": .SetText 5, 7, "125":
'                .SetText 1, 8, "�ӻ�ȭ��": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
'                .SetText 5, 9, "���۰Ǽ�":
'                .SetText 5, 10, "CVR�Ǽ�":
'                .SetText 5, 11, "������":
                .MaxRows = 11
                .SetText 1, 1, "�ӻ�ȭ��": .SetText 2, 1, "Glucose": .SetText 3, 1, "C3711": .SetText 4, 1, "35": .SetText 5, 1, "550":
                .SetText 1, 2, "�ӻ�ȭ��": .SetText 2, 2, "Ca": .SetText 3, 2, "C3795": .SetText 4, 2, "5.0": .SetText 5, 2, "14.0":
                .SetText 1, 3, "�ӻ�ȭ��": .SetText 2, 3, "Na": .SetText 3, 3, "C3791": .SetText 4, 3, "110": .SetText 5, 3, "155":
                .SetText 1, 4, "�ӻ�ȭ��": .SetText 2, 4, "K": .SetText 3, 4, "C3792": .SetText 4, 4, "2.2": .SetText 5, 4, "7.0":
                .SetText 1, 5, "�ӻ�ȭ��": .SetText 2, 5, "RPR-H": .SetText 3, 5, "C4602A": .SetText 4, 5, "": .SetText 5, 5, "1.0�̻�":
                .SetText 1, 6, "�ӻ�ȭ��": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "�ӻ�ȭ��": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "�ӻ�ȭ��": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "05"
                .MaxRows = 11
                .SetText 1, 1, "��������": .SetText 2, 1, "Rho(D)": .SetText 3, 1, "B2021": .SetText 4, 1, "": .SetText 5, 1, "����":
                .SetText 1, 2, "��������": .SetText 2, 2, "Antibody Screening": .SetText 3, 2, "B2061": .SetText 4, 2, "": .SetText 5, 2, "�缺":
                .SetText 1, 3, "��������": .SetText 2, 3, "Direct Coombs'": .SetText 3, 3, "B2047": .SetText 4, 3, "": .SetText 5, 3, "�缺":
                .SetText 1, 4, "��������": .SetText 2, 4, "": .SetText 3, 4, "": .SetText 4, 4, "": .SetText 5, 4, "":
                .SetText 1, 5, "��������": .SetText 2, 5, "": .SetText 3, 5, "": .SetText 4, 5, "": .SetText 5, 5, "":
                .SetText 1, 6, "��������": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "��������": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "��������": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "03"
                .MaxRows = 11
                .SetText 1, 1, "�鿪��û": .SetText 2, 1, "�ŵ���������": .SetText 3, 1, "C4612A": .SetText 4, 1, "": .SetText 5, 1, "�缺": 'C4612AB
                .SetText 1, 2, "�鿪��û": .SetText 2, 2, "HIV": .SetText 3, 2, "C4712A": .SetText 4, 2, "": .SetText 5, 2, "Positive":
                .SetText 1, 3, "�鿪��û": .SetText 2, 3, "O.Widal": .SetText 3, 3, "C46401": .SetText 4, 3, "": .SetText 5, 3, ">1:80":
                .SetText 1, 3, "�鿪��û": .SetText 2, 4, "H.Widal": .SetText 3, 4, "C46402": .SetText 4, 4, "": .SetText 5, 4, ">1:160":
                .SetText 1, 4, "�鿪��û": .SetText 2, 5, "HAVAb IgM": .SetText 3, 5, "C4862A": .SetText 4, 5, "": .SetText 5, 5, "Positive": 'C4862A
'                .SetText 1, 5, "�鿪��û": .SetText 2, 5, "RPR CRAD": .SetText 3, 5, "C4601": .SetText 4, 5, "": .SetText 5, 5, "1.0�̻�":
                .SetText 1, 5, "�鿪��û": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 6, "�鿪��û": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 7, "�鿪��û": .SetText 2, 8, "": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 1, 8, "�鿪��û": .SetText 2, 9, "��Ÿ": .SetText 3, 9, "": .SetText 4, 9, "": .SetText 5, 9, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "07"
                .MaxRows = 11
                .SetText 1, 1, "�����": .SetText 2, 1, "Ketone": .SetText 3, 1, "B00306": .SetText 4, 1, "": .SetText 5, 1, ">3+":
                .SetText 1, 2, "�����": .SetText 2, 2, "U.Micro : Tumor cell": .SetText 3, 2, "B004123": .SetText 4, 2, "": .SetText 5, 2, "":
                .SetText 1, 3, "�����": .SetText 2, 3, "": .SetText 3, 3, "": .SetText 4, 3, "": .SetText 5, 3, "":
                .SetText 1, 4, "�����": .SetText 2, 4, "": .SetText 3, 4, "": .SetText 4, 4, "": .SetText 5, 4, "":
                .SetText 1, 5, "�����": .SetText 2, 5, "": .SetText 3, 5, "": .SetText 4, 5, "": .SetText 5, 5, "":
                .SetText 1, 6, "�����": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "�����": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "�����": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "OT"
                .MaxRows = 11
                .SetText 1, 1, "���ְ˻�": .SetText 2, 1, "Triple marker(DOWN)": .SetText 3, 1, "27BM": .SetText 4, 1, "": .SetText 5, 1, "":
                .SetText 1, 2, "���ְ˻�": .SetText 2, 2, "17a-OHP(hydroxyprogesteron)": .SetText 3, 2, "": .SetText 4, 2, "": .SetText 5, 2, "":
                .SetText 1, 3, "���ְ˻�": .SetText 2, 3, "Methionine": .SetText 3, 3, "": .SetText 4, 3, "": .SetText 5, 3, "":
                .SetText 1, 4, "���ְ˻�": .SetText 2, 4, "NST": .SetText 3, 4, "": .SetText 4, 4, "": .SetText 5, 4, "":
                .SetText 1, 5, "���ְ˻�": .SetText 2, 5, "FISH,X/Y(Blood)": .SetText 3, 5, "": .SetText 4, 5, "": .SetText 5, 5, "":
                .SetText 1, 6, "���ְ˻�": .SetText 2, 6, "MS/MS�����˻�缺": .SetText 3, 6, "Y995": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "���ְ˻�": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "���ְ˻�": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "04"
                .MaxRows = 11
                .SetText 1, 1, "MicroBiology": .SetText 2, 1, "���׹��": .SetText 3, 1, "B40561": .SetText 4, 1, "": .SetText 5, 1, "���׹���߰�����":
                .SetText 1, 2, "MicroBiology": .SetText 2, 2, "Indiaink": .SetText 3, 2, "B4111": .SetText 4, 2, "": .SetText 5, 2, "Positive":
                .SetText 1, 3, "MicroBiology": .SetText 2, 3, "": .SetText 3, 3, "": .SetText 4, 3, "": .SetText 5, 3, "":
                .SetText 1, 4, "MicroBiology": .SetText 2, 4, "": .SetText 3, 4, "": .SetText 4, 4, "": .SetText 5, 4, "":
                .SetText 1, 5, "MicroBiology": .SetText 2, 5, "": .SetText 3, 5, "": .SetText 4, 5, "": .SetText 5, 5, "":
                .SetText 1, 6, "MicroBiology": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "MicroBiology": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "MicroBiology": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "08"
                .MaxRows = 11
                .SetText 1, 1, "���ٰ˻�": .SetText 2, 1, "AFB ����������": .SetText 3, 1, "B4021AE": .SetText 4, 1, "1+/2+/3+": .SetText 5, 1, "":
                .SetText 1, 2, "���ٰ˻�": .SetText 2, 2, "": .SetText 3, 2, "": .SetText 4, 2, "": .SetText 5, 2, "":
                .SetText 1, 3, "���ٰ˻�": .SetText 2, 3, "": .SetText 3, 3, "": .SetText 4, 3, "": .SetText 5, 3, "":
                .SetText 1, 4, "���ٰ˻�": .SetText 2, 4, "": .SetText 3, 4, "": .SetText 4, 4, "": .SetText 5, 4, "":
                .SetText 1, 5, "���ٰ˻�": .SetText 2, 5, "": .SetText 3, 5, "": .SetText 4, 5, "": .SetText 5, 5, "":
                .SetText 1, 6, "���ٰ˻�": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "���ٰ˻�": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "���ٰ˻�": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
            Case "15"
                .MaxRows = 11
                .SetText 1, 1, "��������": .SetText 2, 1, "CPE": .SetText 3, 1, "CPEPCR": .SetText 4, 1, "Detected": .SetText 5, 1, "":
                .SetText 1, 2, "��������": .SetText 2, 2, "Legionella pneumophila": .SetText 3, 2, "PNBPCR3": .SetText 4, 2, "Positive": .SetText 5, 2, "":
                .SetText 1, 3, "��������": .SetText 2, 3, "Bordetella pertussis": .SetText 3, 3, "PNBPCR4": .SetText 4, 3, "Positive": .SetText 5, 3, "":
                .SetText 1, 4, "��������": .SetText 2, 4, "Bordetella parapertussis": .SetText 3, 4, "PNBPCR5": .SetText 4, 4, "Positive": .SetText 5, 4, "":
                .SetText 1, 5, "��������": .SetText 2, 5, "": .SetText 3, 5, "": .SetText 4, 5, "": .SetText 5, 5, "":
                .SetText 1, 6, "��������": .SetText 2, 6, "": .SetText 3, 6, "": .SetText 4, 6, "": .SetText 5, 6, "":
                .SetText 1, 7, "��������": .SetText 2, 7, "": .SetText 3, 7, "": .SetText 4, 7, "": .SetText 5, 7, "":
                .SetText 1, 8, "��������": .SetText 2, 8, "��Ÿ": .SetText 3, 8, "": .SetText 4, 8, "": .SetText 5, 8, "":
                .SetText 5, 9, "CVR�Ǽ�":
                .SetText 5, 10, "Ȯ�ΰǼ�":
                .SetText 5, 11, "������":
        End Select
    
    End With

'    cmdPrint.Enabled = False
    cmdExcel.Enabled = True
End Sub

'Private Sub txtAccDt_LostFocus()
'    If Trim(txtAccDt.Text) <> "" And Len(txtAccDt.Text) >= 2 Then
'        dtpFromDt.Year = "20" & Mid(txtAccDt.Text, 1, 2)
'    End If
'End Sub
'
'Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtDeptCd_LostFocus()
''    Dim objDept As clsBasisData
'    Dim strDept As String
'
'    If Trim(txtDeptCd.Text) = "" Then
'        lblDeptNm.Caption = ""
'        Exit Sub
'    End If
'
''    Set objDept = New clsBasisData
'    strDept = GetDeptNm(txtDeptCd.Text)
''    Set objDept = Nothing
'
'    If strDept <> "" Then
'        lblDeptNm.Caption = strDept
'    Else
'        medBeep (1)
'        txtDeptCd.Text = ""
'        lblDeptNm.Caption = ""
'        txtDeptCd.SetFocus
'        Exit Sub
'    End If
''
''    With ObjAPSComCode.DeptCd
''
''        If .Exists(Trim(txtDeptCd.Text)) = True Then
''            .KeyChange Trim(txtDeptCd.Text)
''            lblDeptNm.Caption = .Fields("deptnm")
''        Else
''            medbeep (1)
''            txtDeptCd.Text = ""
''            lblDeptNm.Caption = ""
''            txtDeptCd.SetFocus
''            Exit Sub
''        End If
''    End With
'End Sub

Private Sub txtFromSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

'Private Sub txtPtId_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtPtId_LostFocus()
'    Dim objPatient As New clsPatient      'ȯ�� Ŭ����
'
'    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
'
'    With objPatient
'        If Trim(txtPtId.Text) <> "" Then
'            If .GETPatient(txtPtId.Text) Then
'                lblPtInfo.Caption = .PtNm & "   " & .SEXNM & " / " & .Age & " " & .AGEDIV
'            Else
'                lblPtInfo.Caption = ""
'                MsgBox "��ϵ��� ���� ȯ��ID �Դϴ�.", vbExclamation, "�޼���"
'                Exit Sub
'            End If
'        Else
'            lblPtInfo.Caption = ""
'        End If
'    End With
'    Set objPatient = Nothing
'End Sub

'Private Sub txtRst_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then
'        txtRst.Visible = False
'    End If
'End Sub
'
'Private Sub txtRstCd_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtRstCd_LostFocus(Index As Integer)
'
'    If Trim(txtRstCd(Index).Text) = "" Then Exit Sub
'
'    With objRstCd
'        If .Exists(Trim(txtRstCd(Index).Text)) = True Then
'            Exit Sub
'        Else
'            medBeep (1)
'            txtRstCd(Index).Text = ""
'            Exit Sub
'        End If
'    End With
'
'End Sub

Private Sub PrintSpread()
    Dim objValue    As New clsDictionary
    Dim I           As Long
    Dim J           As Long
    Dim strLabNo    As String
    Dim strPtNm     As String
    Dim strPtId     As String
    Dim strSpcnm    As String
    Dim strDeptCd   As String
    Dim strDx       As String
    Dim strData     As String
    
    objValue.Clear
    objValue.FieldInialize "labno", "ptnm,ptid,spcnm,deptcd,dx"
    
    With tblResult
        For I = 1 To .MaxRows
            .Row = I
            For J = 1 To .MaxCols
                .Col = J
                Select Case J
                    Case 1: strLabNo = .Value
                    Case 2: strPtNm = .Value
                    Case 3: strPtId = .Value
                    Case 5: strSpcnm = .Value
                    Case 9: strDeptCd = .Value
                    Case 11: strDx = .Value
                End Select
            Next J
            strData = Join(Array(strPtNm, strPtId, strSpcnm, strDeptCd, strDx), COL_DIV)
            objValue.AddNew strLabNo, strData
        Next I
    End With
    
    Set objValue = Nothing
    
End Sub

Private Sub GetIcdList()

    Dim objRs As Recordset
'    Dim objIcdSql   As New clsBasisData  'clsHosComSQLStmt
    Dim objStatus As New jProgressBar.clsProgress
    
    With objStatus
        .Container = Me
        .Width = lblTitle.Width
        .Left = lblTitle.Left
        .Top = lblTitle.Top
        .Height = 280
        .Message = "���ڵ� �����͸� �ε��ϰ� �ֽ��ϴ�..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = lblTitle.Width
'        .XPos = lblTitle.Left
'        .YPos = lblTitle.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "���ڵ� �����͸� �ε��ϰ� �ֽ��ϴ�..."
'        .Value = 0
    End With

    Set objIcdList = New clsDictionary
    objIcdList.Clear
    objIcdList.FieldInialize "icd", "icdenm"
    
    Set objRs = New Recordset
    objRs.Open GetSQLIcdList, DBConn
    
    objStatus.Max = objRs.RecordCount
    
    objIcdList.Sort = False
    While Not objRs.EOF
        objStatus.Value = objStatus.Value + 1
        objStatus.Message = "���ڵ� �����͸� �ε��ϰ� �ֽ��ϴ�...(" & CInt(objStatus.Value / objStatus.Max * 100) & "%)"
        objIcdList.AddNew objRs.Fields("icd").Value & "", objRs.Fields("ienm").Value & ""
        objRs.MoveNext
    Wend
    
    Set objRs = Nothing
'    Set objIcdSql = Nothing
    Set objStatus = Nothing
    
End Sub

Private Sub GetRstCdList()

    Dim objRs As Recordset
    Dim objRstSQL As New clsLISSqlETest

    Set objRstCd = New clsDictionary
    objRstCd.Clear
    objRstCd.FieldInialize "rstcd", "rstnm"
    
    Set objRs = New Recordset
    objRs.Open objRstSQL.SqlGetSpeRstCode(txtTestCd.Text), DBConn
    
    objRstCd.Sort = False
    While Not objRs.EOF
        objRstCd.AddNew objRs.Fields("rstcd").Value & "", objRs.Fields("rstnm").Value & ""
        objRs.MoveNext
    Wend
    objRstCd.Sort = True
    
    Set objRs = Nothing
    Set objRstSQL = Nothing
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        Call txtTestCd_LostFocus
    End If
End Sub

Private Sub txtTestCd_LostFocus()

    Dim strSQL As String
    Dim objRs As Recordset
    
    Call TxtClear
    Call txtTblClear
    
    If Trim(txtTestCd.Text) = "" Then Exit Sub
    
    strSQL = objSQL.GetAccTest(txtTestCd.Text)
    Set objRs = New Recordset
    objRs.Open strSQL, DBConn
    
    If objRs.EOF Then
        MsgBox "ó���ڵ带 �ٽ� �Է��Ͻʽÿ�.", vbInformation, "ó���ڵ� �Է�"
        Set objRs = Nothing
        txtTestCd.SelStart = 0
        txtTestCd.SelLength = Len(txtTestCd.Text)
        txtTestCd.SetFocus
        Exit Sub
    Else
        lblTestNm.Caption = "" & objRs.Fields("abbrnm10").Value
    End If
    
    Set objRs = Nothing
    
    Call GetRstCdList
End Sub

Private Sub txtToSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdPrint_Click()
    Dim objSpread    As vaSpread
    Dim strTitle     As String
    Dim strPrintDate As String
    Dim strTestNm    As String
    Dim strPDate     As String
    Dim tmpTitle     As String
    Dim strDate      As String
    Dim strGb        As String
    
    strGb = ""
    strPDate = Format(Now, "yyyy-mm-dd hh:mm:ss")

    With tblResult
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .FontBold = False
        .FontSize = 9
        .BlockMode = False
               
        .PrintJobName = "CVR ����ü�� ������"

        .PrintAbortMsg = "CVR ����ü�� �������� ������Դϴ�. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
        
        tmpTitle = "CVR ����ü�� ������"
'        strTitle = "/fn""����ü""/fz""18""/fb1/fi0/fu1/fk0/fs1" _
'              & "/f1/c" & tmpTitle & "/n/n/n"
        strTitle = "/fn""����ü"" /fz""18"" /fb1/fi0/fu0/fk0/fs1" _
                  & "/f1/c" & tmpTitle & "/n/n/n"
        strPrintDate = "/fn""����ü"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "������� : " & strPDate & "/n/n"
        strTestNm = "/fn""����ü"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "WorkArea : " & cboWA.Text & "   �˻��׸� : " & lblTestNm.Caption & "/n"
        strDate = "/fn""����ü"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "��ȸ�Ⱓ : " & Format(dtpFromDt.Value, "yyyy") & "/n"
        .PrintHeader = strTitle & strTestNm & strDate 'strPrintDate
        .PrintMarginLeft = 10
'        .PrintMarginRight = 10
        .PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape
'        .PrintOrientation = PrintOrientationLandscape 'PrintOrientationLandscape
        
        
'        P_HOSPITALNAME = "�Ѹ������׿�"
        .PrintFooter = " /l " & String(130, Chr(6)) & "/n/l " & P_HOSPITALNAME & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintShadows = True
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

    End With
End Sub

'Private Sub CaseStudyHead()
'    Dim strTmp  As String
'
'    lngCurYPos = 10
'    Printer.DrawStyle = 0: Printer.DrawWidth = 6
'    Printer.FontSize = 20: Printer.FontBold = True
'    Call Print_Setting("Case Study", 0, LineSpace * 3, Printer.ScaleWidth - 0, "C", "C", True)
'    Printer.FontSize = 9: Printer.FontBold = False
'
'    strTmp = Format(dtpFromDt.Value, "YYYY�� MM�� DD��") & " ~ " & Format(dtpToDt.Value, "YYYY�� MM�� DD��")
'
'    Call Print_Setting("��ȸ�Ⱓ : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C")
'    Call Print_Setting("�������� : " & cboWA.Text, 120, LineSpace, Printer.ScaleWidth, "L", "C", False)
'    Call Print_Setting("�˻��׸� : " & txtTestCd.Text & "[" & lblTestNm.Caption & "]", 0, LineSpace, Printer.ScaleWidth, "L", "C")
'    strTmp = "[ ��ü ]": If txtPtId.Text <> "" Then strTmp = "[ " & txtPtId.Text & " ] " & lblPtInfo.Caption
'    Call Print_Setting("ȯ������ : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C", False)
'    strTmp = "[ ��ü ]": If txtAccDt.Text <> "" Then strTmp = "[ " & txtAccDt.Text & " ] " & txtFromSeq.Text & " ~ " & txtToSeq.Text
'    Call Print_Setting("������ȣ : " & strTmp, 120, LineSpace, Printer.ScaleWidth, "L", "C")
'    strTmp = "[ ��ü ]": If txtRstCd(0).Text <> "" Then strTmp = "[ " & txtRstCd(0).Text & " ] " & txtRstCd(1).Text & " ~ " & txtRstCd(2).Text
'    Call Print_Setting("����ڵ� : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C", False)
'    strTmp = "[ ��ü ]": If txtDeptCd.Text <> "" Then strTmp = "[ " & txtDeptCd.Text & " ] " & lblDeptNm.Caption
'    Call Print_Setting("�� �� �� : " & strTmp, 120, LineSpace, Printer.ScaleWidth, "L", "C")
'    strTmp = Format(GetSystemDate, "YYYY�� MM�� DD��")
'    Call Print_Setting("�� �� �� : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C")
'
'    Printer.Line (0, lngCurYPos)-(Printer.Width - 0, lngCurYPos)
'
'    '-- ����
''    Call CaseStudyBody("������ȣ", "ȯ��ID", "ȯ�ڸ�", "��/����", "��ü��", "��������", "�����", _
'                       "����", "���1", "���2", "���3", "text���")
'
'    Call CaseStudyBody("������ȣ", "ȯ��ID", "ȯ�ڸ�", "��/����", "��ü��", "��������", "��������", _
'                       "�����", "����", "���1", "", "", "")
'
'    Printer.DrawStyle = 0: Printer.DrawWidth = 6
'    Printer.Line (0, lngCurYPos)-(Printer.Width - 0, lngCurYPos)
'End Sub
'
'Private Sub CaseStudyBody(ByVal sAccno As String, ByVal sPtid As String, ByVal sPtnm As String, _
'                          ByVal sSexAge As String, ByVal sSpcNm As String, ByVal sAccDt As String, _
'                          ByVal sVfydt As String, ByVal sDept As String, ByVal sWard As String, ByVal sRst1 As String, _
'                          ByVal sRst2 As String, ByVal sRst3 As String, ByVal sTxtFg As String)
'
'    If lngCurYPos > Printer.ScaleHeight - 6 Then
'        Printer.NewPage
'        Call CaseStudyHead
'    End If
'
'    Call Print_Setting(sAccno, 0, LineSpace, 30, "L", "C", False)
'    Call Print_Setting(sPtid, 25, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sPtnm, 40, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sSexAge, 55, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sSpcNm, 70, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sAccDt, 85, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sVfydt, 120, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sDept, 155, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sWard, 170, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sRst1, 185, LineSpace, 15, "L", "C")
'
'    '** ���� -------------------------------------------------------
''    Call Print_Setting(sDept, 105, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sWard, 120, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sRst1, 135, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sRst2, 150, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sRst3, 165, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sTxtFg, 180, LineSpace, 35, "L", "C")
'    '---------------------------------------------------------------
'    Printer.DrawStyle = 2: Printer.DrawWidth = 2
'    Printer.Line (0, lngCurYPos)-(Printer.Width - 0, lngCurYPos)
'End Sub
'
'Private Sub PrintCaseStudy()
'    Dim sAccno  As String
'    Dim sPtid   As String
'    Dim sPtnm   As String
'    Dim sSexAge As String
'    Dim sSpcNm  As String
'    Dim sAccDt  As String
'    Dim sVfydt  As String
'    Dim sDept   As String
'    Dim sWard   As String
'    Dim sRst1   As String
'    Dim sRst2   As String
'    Dim sRst3   As String
'    Dim sTxtFg  As String
'
'    Dim ii          As Integer
'
'    If tblResult.DataRowCnt < 1 Then Exit Sub
'
'    Call P_PrtSet
'    Call CaseStudyHead
'
'    With tblResult
'        For ii = 1 To .DataRowCnt
'            .Row = ii
'            .Col = 1:   sAccno = .Value
'            .Col = 2:   sPtid = .Value
'            .Col = 3:   sPtnm = .Value
'            .Col = 4:   sSexAge = .Value
'            .Col = 5:   sSpcNm = .Value
'            .Col = 7:   sAccDt = .Value
'            .Col = 8:   sVfydt = .Value
'            .Col = 9:   sDept = .Value
'            .Col = 10:   sWard = .Value
'            .Col = 11:   sRst1 = .Value
''            .Col = 12:  sRst2 = .Value
''            .Col = 13:  sRst3 = .Value
''            .Col = 14:  sTxtFg = "Y"
'                        If .CellType = CellTypeStaticText Then sTxtFg = ""
'            Call CaseStudyBody(sAccno, sPtid, sPtnm, sSexAge, sSpcNm, sAccDt, sVfydt, sDept, sWard, sRst1, sRst2, sRst3, sTxtFg)
'        Next
'    End With
'
'    Printer.EndDoc
'End Sub



