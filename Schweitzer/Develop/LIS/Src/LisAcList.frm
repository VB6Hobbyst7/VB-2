VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmACList 
   BackColor       =   &H00DBE6E6&
   Caption         =   "원무환불취소"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14670
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14670
   Tag             =   "25200"
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "검 색 (&Q)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "158"
      Top             =   60
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "환불취소 (&D)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "132"
      Top             =   60
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "To &Excel"
      Height          =   510
      Left            =   11760
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "127"
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox cboWorkArea 
      Height          =   300
      Left            =   5490
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.CheckBox chkWorkArea 
      Appearance      =   0  '평면
      BackColor       =   &H8000000D&
      Caption         =   "WorkArea"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   180
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00DBE6E6&
      Caption         =   "확인된결과"
      Height          =   315
      Left            =   7860
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1815
   End
   Begin FPSpread.vaSpread ssCmtList 
      Height          =   7740
      Left            =   75
      TabIndex        =   7
      Tag             =   "45506"
      Top             =   690
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   13653
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   11
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      SpreadDesigner  =   "LisAcList.frx":0000
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   375
      Left            =   1050
      TabIndex        =   8
      Top             =   150
      Width           =   1425
      _ExtentX        =   2514
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
      Format          =   83296259
      CurrentDate     =   36328
   End
   Begin MSComCtl2.DTPicker dtpEndDt 
      Height          =   360
      Left            =   2730
      TabIndex        =   9
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
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
      Format          =   83296259
      CurrentDate     =   36328
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Index           =   0
      Left            =   75
      TabIndex        =   10
      Top             =   150
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   661
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "접수일자"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   11
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
      SpreadDesigner  =   "LisAcList.frx":1B9D
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "-"
      Height          =   240
      Left            =   2520
      TabIndex        =   12
      Top             =   225
      Width           =   270
   End
End
Attribute VB_Name = "frmACList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()
Private objRst As New clsPatientInfo
Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

Private Sub cmdExcel_Click()
    Dim strTmp  As String
    
    If ssCmtList.DataRowCnt = 0 Then Exit Sub
    
    With ssCmtList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = Trim(strTmp)
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "EmmaList"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()

    Dim SQL      As String
    Dim objProBar   As jProgressBar.clsProgress
    Dim rsGetinfo   As Recordset
    Dim RS          As Recordset
    Dim sStartDt    As String
    Dim SendDt      As String
    Dim i%
    Dim strTransNm  As String
    Dim strTAT      As String
    Dim strHH       As String
    Dim strDD       As String
    Dim strTATDt    As String
    Dim strTransFg  As String
    Dim strWorkArea As String
    Dim strSDate     As String
    Dim strEDate     As String
    Dim strNDate     As String
    Dim RS1          As Recordset
    
    strSDate = Format(dtpStartDt.Value, "YYYY-MM-DD")
    strEDate = Format(dtpEndDt.Value, "YYYY-MM-DD")
        
    SQL = ""
    SQL = " SELECT * FROM aclisrpt WHERE MEDDATE BETWEEN '" & strSDate & "'  AND '" & strEDate & "' "

'    If chkWorkArea.Value = 1 Then
'        SQL = SQL & "   AND a.WORKAREA = '" & Mid(cboWorkArea, 1, 2) & "'"
'    End If
'
'    If chkAll.Value = 1 Then
'        SQL = SQL & "   AND a.VERICHK = '1' "
'    Else
'        SQL = SQL & "   AND a.VERICHK is null "
'    End If
    
    Set objProBar = New jProgressBar.clsProgress
    
    With objProBar
        .Container = Me
        .Width = ssCmtList.Width
        .Left = ssCmtList.Left
        .Top = ssCmtList.Top - 280
        .Height = 280
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
    
    Set rsGetinfo = New Recordset
    rsGetinfo.Open SQL, DBConn
        
    If rsGetinfo.RecordCount > 0 Then
        objProBar.Max = rsGetinfo.RecordCount
    Else
        MsgBox "데이타가 없습니다.."
    End If
    
    ClearssCmtList
    
    For i = 1 To rsGetinfo.RecordCount
        ssCmtList.MaxRows = ssCmtList.MaxRows + 1
        ssCmtList.Row = ssCmtList.MaxRows + 1
        objProBar.Value = i
        DoEvents
        With rsGetinfo
            If "" & .Fields("RCPYN").Value = "Y" Then
                strTransFg = "취소확인"
            Else
                strTransFg = "미확인"
            End If
                               
            ssCmtList.SetText 2, i, GetPtNm("" & .Fields("PATNO").Value)
            ssCmtList.SetText 3, i, "" & .Fields("PATNO").Value
            ssCmtList.SetText 4, i, "" & .Fields("MEDDATE").Value
            
            SQL = ""
            SQL = " SELECT TESTNM FROM S2LAB001 WHERE TESTCD = '" & .Fields("SUGACODE").Value & "' "
            
            Set RS = New Recordset
            RS.Open SQL, DBConn
    
            ssCmtList.SetText 6, i, "" & RS.Fields("TESTNM").Value
            ssCmtList.SetText 5, i, GetEmpNm("" & .Fields("MEDDR").Value)
            ssCmtList.SetText 7, i, "" & .Fields("EDITID").Value
            ssCmtList.SetText 8, i, "" & .Fields("EDITDATE").Value
            ssCmtList.SetText 9, i, strTransFg
            ssCmtList.SetText 10, i, "" & .Fields("SUGACODE").Value
            ssCmtList.SetText 11, i, "" & .Fields("WORKAREA").Value
        End With
        rsGetinfo.MoveNext
    Next i
    
    Set rsGetinfo = Nothing
    Set RS = Nothing
    Set RS1 = Nothing
    Set objProBar = Nothing
    
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim strWA As String
    
    dtpStartDt.Value = GetSystemDate
    dtpEndDt.Value = GetSystemDate
    
    ClearssCmtList
    
    Call objRst.Load_WorkArea(cboWorkArea)
    
    '설정된 Workarea 가 있는경우 읽기
    
    strWA = GetSetting("Schweitzer2000 LIS", "Options", "UnvfyForWA", vbNullString)
    
    If strWA <> vbNullString Then
        cboWorkArea.ListIndex = Val(strWA)
    Else
        cboWorkArea.ListIndex = 0
    End If
End Sub

Private Sub ClearssCmtList()

    With ssCmtList
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With

End Sub

Private Sub dtpEndDt_Validate(Cancel As Boolean)
    ClearssCmtList
End Sub

Private Sub dtpStartDt_Validate(Cancel As Boolean)
    ClearssCmtList
End Sub


Private Sub optOption_Click(Index As Integer)
    ClearssCmtList
End Sub

Private Sub cmdPrint_Click()
    Dim strSaveDt As String
    Dim SQL       As String
    Dim strEmpId  As String
    Dim iCnt      As Integer
    Dim varTmp
    Dim strPtID, strTestCd As String
    Dim strTmp    As String
    Dim strOrdDt  As String
    
    strSaveDt = Format(GetSystemDate, "YYYY-MM-DD HH:MM:SS")
    strEmpId = ObjSysInfo.EmpId
    
    strTmp = MsgBox("확인 된 결과를 삭제합니다." & vbCRLF & "삭제하시겠습니까?", vbYesNo, "삭제확인")
    If strTmp = vbYes Then
        Screen.MousePointer = vbHourglass
        With ssCmtList
            For iCnt = 1 To .MaxRows
                .GetText 1, iCnt, varTmp
                If varTmp = 1 Then
                    .GetText 3, iCnt, varTmp: strPtID = varTmp
                    .GetText 4, iCnt, varTmp: strOrdDt = varTmp
                    .GetText 10, iCnt, varTmp: strTestCd = varTmp
                    
                    SQL = "DELETE ACLISRPT "
                    SQL = SQL + " WHERE PATNO = '" & strPtID & "'                                                            "
                    SQL = SQL + "   AND MEDDATE = '" & strOrdDt & "'                                                       "
                    SQL = SQL + "   AND SUGACODE = '" & strTestCd & "'                                                       "
                    
                    DBConn.Execute SQL
                End If
            Next
        End With
        MsgBox "삭제되었습니다.."
        ClearssCmtList
        Screen.MousePointer = vbDefault
    Else
        Exit Sub
    End If
End Sub



