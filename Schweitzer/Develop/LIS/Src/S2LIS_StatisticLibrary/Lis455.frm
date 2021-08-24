VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm455AnalysisList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "AnalysisList"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   14610
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   Tag             =   "45500"
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "To &Excel"
      Height          =   510
      Left            =   11760
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "127"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "132"
      Top             =   60
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "검 색 (&Q)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "158"
      Top             =   60
      Width           =   1320
   End
   Begin VB.Frame fraInOut 
      BackColor       =   &H00DBE6E6&
      Height          =   465
      Left            =   5085
      TabIndex        =   5
      Top             =   60
      Width           =   3960
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "모두"
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   750
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과수정사유"
         Height          =   315
         Index           =   1
         Left            =   870
         TabIndex        =   8
         Top             =   120
         Width           =   1470
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수취소사유"
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread ssCmtList 
      Height          =   7740
      Left            =   75
      TabIndex        =   2
      Tag             =   "45506"
      Top             =   615
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
      MaxCols         =   13
      OperationMode   =   1
      Protect         =   0   'False
      ShadowColor     =   14737632
      SpreadDesigner  =   "Lis455.frx":0000
      VisibleCols     =   7
      VisibleRows     =   500
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   375
      Left            =   1050
      TabIndex        =   3
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
      Format          =   83034115
      CurrentDate     =   36328
   End
   Begin MSComCtl2.DTPicker dtpEndDt 
      Height          =   360
      Left            =   2730
      TabIndex        =   4
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
      Format          =   83034115
      CurrentDate     =   36328
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Index           =   0
      Left            =   75
      TabIndex        =   6
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
      Caption         =   "검색기간"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   360
      Index           =   1
      Left            =   4140
      TabIndex        =   10
      Top             =   150
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   635
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
      Caption         =   "조회유형"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   14
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
      SpreadDesigner  =   "Lis455.frx":1E67
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
      TabIndex        =   0
      Top             =   225
      Width           =   270
   End
End
Attribute VB_Name = "frm455AnalysisList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()

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
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "AnalysisList"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdStart_Click()

    Dim strSQL      As String
    Dim objProBar   As jProgressBar.clsProgress
    Dim rsGetinfo   As Recordset
    Dim sStartDt    As String
    Dim SendDt      As String
    Dim i%
    Dim strDoct     As String
    
    sStartDt = Format(dtpStartDt.Value, CS_DateDbFormat)
    SendDt = Format(dtpEndDt.Value, CS_DateDbFormat)
          
''''    strSQL = " SELECT/*+ INDEX(a S2LAB201_IDX2) */  a.workarea, a.accdt, a.accseq, a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
''''             "        a.hosilid as roomid , a.bedid as bedid, b.empnm, c.vfyid, c.rsttxt as rsttxt, a.rcvdt, a.rcvtm, " & _
''''             "        d." & F_PTNM & " as ptnm, '' as doctnm, '' as MFYDT, '' as MFYTM, '' as MFYID " & _
''''             " FROM " & T_LAB201 & " a, " & T_LAB304 & " c, " & T_LAB015 & " b, " & _
''''                        T_HIS001 & " d " & _
''''             " WHERE " & _
''''                       DBW("a.rcvdt >= ", sStartDt) & _
''''             " AND " & DBW("a.rcvdt <= ", SendDt) & _
''''             " AND c.workarea =  a.workarea" & _
''''             " AND c.accdt    =  a.accdt" & _
''''             " AND c.accseq   =  a.accseq" & _
''''             " AND a.ptid     = d." & F_PTID & _
''''             " AND b.empid    = c.vfyid "
''''
''''    If optOption(1).Value Then
''''        strSQL = " SELECT/*+ INDEX(a S2LAB201_IDX2) */  a.workarea, a.accdt, a.accseq, a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
''''                 "        a.hosilid as roomid , a.bedid as bedid, b.empnm, c.vfyid, c.rsttxt as rsttxt, a.rcvdt, a.rcvtm, " & _
''''                 "        d." & F_PTNM & " as ptnm, '' as doctnm, f.MFYDT, f.MFYTM, f.MFYID " & _
''''                 " FROM " & T_LAB201 & " a, " & T_LAB304 & " c, " & T_LAB015 & " b, " & _
''''                            T_HIS001 & " d, " & T_LAB308 & " f " & _
''''                 " WHERE " & _
''''                           DBW("a.rcvdt >= ", sStartDt) & _
''''                 " AND " & DBW("a.rcvdt <= ", SendDt) & _
''''                 " AND c.workarea =  a.workarea" & _
''''                 " AND c.accdt    =  a.accdt" & _
''''                 " AND c.accseq   =  a.accseq" & _
''''                 " AND a.ptid     = d." & F_PTID & _
''''                 " AND b.empid    = c.vfyid " & _
''''                 " AND a.workarea =  f.workarea" & _
''''                 " AND a.accdt    =  f.accdt" & _
''''                 " AND a.accseq   =  f.accseq"
''''    ElseIf optOption(2).Value Then
''''        strSQL = strSQL & " AND   a.stscd = '" & enStsCd.StsCd_LIS_Cancel & "'"
''''    End If
          
          
    If optOption(0).Value Then
        strSQL = " SELECT/*+ INDEX(a S2LAB201_IDX2) */  DISTINCT a.workarea, a.accdt, a.accseq, a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
                 "        a.hosilid as roomid , a.bedid as bedid, b.empnm, c.vfyid, c.rsttxt as rsttxt, a.rcvdt, a.rcvtm, " & _
                 "        d." & F_PTNM & " as ptnm, e." & F_DOCTNM & " as doctnm, '' as MFYDT, '' as MFYTM, '' as MFYID  " & _
                 " FROM " & T_LAB201 & " a, " & T_LAB304 & " c, " & T_LAB015 & " b, " & _
                            T_HIS001 & " d, " & T_HIS005 & " e " & _
                 " WHERE " & _
                           DBW("a.rcvdt >= ", sStartDt) & _
                 " AND " & DBW("a.rcvdt <= ", SendDt) & _
                 " AND c.workarea =  a.workarea" & _
                 " AND c.accdt    =  a.accdt" & _
                 " AND c.accseq   =  a.accseq" & _
                 " AND a.ptid     = d." & F_PTID & _
                 " AND b.empid    = c.vfyid " & _
                 " AND e." & F_DOCTID & " = a.orddoct "

    ElseIf optOption(1).Value Then
        strSQL = " SELECT/*+ INDEX(a S2LAB201_IDX2) */ DISTINCT a.workarea, a.accdt, a.accseq, a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
                 "        a.hosilid as roomid , a.bedid as bedid, b.empnm, c.vfyid, c.rsttxt as rsttxt, a.rcvdt, a.rcvtm, " & _
                 "        d." & F_PTNM & " as ptnm, e." & F_DOCTNM & " as doctnm, f.MFYDT, f.MFYTM as mfytm, f.MFYID " & _
                 " FROM " & T_LAB201 & " a, " & T_LAB304 & " c, " & T_LAB015 & " b, " & _
                            T_HIS001 & " d, " & T_HIS005 & " e, " & T_LAB308 & " f " & _
                 " WHERE " & _
                           DBW("a.rcvdt >= ", sStartDt) & _
                 " AND " & DBW("a.rcvdt <= ", SendDt) & _
                 " AND c.workarea =  a.workarea" & _
                 " AND c.accdt    =  a.accdt" & _
                 " AND c.accseq   =  a.accseq" & _
                 " AND a.ptid     = d." & F_PTID & _
                 " AND b.empid    = c.vfyid " & _
                 " AND a.workarea =  f.workarea" & _
                 " AND a.accdt    =  f.accdt" & _
                 " AND a.accseq   =  f.accseq" & _
                 " AND e." & F_DOCTID & " = a.orddoct "
        strSQL = strSQL & " AND  exists (SELECT * FROM " & T_LAB308 & _
                                        " WHERE workarea = a.workarea " & _
                                        " AND   accdt = a.accdt " & _
                                        " AND   accseq = a.accseq ) "
    ElseIf optOption(2).Value Then
        strSQL = strSQL & " SELECT/*+ INDEX(a S2LAB201_IDX2) */  DISTINCT a.workarea, a.accdt, a.accseq, a.ptid as ptid, a.sex as sex, a.ageday as ageday, a.wardid as wardid, " & _
                 "        a.hosilid as roomid , a.bedid as bedid, b.empnm, c.vfyid, c.rsttxt as rsttxt, a.rcvdt, a.rcvtm, " & _
                 "        d." & F_PTNM & " as ptnm, a.orddoct as doctnm, c.MFYDT, c.MFTTM as mfytm, '' as MFYID " & _
                 " FROM " & T_LAB201 & " a, " & T_LAB304 & " c, " & T_LAB015 & " b, " & _
                            T_HIS001 & " d " & _
                 " WHERE " & _
                           DBW("a.rcvdt >= ", sStartDt) & _
                 " AND " & DBW("a.rcvdt <= ", SendDt) & _
                 " AND c.workarea =  a.workarea" & _
                 " AND c.accdt    =  a.accdt" & _
                 " AND c.accseq   =  a.accseq" & _
                 " AND a.ptid     = d." & F_PTID & _
                 " AND b.empid    = c.vfyid "
        strSQL = strSQL & " AND   a.stscd = '" & enStsCd.StsCd_LIS_Cancel & "'"
    End If
                                        
    Set objProBar = New jProgressBar.clsProgress
    
    With objProBar
        .Container = Me
        .Width = ssCmtList.Width
        .Left = ssCmtList.Left
        .Top = ssCmtList.Top - 280
        .Height = 280
        .Message = "자료를 읽기 위해 준비중입니다..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = ssCmtList.Width
'        .XPos = ssCmtList.Left
'        .YPos = ssCmtList.Top - 280
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "자료를 읽기 위해 준비중입니다..."
'        .Value = 1
    End With
    
    Set rsGetinfo = New Recordset
    rsGetinfo.Open strSQL, DBConn
    
'    objProBar.Msg = ""
    
    
    If rsGetinfo.RecordCount > 0 Then
        objProBar.Max = rsGetinfo.RecordCount
    Else
        MsgBox "데이타가 없습니다.."
    End If
'    barStatus.Value = 0  GetDoctNm
    '.Fields("statfg").Value
    ClearssCmtList
    
    For i = 1 To rsGetinfo.RecordCount
'        barStatus.Value = barStatus.Value + 1
        objProBar.Value = i
        DoEvents
        With rsGetinfo
            If optOption(2).Value Then
                strDoct = GetDoctNm("" & .Fields("doctnm").Value)
            Else
                strDoct = "" & .Fields("doctnm").Value
            End If
            Call DspSpd("" & .Fields("WorkArea").Value, "" & .Fields("AccDt").Value, "" & .Fields("AccSeq").Value, _
                        "" & .Fields("PtId").Value, "" & .Fields("ptnm").Value, "" & .Fields("Sex").Value, "" & .Fields("AgeDay").Value, _
                        "" & .Fields("WardId").Value, "" & .Fields("RoomId").Value, "" & .Fields("rsttxt").Value, _
                        "" & .Fields("EmpNm").Value, "" & .Fields("RcvDt").Value, "" & .Fields("RcvTM").Value, strDoct, _
                        "" & .Fields("MFYID").Value, "" & .Fields("mfydt").Value, "" & .Fields("mfytm").Value, i)
        End With
      rsGetinfo.MoveNext
    Next i
    
'    MouseDefault   2001/04/18
    
    Set rsGetinfo = Nothing
    Set objProBar = Nothing
    
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    dtpStartDt.Value = Now
    dtpEndDt.Value = Now
    ClearssCmtList
    optOption(0).Value = True
End Sub

Private Sub DspSpd(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As Long, ByVal PTid As String, _
                   ByVal PtNm As String, ByVal Sex As String, ByVal AgeDay As String, ByVal WardId As String, _
                   ByVal RoomId As String, ByVal RstTxt As String, ByVal EmpNm As String, ByVal RcvDt As String, _
                   ByVal RcvTM As String, ByVal OrdDocT As String, _
                   ByVal MfyNm As String, ByVal MfyDt As String, ByVal MfyTm As String, ByVal lngRow As Long)
                    
    Dim sAge As String
    Dim Age As Integer
    Dim Location As String
    Dim tmpPtid As String
    Dim lngMaxHeight As Long
    
    Age = (AgeDay / 365) + 1
    sAge = Sex & "/" & CStr(Age)
    
    Location = WardId & "-" & RoomId    '& "-" & BedId
    

    With ssCmtList
        .MaxRows = lngRow
        .Row = lngRow
        .Col = 1: .Text = WorkArea & "-" & Mid(AccDt, 3) & "-" & AccSeq
        .Col = 2: .Text = Trim(PTid)
        .Col = 3: .Text = Trim(PtNm)
        .Col = 4: .Text = sAge
        .Col = 5: .Text = WardId & "-" & RoomId
        
        
        RstTxt = Replace(RstTxt, vbCrLf, " ")
        .Col = 6: .Text = Trim(RstTxt)
        If .MaxTextCellHeight > lngMaxHeight Then lngMaxHeight = .MaxTextCellHeight
        
        .Col = 7: .Text = Trim(EmpNm)
        .Col = 8: .Text = Format(Trim(RcvDt), CS_DateLongMask)
        .Col = 9: .Text = Format(Trim(RcvTM), CS_TimeLongMask)
        .Col = 10: .Text = Trim(OrdDocT)
        .Col = 11: .Text = GetEmpNm(Trim(MfyNm))
        .Col = 12: .Text = Format(Trim(MfyDt), CS_DateLongMask)
        .Col = 13: .Text = Format(Trim(MfyTm), CS_TimeLongMask)
        
        .RowHeight(lngRow) = lngMaxHeight
    End With
    
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


Private Sub AnalysisHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    strTmp = "AnalysisList"
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 8

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("AnalysisList", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    strTmp = "조회기간 : " & Format(dtpStartDt.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpEndDt.Value, "YYYY년 MM월 DD일")
    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
    
    strTmp = "조회조건 : "
    For ii = 0 To 2
        If optOption(ii).Value Then
            Select Case ii
                Case 0: strTmp = strTmp & "     " & "(√)모두"
                Case 1: strTmp = strTmp & "     " & "(√)결과수정사유 "
                Case 2: strTmp = strTmp & "     " & "(√)접수취소사유"
            End Select
        Else
            Select Case ii
                Case 0: strTmp = strTmp & "     " & "(  )모두"
                Case 1: strTmp = strTmp & "     " & "(  )결과수정사유 "
                Case 2: strTmp = strTmp & "     " & "(  )접수취소사유"
            End Select
        End If
    Next
    
    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    Call PrintString("접수번호", "환자ID", "환자명", "성별/나이", "병실", "소견사유", "입력자", "입력일", "처방의", "수정자", "수정일", "수정시간", True)
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub

Private Sub PrintString(ByVal sAccno As String, ByVal sPtid As String, ByVal sPtnm As String, ByVal sSexAge As String, ByVal sLocation As String, _
                        ByVal sMesg As String, ByVal sEntNm As String, ByVal sEntDt As String, ByVal sOrdDt As String, _
                        ByVal sMfyNm As String, ByVal sMfyDT As String, ByVal sMfyTM As String, Optional ByVal blnHead As Boolean = False)
    Dim arytmp()    As String
    Dim strTmp      As String
    Dim ii          As Integer
    
    
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call AnalysisHead
    End If
    
    Call Print_Setting(sAccno, PrtLeft, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sPtid, 25, LineSpace, 15, "L", "C", False)
    Call Print_Setting(sPtnm, 40, LineSpace, 55, "L", "C", False)
    Call Print_Setting(sSexAge, 55, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sLocation, 75, LineSpace, 20, "L", "C", False)
    Call Print_Setting(sEntNm, 150, LineSpace, 15, "L", "C", False)
    Call Print_Setting(sEntDt, 165, LineSpace, 15, "L", "C", False)
    Call Print_Setting(sOrdDt, 185, LineSpace, 10, "L", "C", Not blnHead)
    
    If Len(sMesg) = 1 Then
        sMesg = ""
    End If
    
    If sMesg <> "" Then
        If blnHead = True Then
            Call Print_Setting(sMesg, 95, LineSpace, 55, "L", "C")
        Else
            Printer.FontBold = True
            For ii = 1 To 5
                If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                    sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                    If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
                        sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
                    End If
                End If
            Next
            
            arytmp() = Split(Trim(sMesg), vbCrLf)
            For ii = LBound(arytmp) To UBound(arytmp)
                If lngCurYPos > Printer.ScaleHeight - 6 Then
                    Printer.NewPage
                    Call AnalysisHead
                End If
                Call Print_Setting(arytmp(ii), PrtLeft + Printer.TextWidth("소견사유 : "), LineSpace, 55, "L", "C")
            Next
            Printer.FontBold = False
            Printer.DrawStyle = 1: Printer.DrawWidth = 2
            Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    
    Dim strAccNo    As String
    Dim strPtId     As String
    Dim strPtNm     As String
    Dim strSEXAGE   As String
    Dim strLocation As String
    Dim strMesg     As String
    Dim strEntNm    As String
    Dim strEntDT    As String
    Dim strOrdDt    As String
    Dim strMfyNm    As String
    Dim strMfyDT    As String
    Dim strMfyTM    As String
    
    Dim ii As Integer
    If ssCmtList.DataRowCnt < 1 Then Exit Sub
    
    Call P_PrtSet
    Call AnalysisHead
    
    With ssCmtList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1:   strAccNo = .Value
            .Col = 2:   strPtId = .Value
            .Col = 3:   strPtNm = .Value
            .Col = 4:   strSEXAGE = .Value
            .Col = 5:   strLocation = .Value
            .Col = 6:   strMesg = .Value
            .Col = 7:   strEntNm = .Value
            .Col = 8:   strEntDT = .Value
            .Col = 10:   strOrdDt = .Value
            .Col = 11:   strMfyNm = .Value
            .Col = 12:   strMfyDT = .Value
            .Col = 13:   strMfyTM = .Value
            
            Call PrintString(strAccNo, strPtId, strPtNm, strSEXAGE, strLocation, strMesg, strEntNm, strEntDT, strOrdDt, strMfyNm, strMfyDT, strMfyTM)
        Next
    End With
    
    Printer.EndDoc
End Sub

Private Sub ssCmtList_Click(ByVal Col As Long, ByVal Row As Long)
    Static iSortOrder As Integer
    
    With ssCmtList
        If Row = 0 Then  'Sort...
            .Row = 0: .Col = Col
            .Row = -1: .Col = -1
            .SortBy = SortByRow
            .SortKey(1) = Col
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
            End If
            .Action = ActionSort
            Exit Sub
        End If
    End With
    
End Sub
