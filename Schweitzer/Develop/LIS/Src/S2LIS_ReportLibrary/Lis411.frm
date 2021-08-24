VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm411PCollectionList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FCEFE9&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   3900
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   495
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2880
      Left            =   90
      TabIndex        =   7
      Top             =   1005
      Width           =   10800
      Begin VB.Frame fraPrtOption 
         BackColor       =   &H00DBE6E6&
         Height          =   435
         Left            =   4620
         TabIndex        =   10
         Top             =   1260
         Width           =   2925
         Begin VB.OptionButton optBusiDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "병동(채혈)"
            Height          =   195
            Index           =   1
            Left            =   1470
            TabIndex        =   3
            Tag             =   "2"
            Top             =   165
            Width           =   1320
         End
         Begin VB.OptionButton optBusiDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외래(접수)"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Tag             =   "1"
            Top             =   165
            Width           =   1230
         End
      End
      Begin VB.ComboBox cboBuilding 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2220
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   735
         Width           =   2925
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   1350
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   57475075
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   315
         Left            =   2220
         TabIndex        =   5
         Top             =   2055
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   57475074
         UpDown          =   -1  'True
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   315
         Left            =   2220
         TabIndex        =   4
         Top             =   1725
         Width           =   1755
         _ExtentX        =   3096
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
         CalendarBackColor=   16707582
         Format          =   57475074
         UpDown          =   -1  'True
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   705
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   609
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
         Caption         =   "Delivery Location"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   270
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "부터"
         Height          =   180
         Left            =   4050
         TabIndex        =   9
         Top             =   1860
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "까지 채혈(접수)된 검체"
         Height          =   180
         Left            =   4050
         TabIndex        =   8
         Top             =   2145
         Width           =   1890
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "채혈리스트 출력 - 개별채혈"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1035
      TabIndex        =   11
      Top             =   585
      Width           =   4155
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0088A2B7&
      FillColor       =   &H00DEEEFE&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   90
      Shape           =   4  '둥근 사각형
      Top             =   480
      Width           =   6390
   End
End
Attribute VB_Name = "frm411PCollectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPWACoHeRow As Integer
Dim fWhich As Object
Dim iPageWidth As Integer
Dim iPageHeight As Integer
Dim iCurY As Integer
Dim DataExist As Boolean
Dim sLastDt As String
Dim sLastTm As String
Dim iRecordCount As Integer
Dim SvBuildNm As String
Dim SvColTime As String
Dim SvBuildCd As String

Const iCm = 567
Const iLineHeight = 10

Dim iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%

Public Event FormClose()

Private Sub cboBuilding_Click()
   Call LoadColTime
End Sub

Private Sub cmdExit_Click()
   Unload Me
'   Set frm411PCollectionList = Nothing

    RaiseEvent FormClose
End Sub

Private Sub cmdReport_Click()
    
    Dim SqlStmt As String
    
    If Format(dtpStartTime.Value, CS_TimeDbFormat) > Format(dtpEndTime.Value, CS_TimeDbFormat) Then
         MsgBox "시간 설정이 잘못 되었습니다.. 확인하시고 다시 출력하십시오."
         dtpEndTime.SetFocus
         Exit Sub
    End If
    
    If cboBuilding.ListIndex < 0 Then
         MsgBox "검체전달 장소(Delivery Location)를 선택하세요."
         cboBuilding.SetFocus
         Exit Sub
    End If
    
    Call PrtBody
    
   Dim tmpRs As Recordset
   
   SqlStmt = " select field1 as ColDate, field2 as ColTime " & _
             " from   " & T_LAB031 & _
             " where  " & DBW("cdindex", LC2_ColListTm, 2) & _
             " and    " & DBW("cdval1", ObjSysInfo.BuildingCd, 2) & _
             " and    " & DBW("cdval2", medGetP(cboBuilding.Text, 1, " "), 2)
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs.EOF Then
      DataExist = False
   Else
      DataExist = True
   End If
   
   Set tmpRs = Nothing
    
    If DataExist Then
      SqlStmt = " update " & T_LAB031 & _
                " set " & _
                            DBW("field1", Format(dtpDate.Value, CS_DateDbFormat), 3) & _
                            DBW("field2", sLastTm, 2) & _
                " where " & DBW("cdindex", LC2_ColListTm, 2) & _
                " and   " & DBW("cdval1", ObjSysInfo.BuildingCd, 2) & _
                " and   " & DBW("cdval2", medGetP(cboBuilding.Text, 1, " "), 2)
    Else
      SqlStmt = " insert into " & T_LAB031 & "(cdindex, cdval1, cdval2, field1, field2) " & _
                " values (" & _
                                DBV("cdindex", LC2_ColListTm, 1) & _
                                DBV("cdval1", ObjSysInfo.BuildingCd, 1) & _
                                DBV("cdval2", medGetP(cboBuilding.Text, 1, " "), 1) & _
                                DBV("field1", Format(dtpDate.Value, CS_DateDbFormat), 1) & _
                                DBV("field2", sLastTm) & ") "
    End If
    
    DBConn.BeginTrans
    DBConn.Execute (SqlStmt)
    DBConn.CommitTrans
    
End Sub

Private Sub Form_Load()
   optBusiDiv(0).Value = True
   Call LoadBuildingList
   dtpDate.Value = Now
End Sub

Public Sub LoadColTime()

   Dim tmpRs As Recordset
   Dim SqlStmt As String
   
   SqlStmt = " select field1 as ColDate, field2 as ColTime " & _
             " from   " & T_LAB031 & _
             " where  " & DBW("cdindex", LC2_ColListTm, 2) & _
             " and    " & DBW("cdval1", ObjSysInfo.BuildingCd, 2) & _
             " and    " & DBW("cdval2", medGetP(cboBuilding.Text, 1, " "), 2)
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   If tmpRs.EOF Then
      dtpDate.Value = Format(Now, "yyyy-mm-dd")
      dtpStartTime.Value = Format(DateAdd("h", -2, Now), "hh:mm:ss")    '2시간 전
      dtpEndTime.Value = Format(Now, "hh:mm:ss")
      DataExist = False
   Else
      dtpDate.Value = Format(Trim(tmpRs.Fields("ColDate").Value), "####-##-##")
      dtpStartTime.Value = Format(Trim(tmpRs.Fields("ColTime").Value), "##:##:##")    '최근 마지막 시간
      dtpEndTime.Value = Format(Now, "hh:mm:ss")
      DataExist = True
   End If
   
   sLastDt = Format(dtpDate.Value, CS_DateDbFormat)
   sLastTm = Format(dtpStartTime.Value, CS_TimeDbFormat)
   
   Set tmpRs = Nothing
End Sub

Public Sub LoadBuildingList()

   Dim i As Integer
   Dim SqlStmt As String
   Dim tmpRs As Recordset
   
   SqlStmt = " Select cdval1 as BuildCd, field1 as BuildNm " & _
             " from   " & T_LAB032 & _
             " where  " & DBW("cdindex", LC3_Buildings, 2) & _
             " order by BuildCd "
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   
   cboBuilding.Clear
   For i = 1 To tmpRs.RecordCount
      cboBuilding.AddItem Trim("" & tmpRs.Fields("BuildCd").Value) & "   " & Trim("" & tmpRs.Fields("BuildNm").Value)
      tmpRs.MoveNext
   Next
   
   Set tmpRs = Nothing
End Sub

Public Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
End Sub

Public Sub PrtHeader()
   
    Dim Title As String
    
    '/* 보고서 제목
    Title = "채혈 리스트"
    
    
    '        "채혈 리스트"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm * 1.5)
    
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 2
    iposWorkNo = iposSEQ + iCm - iCm / 4
    iposPtName = iposWorkNo + 2.3 * iCm
    'iposPtID = iposPtName + iCm + iCm / 2
    iposPtID = (iposPtName + iCm + iCm / 2) - 200
    iposSAge = (iposPtID + iCm + iCm / 2) + 100
    'iposSAge = iposPtID + iCm + iCm / 2
    iposIO = iposSAge + iCm
    iposRcv = iposIO + iCm
    iposSF = iposRcv + 2.5 * iCm
    iposTestCD = iposSF + 0.8 * iCm
    iposSpccd = (iposTestCD + 7.5 * iCm) - 300
    'iposSpccd = iposTestCD + 7.5 * iCm
    
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposWorkNo, "    Work No", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "  환자ID", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    Call WriteStr(iCurY, iposIO, "  I/O", iCurY, 0)
    Call WriteStr(iCurY, iposRcv, "      채혈일시", iCurY, 0)
    Call WriteStr(iCurY, iposSF, " S/F", iCurY, 0)
    Call WriteStr(iCurY, iposTestCD, "                              검사종목", iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "  검체", iCurY, iCm / 2)
    
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, 0)
    
 
End Sub

Public Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str
End Sub

Public Sub prtTitle(Title As String, iSpace As Integer)

    Dim oldFontSize As Integer
    
    oldFontSize = Printer.FontSize
    Printer.FontSize = 14
    Printer.FontBold = True
    '/* Tile이 중앙으로 오도록 string길이에 따라 위치를 계산한다.
    
    Printer.CurrentY = 0
    Printer.CurrentX = iPageWidth / 2 - Printer.TextWidth(Title) / 2
    iCurY = Printer.CurrentY + Printer.TextHeight(Title) + iSpace
    
    Printer.Print Title
    Printer.FontSize = oldFontSize
    Printer.FontBold = False

End Sub

Public Sub ChangeLine(iLineSpace As Integer)
    iCurY = iCurY + iLineSpace
    Printer.CurrentY = iCurY
    Printer.CurrentX = iCm / 2
    
End Sub

Public Sub PrtBody()

    Dim sSQL1 As String
    Dim sSQL2 As String
    Dim rsWorksheet As Recordset
    Dim rsPtName As Recordset
    
    Dim sStart As String, sEnd As String, tmpStr As String, sStsCd As String
    Dim i%
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sStartTime As String
    Dim sEndTime As String
    
    'If cboBuilding.ListIndex = 0 Then
    '  tmpStr = ""
    'Else
    '  tmpStr = "   a.buildcd = '" & medGetP(cboBuilding.Text, 1, " ") & "' and "
    'End If
   
    Printer.FontSize = 9
   
    sStartDate = Format(dtpDate.Value, CS_DateDbFormat)
    sStartTime = Format(dtpStartTime.Value, CS_TimeDbFormat)
    sEndTime = Format(dtpEndTime.Value, CS_TimeDbFormat)

    sStart = sStartDate & sStartTime
    sEnd = sEndDate & sEndTime
   
    If optBusiDiv(0).Value Then
         sStsCd = enStsCd.StsCd_LIS_Accession    '외래-->접수상태
    Else
         sStsCd = enStsCd.StsCd_LIS_Collection   '병동-->채혈상태
    End If



    sSQL1 = " select a.workarea, a.accdt, a.accseq, a.ptid, a.sex, a.ageday, a.deptcd, a.coldt, " & _
            " a.coltm, a.storecd, a.spccd, a.testdiv, a.buildcd, b.field1 as BuildNm, c.field3 as SpcNm " & _
            " from  " & T_LAB201 & " a, " & T_LAB032 & " b, " & T_LAB032 & " c," & T_LAB204 & " d" & _
            " where " & DBW("a.orgbuildcd", ObjSysInfo.BuildingCd, 2) & _
            " and   " & DBW("a.coldt", sStartDate, 2) & _
            " and   " & DBW("a.coltm >= ", sStartTime) & _
            " and   " & DBW("a.stscd", sStsCd, 2) & _
            " and   " & DBW("a.buildcd", medGetP(cboBuilding.Text, 1, " "), 2) & _
            " and   " & DBW("a.coltm <= ", sEndTime) & _
            " and   "
    
    If optBusiDiv(0).Value Then
        sSQL1 = sSQL1 & "  ( a.wardid = '' or a.wardid is null ) and "
    Else
        sSQL1 = sSQL1 & "  a.wardid <> ' ' and  a.wardid is not null  and "
    End If
    
    sSQL1 = sSQL1 & _
                        DBW("b.cdindex", LC3_Buildings, 2) & _
            " and   b.cdval1 = a.buildcd" & _
            " and   " & DBW("c.cdindex", LC3_Specimen, 2) & _
            " and   c.cdval1 = a.spccd " & _
            " and a.coldt=d.workdt " & _
            " and substring(a.coltm,1,4)=substring(d.worktm,1,4) " & _
            " and a.workarea=d.workarea " & _
            " and a.accdt=d.accdt " & _
            " and a.accseq=d.accseq " & _
            " order by a.buildcd, a.rcvdt, a.rcvtm"
            '" and   " & _
            "       not exists (select * from " & T_LAB204 & " where workdt = a.coldt and " & _
            "                   worktm = a.coltm and workarea = a.workarea and accdt = a.accdt and accseq = a.accseq) " & _
            " order by a.buildcd, a.rcvdt, a.rcvtm "  '접수시간순으로.. 2000/2/17 by 김미경&정미경
            '" order by a.buildcd, a.workarea, a.accdt, a.accseq "

    Set rsWorksheet = New Recordset
    rsWorksheet.Open sSQL1, DBConn
   
    If rsWorksheet.EOF = True Then ' record가 존재하지 않을경우
       MsgBox " 데이타가 존재하지 않습니다. "
       GoTo Nodata
    End If
    
    iRecordCount = rsWorksheet.RecordCount
    
    SvBuildCd = Trim(rsWorksheet.Fields("buildcd").Value)
    SvBuildNm = Trim(rsWorksheet.Fields("buildnm").Value)
    
    Call InitReport
    Call PrtHeader
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark
    
    Dim temp1 As String, temp2 As String
    Dim sAge As String
    Dim iSeqNum As Integer
    
    sLastTm = ""
    With rsWorksheet
    
        .MoveFirst
        For i = 1 To .RecordCount
            
            If sLastTm < .Fields("ColTm").Value Then sLastTm = .Fields("ColTm").Value
                
            temp1 = Mid(.Fields("ColTm").Value, 1, 4)
            temp2 = Format(temp1, "00:00")
            
            sSQL2 = " select " & F_PTNM & "  as ptnm " & _
                    " from   " & T_HIS001 & _
                    " where  " & DBW(F_PTID, rsWorksheet.Fields("ptid").Value, 2)
            Set rsPtName = New Recordset
            rsPtName.Open sSQL2, DBConn

            If chkTestCD(.Fields("WorkArea").Value, .Fields("accdt").Value, .Fields("accseq").Value, _
                         .Fields("testdiv").Value) = True Then                   ' Exists
                Call ChangeLine(iCm / 2)
                
               If SvBuildCd <> Trim(.Fields("buildcd").Value) Then
                    SvBuildCd = Trim(.Fields("buildcd").Value)
                    SvBuildNm = Trim(.Fields("buildnm").Value)
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
               End If
                
                If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                    Call Print_WaterMark
                End If
            
                iSeqNum = iSeqNum + 1
            
                sAge = (.Fields("AgeDay").Value \ 365) + 1
                Call WriteStr(iCurY, iposSEQ + iCm / 6, CStr(iSeqNum), iCurY, 0)
                Call WriteStr(iCurY, iposWorkNo + iCm / 6, Trim(CStr(.Fields("WorkArea").Value)) & "-" & Del2Chr(.Fields("AccDt").Value) & _
                              "-" & Trim(CStr(.Fields("AccSeq").Value)), iCurY, 0)
                Call WriteStr(iCurY, iposPtName + iCm / 6, rsPtName.Fields("ptnm").Value, iCurY, 0)

                Call WriteStr(iCurY, iposPtID + iCm / 6, (CStr(.Fields("PtId").Value)), iCurY, 0)
                
                Call WriteStr(iCurY, iposSAge + iCm / 6, Trim(.Fields("Sex").Value) & "/" & Trim(CStr(sAge)), iCurY, 0)
                Call WriteStr(iCurY, iposIO + iCm / 6, Trim(.Fields("DeptCd").Value), iCurY, 0)
                Call WriteStr(iCurY, iposRcv + iCm / 6, Del2Chr(.Fields("ColDt").Value) & _
                              "   " & temp2, iCurY, 0)
                Call WriteStr(iCurY, iposSF + iCm / 6, Trim(.Fields("StoreCd").Value), iCurY, 0)
                Call WriteStr(iCurY, iposSpccd + iCm / 6, Trim(.Fields("SpcNm").Value), iCurY, 0)
            
                Call WriteTestCD(.Fields("WorkArea").Value, .Fields("AccDt").Value, .Fields("AccSeq").Value, .Fields("TestDiv").Value)
            End If
            .MoveNext
        Next i
        Set rsPtName = Nothing
    End With
        
    Call Print_WaterMark
    Printer.EndDoc
    
Nodata:
    Set rsWorksheet = Nothing
End Sub

Private Function Del2Chr(sStr As String) As String
    Del2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function chkTestCD(sWorkarea As String, sAccDt As String, sAccSeq As String, stestdiv As String) As Boolean
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    
    sSQL2 = " select ordcd " & _
            " from   " & T_LAB102 & _
            " where  " & DBW("workarea", sWorkarea, 2) & _
            " and    " & DBW("accdt", Trim(sAccDt), 2) & _
            " and    " & DBW("accseq", Trim(sAccSeq), 2)
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        chkTestCD = False         ' not exitst
    Else
        chkTestCD = True              ' Exist
    End If
    
    Set rsTestCode = Nothing
End Function

Public Sub WriteTestCD(sWorkarea As String, sAccDt As String, sAccSeq As String, stestdiv As String)
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    Dim i%, tmpiposTestCD
    Dim sTable As String
    
    sSQL2 = " select ordcd " & _
            " from  " & T_LAB102 & _
            " where " & DBW("workarea", sWorkarea, 2) & _
            " and   " & DBW("accdt", Trim(sAccDt), 2) & _
            " and   " & DBW("accseq", Trim(sAccSeq), 2)
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
'        MsgBox " 레코드가 없다"
        GoTo Nodata
    End If
    
    With rsTestCode
        tmpiposTestCD = iposTestCD
        rsTestCode.MoveFirst
        For i = 1 To rsTestCode.RecordCount

            Call WriteStr(iCurY, tmpiposTestCD + iCm / 6, Trim(.Fields("OrdCd").Value), iCurY, 0)
            tmpiposTestCD = tmpiposTestCD + 1.5 * iCm
            If (i Mod 5 = 0) Then
                Call ChangeLine(iCm / 2)
                If iCurY > iPageHeight - 2 * iCm Then  ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                End If
                tmpiposTestCD = iposTestCD
            End If
            rsTestCode.MoveNext
        Next i
        If (rsTestCode.RecordCount Mod 5) = 0 Then
            iCurY = iCurY - iCm / 2
        End If
    End With
    
Nodata:
    Set rsTestCode = Nothing
End Sub

Public Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select
         
    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    iCurY = Printer.CurrentY + iSpace
End Sub

Public Sub prtPageNum()
    
    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    
    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
    
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + iCm / 6
    Printer.Print "RUN-TIME : " & sTime
        
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
End Sub

Public Sub prtTerm()
    Dim oldX As Integer, oldY As Integer
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sStartTime As String
    Dim sEndTime As String
    Dim oldFontSize As Integer
    
    
    sStartDate = Format(dtpDate.Value, CS_DateLongFormat)
    sStartTime = Format(dtpStartTime.Value, "HH:MM")
    sEndTime = Format(dtpEndTime.Value, "HH:MM")

     
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    oldFontSize = Printer.FontSize
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "Delivery Location  : " & SvBuildCd & "   " & SvBuildNm
    Printer.CurrentX = iCm * 8
    Printer.CurrentY = 1.3 * iCm
    If optBusiDiv(0).Value Then
         Printer.Print "( 외래검체 - 접수상태 )"
    Else
         Printer.Print "( 병동검체 - 채혈상태 )"
    End If
    
    
    Printer.FontSize = oldFontSize
    Printer.FontBold = False
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("Work Area : ") + iCm / 4
    Printer.Print "채혈장소   : " & ObjSysInfo.BuildingCd & "    " & ObjSysInfo.BuildingNm & "    /    " & "채혈일시   : " & sStartDate & "    " & sStartTime & "  ~  " & sEndTime
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("Work Area : ") + iCm / 4
    Printer.FontBold = True
    Printer.Print "검체수   :   총  " & CStr(iRecordCount) & "  개"
    Printer.FontBold = False
                                    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
End Sub
