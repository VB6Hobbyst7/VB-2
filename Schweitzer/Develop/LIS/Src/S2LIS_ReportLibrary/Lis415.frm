VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form frm415PAccessionList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FCEFE9&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2865
      Left            =   90
      TabIndex        =   7
      Top             =   1005
      Width           =   10755
      Begin VB.OptionButton optOrderBy 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수번호순"
         Height          =   315
         Index           =   1
         Left            =   7830
         TabIndex        =   14
         Top             =   765
         Width           =   1290
      End
      Begin VB.OptionButton optOrderBy 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수시간순"
         Height          =   315
         Index           =   0
         Left            =   6390
         TabIndex        =   13
         Top             =   735
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.CommandButton CmdWACodeHelp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         Height          =   360
         Left            =   2595
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   705
         Width           =   315
      End
      Begin VB.TextBox txtWACode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   360
         HideSelection   =   0   'False
         Left            =   1875
         MaxLength       =   2
         TabIndex        =   0
         Top             =   705
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpEndTime1 
         Height          =   345
         Left            =   3660
         TabIndex        =   4
         Top             =   2070
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         Format          =   21299203
         CurrentDate     =   36370.1878472222
      End
      Begin MSComCtl2.DTPicker dtpEndDate1 
         Height          =   345
         Left            =   1875
         TabIndex        =   3
         Top             =   2070
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   609
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
         Format          =   21299203
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker dtpStartTime1 
         Height          =   345
         Left            =   3645
         TabIndex        =   2
         Top             =   1590
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         Format          =   21299203
         UpDown          =   -1  'True
         CurrentDate     =   36370.0208333333
      End
      Begin MSComCtl2.DTPicker dtpStartDate1 
         Height          =   345
         Left            =   1875
         TabIndex        =   1
         Top             =   1590
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   609
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
         Format          =   21299203
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   705
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Work Area"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   270
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label lblWAName 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   2925
         TabIndex        =   10
         Top             =   705
         Width           =   2940
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DBE6E6&
         Caption         =   " ~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2085
         Width           =   375
      End
   End
   Begin VB.Frame frmWACodeHelp 
      BorderStyle     =   0  '없음
      Height          =   2205
      Left            =   2010
      TabIndex        =   5
      Top             =   2205
      Visible         =   0   'False
      Width           =   2955
      Begin FPSpread.vaSpread spdWACodeHelp 
         Height          =   2205
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   2925
         _Version        =   196608
         _ExtentX        =   5159
         _ExtentY        =   3889
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   8
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis415.frx":0000
         UserResize      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "접수대장 출력"
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
      Left            =   1500
      TabIndex        =   12
      Top             =   480
      Width           =   2085
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0088A2B7&
      FillColor       =   &H00DEEEFE&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   150
      Shape           =   4  '둥근 사각형
      Top             =   390
      Width           =   5220
   End
End
Attribute VB_Name = "frm415PAccessionList"
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

Const iCm = 567
Const iLineHeight = 10

Dim iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%


Public Event FormClose()

Private Sub cmdExit_Click()
   Unload Me
'   Set frm415PAccessionList = Nothing

    RaiseEvent FormClose
End Sub

Private Sub CmdWACodeHelp_Click()
    Dim i%
    Dim sSQL As String
    Dim dsToWACode As Recordset

    sSQL = " select cdval1, field1 " & _
           " from   " & T_LAB032 & _
           " where  " & DBW("cdindex", LC3_WorkArea, 2)

    Set dsToWACode = New Recordset
    dsToWACode.Open sSQL, DBConn

    If dsToWACode.EOF = True Then ' record가 존재하지 않을 경우
        Exit Sub
    End If

    With spdWACodeHelp
        dsToWACode.MoveFirst
        .MaxRows = dsToWACode.RecordCount

        For i = 1 To dsToWACode.RecordCount
            .Row = i

            .Col = 1
            .Text = "" & dsToWACode.Fields("cdval1").Value

            .Col = 2
            .Text = "" & dsToWACode.Fields("field1").Value

            dsToWACode.MoveNext
        Next i
    End With

    frmWACodeHelp.Visible = True
    frmWACodeHelp.ZOrder 0
    Set dsToWACode = Nothing
End Sub

Private Sub cmdReport_Click()
    
    Call PrtBody
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        frmWACodeHelp.Visible = False
       ' caldate.Visible = False
        
    End If
End Sub


Private Sub Form_Load()
    dtpStartDate1.Value = Format(Now, "yyyy-mm-dd")
    dtpStartTime1.Value = "00:00:00"
    dtpEndDate1.Value = Format(Now, "yyyy-mm-dd")
    dtpEndTime1.Value = Format(Now, "hh:mm:ss")
End Sub

Private Sub spdWACodeHelp_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or iPWACoHeRow = Row Then Exit Sub
    
    With spdWACodeHelp
        .Row = Row
        
        .Col = 1
        txtWACode.Text = .Text
        
        .Col = 2
        lblWAName.Caption = .Text
    End With
    
    iPWACoHeRow = -1
    frmWACodeHelp.Visible = False
End Sub

Private Sub txtWACode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sSQL As String, tmp As String
    Dim dsSelWACode As Recordset

    If KeyCode = vbKeyReturn Then
        
        lblWAName.Caption = ""
        sSQL = " select cdval1 , field1 from " & T_LAB032 & _
               " where " & DBW("cdindex", LC3_WorkArea, 2) & _
               " and   " & DBW("cdval1", UCase(txtWACode.Text), 2)
        
        Set dsSelWACode = New Recordset
        dsSelWACode.Open sSQL, DBConn
        
        If dsSelWACode.EOF = True Then
'            medMain.stsBar.Panels(2).Text = " 존재하지 않는 WorkArea Code 입니다"
            txtWACode.Text = ""
            txtWACode.SetFocus

        Else
            txtWACode.Text = UCase(txtWACode.Text)
'            medMain.stsBar.Panels(2).Text = ""
            lblWAName.Caption = dsSelWACode.Fields("field1").Value
           
        End If
        
        Set dsSelWACode = Nothing
    End If
End Sub

Public Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
End Sub

Public Sub PrtHeader()
   
    Dim Title As String
    
    '/* 보고서 제목
    Title = "접수대장"
    
    
    '        "접수대장"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm * 1.5)
    
    ' -----------------------------------------------------------------------------
    Call DrawLine(0, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = 0 ' iCm / 2
    iposWorkNo = iposSEQ + iCm
    iposPtName = iposWorkNo + 2 * iCm
    'iposPtID = iposPtName + iCm + iCm / 2
    iposPtID = (iposPtName + iCm + iCm / 2) '- 200
    'iposSAge = iposPtID + iCm + iCm / 2
    iposSAge = (iposPtID + iCm + iCm / 2) + 100
    iposIO = (iposSAge + iCm)
    iposRcv = iposIO + iCm
    'iposSF = iposRcv + 3 * iCm
    iposSF = iposRcv + 2.5 * iCm
    'iposTestCD = iposSF + 0.8 * iCm
    iposTestCD = iposSF + 0.5 * iCm
    'iposSpccd = iposTestCD + 7.5 * iCm
    iposSpccd = iposTestCD + 7.2 * iCm
    
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposWorkNo, "    Work No", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "  환자ID", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    Call WriteStr(iCurY, iposIO, "  I/O", iCurY, 0)
    Call WriteStr(iCurY, iposRcv, "  검체도착시간", iCurY, 0)
    Call WriteStr(iCurY, iposSF, " S/F", iCurY, 0)
    Call WriteStr(iCurY, iposTestCD, "                           검사종목", iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "검체", iCurY, iCm / 2)
    
    Call DrawLine(0, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, 0)
    
 
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
    Printer.FontSize = 9    'oldFontSize
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
    
    Dim sStart As String, sEnd As String
    Dim i%
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sStartTime As String
    Dim sEndTime As String
    

    sStartDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
    sStartTime = Format(dtpStartTime1.Value, CS_TimeDbFormat)
    sEndDate = Format(dtpEndDate1.Value, CS_DateDbFormat)
    sEndTime = Format(dtpEndTime1.Value, CS_TimeDbFormat)

    sStart = sStartDate & sStartTime
    sEnd = sEndDate & sEndTime
   
    If sStartDate = sEndDate Then
        sSQL1 = " select accdt, accseq, ptid, sex, ageday, deptcd, rcvdt, " & _
                " rcvtm, storecd, spccd, testdiv " & _
                " from   " & T_LAB201 & _
                " where  " & DBW("workarea", txtWACode.Text, 2) & _
                " and    " & DBW("rcvdt = ", sStartDate) & _
                " and    " & DBW("rcvtm >= ", sStartTime) & _
                " and    " & DBW("rcvtm <= ", sEndTime) & _
                " and    " & DBW("buildcd", ObjSysInfo.BuildingCd, 2)
    Else
        sSQL1 = " select accdt, accseq, ptid, sex, ageday, deptcd, rcvdt, " & _
                " rcvtm, storecd, spccd, testdiv " & _
                " from   " & T_LAB201 & _
                " where  " & DBW("workarea", txtWACode.Text, 2) & _
                " and    " & DBW("rcvdt = ", sStartDate) & _
                " and    " & DBW("rcvtm >= ", sStartTime) & _
                " and    " & DBW("buildcd", ObjSysInfo.BuildingCd, 2)
        sSQL1 = sSQL1 & " union all " & _
                " select accdt, accseq, ptid, sex, ageday, deptcd, rcvdt, " & _
                " rcvtm, storecd, spccd, testdiv " & _
                " from   " & T_LAB201 & _
                " where  " & DBW("workarea", txtWACode.Text, 2) & _
                " and    " & DBW("rcvdt > ", sStartDate) & _
                " and    " & DBW("rcvdt < ", sEndDate) & _
                " and    " & DBW("buildcd", ObjSysInfo.BuildingCd, 2)
        sSQL1 = sSQL1 & " union all " & _
                " select accdt, accseq, ptid, sex, ageday, deptcd, rcvdt, " & _
                " rcvtm, storecd, spccd, testdiv " & _
                " from   " & T_LAB201 & _
                " where  " & DBW("workarea", txtWACode.Text, 2) & _
                " and    " & DBW("rcvdt = ", sEndDate) & _
                " and    " & DBW("rcvtm <= ", sEndTime) & _
                " and    " & DBW("buildcd", ObjSysInfo.BuildingCd, 2)
    End If

    If optOrderBy(0).Value Then
        sSQL1 = sSQL1 & _
                " order by rcvdt, rcvtm "   '접수시간순
    Else
        sSQL1 = sSQL1 & _
                " order by accdt, accseq"   '접수번호순
    End If
    
    Set rsWorksheet = New Recordset
    rsWorksheet.Open sSQL1, DBConn
   
    If rsWorksheet.EOF = True Then ' record가 존재하지 않을경우
       MsgBox " 전체레코드가 존재하지 않습니다. "
       Exit Sub
    End If
    
    Call InitReport
    Call PrtHeader
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark
    
    Dim temp1       As String
    Dim temp2       As String
    Dim sAge        As String
    Dim iSeqNum     As Integer
    Dim sICSString  As String
    
    With rsWorksheet
        .MoveFirst
        For i = 1 To .RecordCount
            sSQL2 = " select " & F_PTNM & " as ptnm from " & T_HIS001 & _
                    " where  " & DBW(F_PTID, rsWorksheet.Fields("ptid").Value, 2)

            Set rsPtName = Nothing
            Set rsPtName = New Recordset
            rsPtName.Open sSQL2, DBConn

            sICSString = ICSPatientString(rsWorksheet.Fields("ptid").Value & "", enICSNum.LIS_ALL)

            If chkTestCD(.Fields("accdt").Value & "", .Fields("accseq").Value & "", _
                         .Fields("testdiv").Value & "") = True Then                  ' Exists

                Call ChangeLine(iCm / 2)
                If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                    Call Print_WaterMark
                End If

                iSeqNum = iSeqNum + 1

                sAge = (Abs(.Fields("AgeDay").Value & "") \ 365) + 1
                Call WriteStr(iCurY, iposSEQ + iCm / 6, CStr(iSeqNum), iCurY, 0)
                If sICSString <> "" Then
                    Call WriteStr(iCurY, iposWorkNo + iCm / 6, "Infection : " & sICSString, iCurY, 0)
                    Call ChangeLine(iCm / 2)
                End If

                Call WriteStr(iCurY, iposWorkNo + iCm / 6, Del2Chr(.Fields("AccDt").Value & "") & _
                              "-" & Trim(CStr(.Fields("AccSeq").Value & "")), iCurY, 0)
                Call WriteStr(iCurY, iposPtName + iCm / 6, rsPtName.Fields("ptnm").Value & "", iCurY, 0)

                Call WriteStr(iCurY, iposPtID + iCm / 6, CStr(.Fields("PtId").Value & ""), iCurY, 0)
                Call WriteStr(iCurY, iposSAge + iCm / 6, Trim(.Fields("Sex").Value & "") & "/" & Trim(CStr(sAge)), iCurY, 0)
                Call WriteStr(iCurY, iposIO + iCm / 6, Trim(.Fields("DeptCd").Value & ""), iCurY, 0)

                temp1 = Mid(.Fields("RcvTm").Value, 1, 4)
                temp2 = Format(temp1, "00:00")
                Call WriteStr(iCurY, iposRcv + iCm / 6, Del2Chr(.Fields("RcvDt").Value & "") & _
                              "   " & temp2, iCurY, 0)

                Call WriteStr(iCurY, iposSF + iCm / 6, Trim(.Fields("StoreCd").Value & ""), iCurY, 0)

                Call WriteStr(iCurY, iposSpccd + iCm / 6, Trim(.Fields("SpcCd").Value & ""), iCurY, 0)

                Call WriteTestCD(.Fields("AccDt").Value & "", .Fields("AccSeq").Value & "", .Fields("TestDiv").Value & "")
            End If
            .MoveNext
        Next i
        Set rsPtName = Nothing
    End With
        
    Call Print_WaterMark
    Printer.EndDoc
    
    Set rsWorksheet = Nothing
End Sub

Private Function Del2Chr(sStr As String) As String
    Del2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function chkTestCD(sAccDt As String, sAccSeq As String, stestdiv As String) As Boolean
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
'    Select Case stestdiv
'        Case "0"                ' 일반
            sSQL2 = " select a.testcd, b.abbrnm5 " & _
                    " from  " & T_LAB302 & " a, " & T_LAB001 & " b " & _
                    " where " & DBW("a.workarea", txtWACode.Text, 2) & _
                    " and   " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and   " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and   (a.detailfg = '' or a.detailfg is null or a.rstdiv = '*') " & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = ( select max(applydt) from " & T_LAB001 & _
                                         " where testcd = b.testcd ) "
                
'        Case "1"                ' 기타
            sSQL2 = sSQL2 & " union all " & _
                    " select a.testcd, b.abbrnm5 " & _
                    " from  " & T_LAB351 & " a, " & T_LAB001 & " b " & _
                    " where " & DBW("a.workarea", txtWACode.Text, 2) & _
                    " and   " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and   " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = ( select max(applydt) from " & T_LAB001 & _
                                         " where testcd = b.testcd ) "

'        Case "2"                ' 미생물
            sSQL2 = sSQL2 & " union all " & _
                    " select a.testcd, b.abbrnm5 " & _
                    " from   " & T_LAB404 & " a, " & T_LAB001 & " b " & _
                    " where  " & DBW("a.workarea", txtWACode.Text, 2) & _
                    " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and    (a.detailfg = '' or  a.detailfg is null or a.rstdiv = '*' )  " & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = ( select max(applydt) from " & T_LAB001 & _
                                         " where testcd = b.testcd ) "
                
'    End Select
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        chkTestCD = False         ' not exitst
        Exit Function
    End If
    chkTestCD = True              ' Exist
    
    Set rsTestCode = Nothing
End Function

Public Sub WriteTestCD(sAccDt As String, sAccSeq As String, stestdiv As String)
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    Dim i%, tmpiposTestCD
    Dim sTable As String
    
'    Select Case stestdiv
'        Case "0"                ' 일반
            sSQL2 = " select a.testcd, b.abbrnm5 " & _
                    " from   " & T_LAB302 & " a, " & T_LAB001 & " b " & _
                    " where  " & DBW("a.workarea", txtWACode.Text, 2) & _
                    " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and    (a.detailfg = '' or a.detailfg is null or a.rstdiv = '*' )" & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = ( select max(applydt) from " & T_LAB001 & _
                                         " where testcd = b.testcd ) "
'        Case "1"                ' 기타
            sSQL2 = sSQL2 & " union all " & _
                    " select a.testcd, b.abbrnm5 " & _
                    " from   " & T_LAB351 & " a, " & T_LAB001 & " b " & _
                    " where  " & DBW("a.workarea", txtWACode.Text, 2) & _
                    " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = ( select max(applydt) from " & T_LAB001 & _
                                         " where testcd = b.testcd ) "
'        Case "2"                ' 미생물
            sSQL2 = sSQL2 & " union all " & _
                    " select a.testcd, b.abbrnm5 " & _
                    " from   " & T_LAB404 & " a, " & T_LAB001 & " b " & _
                    " where  " & DBW("a.workarea", txtWACode.Text, 2) & _
                    " and    " & DBW("a.accdt", Trim(sAccDt), 2) & _
                    " and    " & DBW("a.accseq", Trim(sAccSeq), 2) & _
                    " and    (a.detailfg = '' or a.detailfg is null or  a.rstdiv = '*' )" & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = ( select max(applydt) from " & T_LAB001 & _
                                         " where testcd = b.testcd ) "
'    End Select
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        Exit Sub
    End If
    
    With rsTestCode
        tmpiposTestCD = iposTestCD
        rsTestCode.MoveFirst
        For i = 1 To rsTestCode.RecordCount
'
            'Call WriteStr(iCurY, tmpiposTestCD + iCm / 6, Trim(.Fields("TestCd").Value), iCurY, 0)
            Call WriteStr(iCurY, tmpiposTestCD + iCm / 6, Trim(.Fields("abbrnm5").Value & ""), iCurY, 0)
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
    
    sStartDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
    sStartTime = Format(dtpStartTime1.Value, "HH:MM")
    sEndDate = Format(dtpEndDate1.Value, CS_DateDbFormat)
    sEndTime = Format(dtpEndTime1.Value, "HH:MM")

     
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    Printer.CurrentX = iCm / 2
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "Work Area  : " & lblWAName.Caption
    
    Printer.CurrentX = iCm / 2
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("Work Area : ") + iCm / 6
    Printer.Print "접수기간   : " & sStartDate & "    " & sStartTime & "  ~  " & _
                                    sEndDate & "    " & sEndTime
                                    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
End Sub
