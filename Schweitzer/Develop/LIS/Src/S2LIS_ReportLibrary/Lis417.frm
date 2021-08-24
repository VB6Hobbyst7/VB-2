VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form frm417PModifiedList 
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
      TabIndex        =   11
      Top             =   3900
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Frame frmWACodeHelp 
      BorderStyle     =   0  '없음
      Height          =   2205
      Left            =   1875
      TabIndex        =   2
      Top             =   2025
      Visible         =   0   'False
      Width           =   2955
      Begin FPSpread.vaSpread spdWACodeHelp 
         Height          =   2205
         Left            =   -30
         TabIndex        =   3
         Top             =   15
         Width           =   2955
         _Version        =   196608
         _ExtentX        =   5212
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
         SpreadDesigner  =   "Lis417.frx":0000
         UserResize      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2895
      Left            =   75
      TabIndex        =   4
      Top             =   960
      Width           =   10770
      Begin VB.CommandButton CmdWACodeHelp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         Height          =   360
         Left            =   3585
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox txtWACode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   360
         HideSelection   =   0   'False
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Top             =   705
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker dtpStartDate1 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   1650
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
         Format          =   58195971
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   2
         Left            =   195
         TabIndex        =   9
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
         Index           =   3
         Left            =   195
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1635
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
         Height          =   345
         Left            =   3915
         TabIndex        =   6
         Top             =   705
         Width           =   2940
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "결과수정자/정보입력자 출력"
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
      Left            =   945
      TabIndex        =   8
      Top             =   525
      Width           =   4155
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0088A2B7&
      FillColor       =   &H00DEEEFE&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   405
      Width           =   5850
   End
End
Attribute VB_Name = "frm417PModifiedList"
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

Dim sStartTime As String
Dim sEndTime As String

Public Event FormClose()


Private Sub cmdExit_Click()
   Unload Me
'   Set frm902ModifiedList = Nothing

    RaiseEvent FormClose
End Sub

Private Sub CmdWACodeHelp_Click()
    
    Dim i%
    Dim sSQL As String
    Dim dsToWACode As Recordset

    sSQL = " SELECT cdval1, field1 " & _
           " FROM   " & T_LAB032 & _
           " WHERE  " & DBW("cdindex", LC3_WorkArea, 2)

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
    Call PrtBodyMod
    Call PrtBodyRmk
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        frmWACodeHelp.Visible = False
    End If
End Sub

Private Sub Form_Load()
    dtpStartDate1.Value = Format(Now, "yyyy-mm-dd")
    sStartTime = "00:00:00"
    sEndTime = Format(Now, "hh:mm:ss")

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
        sSQL = " SELECT cdval1 , field1 FROM " & T_LAB032 & _
               " Where  " & DBW("cdindex", LC3_WorkArea, 2) & _
               " and    " & DBW("cdval1", UCase(txtWACode.Text), 2)
        
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

Public Sub PrtHeader(ByVal Title As String)
    
    '    "결과수정자 명단"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(6.8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(6.8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm * 1.5)
    
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 2
    iposPtName = iposSEQ + 2 * iCm
    iposPtID = iposPtName + iCm + iCm / 2
    iposSAge = iposPtID + iCm + iCm
    iposIO = iposSAge + iCm
    
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "  환자번호", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    Call WriteStr(iCurY, iposIO, "  I/O", iCurY, iCm / 2)
    
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
    
    oldFontSize = 9 'Printer.FontSize
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

Public Sub PrtBodyMod()

    Dim sSQL1 As String
    Dim rsModList As Recordset
        
    Dim i%
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim prtTitle As String
    Dim sICSString As String
    
    sStartDate = Format(DateAdd("D", -3, dtpStartDate1.Value), CS_DateDbFormat)
    sEndDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
    sStartTime = Format(sStartTime, CS_TimeDbFormat)
    sEndTime = Format(sEndTime, CS_TimeDbFormat)

   
    sSQL1 = "Select distinct a.ptid as Ptid, b." & F_PTNM & " as PtNm, a.deptcd as Deptcd,  " & _
            "     a.rcvdt, a.rcvtm, a.ageday/365 + 1 as Age, a.sex as Sex " & _
                  "from " & T_LAB201 & " a, " & T_HIS001 & " b " & _
                  "where exists ( Select * from " & T_LAB308 & " b " & _
                  "                where b.workarea = '" & txtWACode.Text & "' " & _
                  "                  and b.accdt <= '" & sEndDate & "' " & _
                  "                  and b.accdt >= '" & sStartDate & "' " & _
                  "                  and b.accseq >= 1 " & _
                  "                  and b.mfydt = '" & sEndDate & "' " & _
                  "                  and b.workarea = a.workarea and b.accdt = a.accdt and b.accseq = a.accseq) " & _
                  "and  a.buildcd = '" & ObjSysInfo.BuildingCd & "' " & _
                  "and  a.ptid = b." & F_PTID & " " & _
                    "order by a.rcvdt, a.rcvtm"
    
    Set rsModList = New Recordset
    rsModList.Open sSQL1, DBConn
   
    If rsModList.EOF = True Then ' record가 존재하지 않을경우
       MsgBox " 결과수정자가 존재하지 않습니다. "
       Exit Sub
    End If
    
    Call InitReport
        '/* 보고서 제목
    prtTitle = "결과수정자 명단 "
    Call PrtHeader(prtTitle)
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark
    
    Dim iSeqNum As Integer
    
    With rsModList
        .MoveFirst
        For i = 1 To .RecordCount
                                    
            Call ChangeLine(iCm / 2)
            If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader(prtTitle)
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                    Call Print_WaterMark
            End If
            sICSString = ICSPatientString(.Fields("ptid").Value & "", enICSNum.LIS_ALL)
            
            iSeqNum = iSeqNum + 1
            
            Call WriteStr(iCurY, iposSEQ + iCm / 6, CStr(iSeqNum), iCurY, 0)
            If sICSString <> "" Then
                Call WriteStr(iCurY, iposSEQ + iCm, "Infection :" & sICSString, iCurY, 0)
                Call ChangeLine(iCm / 2)
            End If
            Call WriteStr(iCurY, iposPtName + iCm / 6, Trim(.Fields("Ptnm").Value), iCurY, 0)

            Call WriteStr(iCurY, iposPtID + iCm / 6, CStr(.Fields("PtId").Value), iCurY, 0)
            Call WriteStr(iCurY, iposSAge + iCm / 6, Trim(.Fields("Sex").Value) & "/" & Format(Trim(CStr(.Fields("Age").Value)), "00"), iCurY, 0)
            Call WriteStr(iCurY, iposIO + iCm / 6, Trim(.Fields("DeptCd").Value), iCurY, 0)
                            
            iCurY = iCurY + 0.8 * iCm
            Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "dot", 1, 0)
            
            .MoveNext
        Next i
        
        
    End With
        
    Call Print_WaterMark
    Printer.EndDoc
    
    Set rsModList = Nothing
    
End Sub
Public Sub PrtBodyRmk()

    Dim sSQL1 As String
    Dim rsRmkList As Recordset
        
    Dim i%
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim prtTitle As String
    Dim sICSString  As String
    
    sStartDate = Format(DateAdd("D", -3, dtpStartDate1.Value), CS_DateDbFormat)
    sEndDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
    sStartTime = Format(sStartTime, CS_TimeDbFormat)
    sEndTime = Format(sEndTime, CS_TimeDbFormat)
    
    sSQL1 = "Select distinct a.ptid as Ptid, b." & F_PTNM & " as PtNm, a.deptcd as Deptcd,  a.ageday/365 + 1 as Age, a.sex as Sex, a.rcvdt, a.rcvtm " & _
                  "from " & T_LAB201 & " a, " & T_HIS001 & " b " & _
                  "where a.workarea = '" & txtWACode.Text & "' " & _
                  "and  a.accdt <= '" & sEndDate & "' " & _
                  "and  a.accdt >= '" & sStartDate & "' " & _
                  "and  a.vfydt = '" & sEndDate & "' " & _
                  "and  a.rmkcd <> '' " & _
                  "and  a.buildcd = '" & ObjSysInfo.BuildingCd & "' " & _
                  "and  a.ptid = b." & F_PTID
    sSQL1 = sSQL1 & " UNION ALL " & _
                  "Select distinct a.ptid as Ptid, b." & F_PTNM & " as PtNm, a.deptcd as Deptcd,  a.ageday/365 + 1 as Age, a.sex as Sex, a.rcvdt, a.rcvtm  " & _
                  "from " & T_LAB201 & " a, " & T_HIS001 & " b " & _
                  "where a.workarea = '" & txtWACode.Text & "' " & _
                  "and  a.accdt <= '" & sEndDate & "' " & _
                  "and  a.accdt >= '" & sStartDate & "' " & _
                  "and  a.vfydt = '" & sEndDate & "' " & _
                  "and  a.footnotefg = '1' " & _
                  "and  a.buildcd = '" & ObjSysInfo.BuildingCd & "' " & _
                  "and  a.ptid = b." & F_PTID
                  
    sSQL1 = sSQL1 & "    order by rcvdt, rcvtm"
    
    Set rsRmkList = New Recordset
    rsRmkList.Open sSQL1, DBConn
   
    If rsRmkList.EOF = True Then ' record가 존재하지 않을경우
       MsgBox " 비고입력자가 존재하지 않습니다. "
       Exit Sub
    End If
    
    Call InitReport
        '/* 보고서 제목
    prtTitle = "비고입력자 명단 "
    Call PrtHeader(prtTitle)
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark
    
'    Dim temp1 As String, temp2 As String
    Dim iSeqNum As Integer
    
    With rsRmkList
        .MoveFirst
        For i = 1 To .RecordCount
            sICSString = ICSPatientString(.Fields("ptid").Value & "", enICSNum.LIS_ALL)
            Call ChangeLine(iCm / 2)
            If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader(prtTitle)
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                    Call Print_WaterMark
            End If
                
            iSeqNum = iSeqNum + 1
            If sICSString <> "" Then
                Call WriteStr(iCurY, iposSEQ + iCm, "Infection :" & sICSString, iCurY, 0)
                Call ChangeLine(iCm / 2)
            End If
            
            Call WriteStr(iCurY, iposSEQ + iCm / 6, CStr(iSeqNum), iCurY, 0)
            Call WriteStr(iCurY, iposPtName + iCm / 6, Trim(.Fields("Ptnm").Value), iCurY, 0)

            Call WriteStr(iCurY, iposPtID + iCm / 6, Trim(CStr(.Fields("PtId").Value)), iCurY, 0)
            Call WriteStr(iCurY, iposSAge + iCm / 6, Trim(.Fields("Sex").Value) & "/" & Format(Trim(CStr(.Fields("Age").Value)), "00"), iCurY, 0)
            Call WriteStr(iCurY, iposIO + iCm / 6, Trim(.Fields("DeptCd").Value), iCurY, 0)
                
             iCurY = iCurY + 0.8 * iCm
             Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "dot", 1, 0)
            
            .MoveNext
        Next i
        
        
    End With
        
    Call Print_WaterMark
    Printer.EndDoc
    
    Set rsRmkList = Nothing
    
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
    Printer.Print "RUN-DATE: " & sDate
        
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
    Dim BuildingNm As String
    
'    objSysInfo.BuildingNm = GetSetting(AppName:=RegHdBld, Section:=RegSsBld, Key:=RegK2Bld, Default:="(건물정보 누락)")
    sStartDate = Format(DateAdd("D", -3, dtpStartDate1.Value), CS_DateDbFormat)
    sEndDate = Format(dtpStartDate1.Value, CS_DateDbFormat)
     
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "WorkArea  : " & lblWAName.Caption & "   ( " & ObjSysInfo.BuildingNm & " )"
    
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("Work Area : ") + iCm / 6
    Printer.Print "보고 일자  : " & sEndDate
                                    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
End Sub
