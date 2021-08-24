VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm413PWorkListM 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FCEFE9&
      Caption         =   "출 력 (&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   12
      Top             =   3900
      Width           =   1320
   End
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00EFFEFE&
      Height          =   2400
      Left            =   5070
      TabIndex        =   6
      Top             =   4590
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.ListBox lstSpeGroup 
      BackColor       =   &H00F7FFF7&
      Height          =   2400
      Left            =   1380
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2850
      Left            =   75
      TabIndex        =   3
      Top             =   1005
      Width           =   10770
      Begin VB.CommandButton cmdWsUnitHelp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         Height          =   345
         Left            =   3900
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   1230
         Width           =   315
      End
      Begin VB.CommandButton cmdWSCodeHelp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         Height          =   360
         Left            =   3900
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtWSUnit 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   345
         HideSelection   =   0   'False
         Left            =   2220
         TabIndex        =   1
         Top             =   1245
         Width           =   1665
      End
      Begin VB.TextBox txtSpeGroupCD 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   360
         HideSelection   =   0   'False
         Left            =   2220
         TabIndex        =   0
         Top             =   735
         Width           =   1665
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   255
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   735
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
         Caption         =   "검체군코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   255
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1245
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
         Caption         =   "WorkSheet Unit"
         Appearance      =   0
      End
      Begin VB.Label lblSpeGroupNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   4245
         TabIndex        =   4
         Top             =   735
         Width           =   5745
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "미생물검사 Work List 출력"
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
      Left            =   1050
      TabIndex        =   9
      Top             =   435
      Width           =   4065
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
      Top             =   345
      Width           =   6090
   End
End
Attribute VB_Name = "frm413PWorkListM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPageWidth As Integer
Dim iPageHeight As Integer
Dim iCurY As Integer
Const iCm = 567
Const iLineHeight = 10

Dim iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%, iposAnti%


Public Event FormClose()


Private Sub cmdExit_Click()
   Unload Me
'   Set frm413PWorkListM = Nothing

    RaiseEvent FormClose
End Sub

Private Sub cmdReport_Click()
    ' 클래스를 이용하여 출력
    Dim MyReport As New clsWorkListM
     
    If txtSpeGroupCD.Text <> "" And txtWSUnit.Text <> "" And _
        lblSpeGroupNm.Caption <> "" Then
        MyReport.Worksheet2 = False
        Call MyReport.GetInputData(txtSpeGroupCD.Text, txtWSUnit, lblSpeGroupNm.Caption)
        Call MyReport.PrintReport
        Set MyReport = Nothing
    End If
    
    ' 클래스를 이용하지 않고 출력
'    If txtSpeGroupCD.Text <> "" And _
'        txtWSUnit.Text <> "" Then
'        Call InitReport
'        Call PrtHeader
'        Call prtPageNum
'        Call prtTerm
'        Call PrtBody
'        Printer.EndDoc
'    End If
End Sub

Private Sub PrtBody()
    Dim LabNo As tLabno
    Dim sSqlGetLabno As String
    Dim rsLabno As Recordset
    Dim i%
    
    sSqlGetLabno = " select workarea, accdt, accseq" & _
                   " from   " & T_LAB402 & _
                   " where  " & DBW("wscd", Trim(txtSpeGroupCD.Text), 2) & _
                   " and    " & DBW("wsunit", Trim(txtWSUnit.Text), 2)
    Set rsLabno = New Recordset
    rsLabno.Open sSqlGetLabno, DBConn
    
    For i = 1 To rsLabno.RecordCount
    
        LabNo.sWorkarea = Trim(rsLabno.Fields("workarea").Value)
        LabNo.sAccDt = Trim(rsLabno.Fields("accdt").Value)
        LabNo.iAccSeq = Trim(rsLabno.Fields("accseq").Value)
        
        If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
            Printer.NewPage
            Call PrtHeader
            Call prtPageNum
            Call prtTerm
        End If

        Call DspSeq(i)
        Call DspWorkno(LabNo.sAccDt, LabNo.iAccSeq)
        Call DspLab201Data(LabNo)
        Call DspPtnm(LabNo)
        Call DspTestCD(LabNo)
'        Call DspAntiBiotic
        Call ChangeLine(iCm * 4.5)
        
        rsLabno.MoveNext
    Next i
    
    Set rsLabno = Nothing
End Sub

Private Sub DspAntiBiotic()
    Dim sSqlGetAnti As String
    Dim rsGetAnti As Recordset
    Dim ioldcury As Integer
    Dim i%
    Dim ioldcurx
    sSqlGetAnti = " select field1 as AntiBoitic" & _
                  " from   " & T_LAB032 & _
                  " where  " & DBW("cdindex", LC3_AntiBiotic, 2) & _
                  " and    ( field5 = null or field5 = '')"
    Set rsGetAnti = New Recordset
    rsGetAnti.Open sSqlGetAnti, DBConn
    
    If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
        Printer.NewPage
        Call PrtHeader
        Call prtPageNum
        Call prtTerm
    End If
    
    ioldcury = iCurY
    ioldcurx = iposAnti
    
    For i = 1 To rsGetAnti.RecordCount
        Call WriteStr(iCurY, iposAnti, rsGetAnti.Fields("AntiBoitic").Value, iCurY, 0)
        
        iposAnti = iposAnti + iCm + iCm / 3
        
        If i Mod 3 = 0 Then
            Call ChangeLine(iCm / 2)
            iposAnti = ioldcurx
        End If
        
        rsGetAnti.MoveNext
    Next i
    
    If i > 9 Then
        iCurY = ioldcury + (iCm / 2) * 3
    End If
    
    iposAnti = ioldcurx
    Set rsGetAnti = Nothing
    
End Sub

Private Sub DspTestCD(LabNo As tLabno)
    Dim sSqlGetTestCD As String
    Dim rsGetTestCD As Recordset
    
    Dim sSqlGetTestNm As String
    Dim rsGetTestNm As Recordset
    Dim i%, oldipostestcd%
    
        
    sSqlGetTestCD = " select testcd " & _
                    " from    " & T_LAB404 & _
                    " where ( " & DBW("workarea", LabNo.sWorkarea, 2) & _
                    " and     " & DBW("accdt", LabNo.sAccDt, 2) & _
                    " and     " & DBW("accseq", LabNo.iAccSeq, 2) & " ) " & _
                    " and     " & _
                    "       ( detailfg = null or detailfg = '' ) " & _
                    " union   " & _
                    " select testcd " & _
                    " from    " & T_LAB404 & _
                    " where ( " & DBW("workarea", LabNo.sWorkarea, 2) & _
                    " and     " & DBW("accdt", LabNo.sAccDt, 2) & _
                    " and     " & DBW("accseq", LabNo.iAccSeq, 2) & " ) " & _
                    " and     " & _
                    "       (detailfg is not null ) and rstdiv = '*' "
                    
    Set rsGetTestCD = New Recordset
    rsGetTestCD.Open sSqlGetTestCD, DBConn
    
    If rsGetTestCD.EOF = True Then Exit Sub
    
    oldipostestcd = iposTestCD
    For i = 1 To rsGetTestCD.RecordCount
        sSqlGetTestNm = " select abbrnm5 " & _
                        " from  " & T_LAB001 & _
                        " where " & DBW("testcd", rsGetTestCD.Fields("testcd").Value, 2)
        Set rsGetTestNm = Nothing
        Set rsGetTestNm = New Recordset
        rsGetTestNm.Open sSqlGetTestNm, DBConn
        
        Call WriteStr(iCurY, iposTestCD, rsGetTestNm.Fields("abbrnm5").Value, iCurY, 0)
        iposTestCD = iposTestCD + 1.3 * iCm
        rsGetTestCD.MoveNext
    Next i
    
    iposTestCD = oldipostestcd
    
    Set rsGetTestCD = Nothing
                        
    Set rsGetTestNm = Nothing
End Sub
Private Sub DspPtnm(LabNo As tLabno)
    Dim sSqlGetPtnm As String
    Dim rsGetPtnm As Recordset
    
    
    sSqlGetPtnm = " select " & F_PTNM & " as ptnm " & _
                  " from   " & T_HIS001 & _
                  " where  " & F_PTID & " = " & _
                  "        (select ptid " & _
                  "         from  " & T_LAB201 & _
                  "         where " & DBW("workarea", LabNo.sWorkarea, 2) & " and " & _
                                      DBW("accdt", LabNo.sAccDt, 2) & " and " & _
                                      DBW("accseq", LabNo.iAccSeq, 2) & ")"
    Set rsGetPtnm = New Recordset
    rsGetPtnm.Open sSqlGetPtnm, DBConn
    
    If rsGetPtnm.EOF = True Then Exit Sub
    
    Call WriteStr(iCurY, iposPtName, rsGetPtnm.Fields("ptnm").Value, iCurY, 0)
    
    Set rsGetPtnm = Nothing
        
End Sub
Private Sub DspSeq(SeqNum As Integer)
    Call WriteStr(iCurY, iposSEQ, CStr(SeqNum), iCurY, 0)
End Sub
Private Sub DspWorkno(sAccDt As String, iAccSeq As Long)
    Dim sWorkno As String
    sWorkno = sAccDt & "-" & CStr(iAccSeq)
    Call WriteStr(iCurY, iposWorkNo, sWorkno, iCurY, 0)
End Sub
Private Sub DspLab201Data(LabNo As tLabno)
    Dim sSqlGetlab201 As String
    Dim rslab201Data As Recordset
    Dim S_Age As String
    sSqlGetlab201 = " select ptid, sex, ageday, deptcd, storecd, " & _
                    "        rcvdt, rcvtm , spccd" & _
                    " from   " & T_LAB201 & _
                    " where  " & DBW("workarea", LabNo.sWorkarea, 2) & " and " & _
                                 DBW("accdt", LabNo.sAccDt, 2) & " and " & _
                                 DBW("accseq", LabNo.iAccSeq, 2)
    Set rslab201Data = New Recordset
    rslab201Data.Open sSqlGetlab201, DBConn

    Call WriteStr(iCurY, iposPtID, rslab201Data.Fields("ptid").Value, iCurY, 0)
    S_Age = rslab201Data.Fields("sex").Value & "/" & _
            (rslab201Data.Fields("ageday").Value \ 365 + 1)
    Call WriteStr(iCurY, iposSAge, S_Age, iCurY, 0)
    Call WriteStr(iCurY, iposIO, rslab201Data.Fields("deptcd").Value, iCurY, 0)
    Call WriteStr(iCurY, iposSF, rslab201Data.Fields("storecd").Value, iCurY, 0)
    Call WriteStr(iCurY, iposRcv, DelFirst2Chr(rslab201Data.Fields("rcvdt").Value), iCurY, 0)
    Call WriteStr(iCurY, iposRcv + iCm + iCm / 6, CvtTmFormat(rslab201Data.Fields("rcvtm").Value), iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, rslab201Data.Fields("spccd").Value, iCurY, 0)
    
    Set rslab201Data = Nothing
End Sub
Private Function DelFirst2Chr(sStr As String) As String
    DelFirst2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function
Private Function DelLast2Chr(sStr As String) As String
    DelLast2Chr = Trim(Mid(sStr, 1, Len(sStr) - 2))
End Function
Private Function CvtTmFormat(sStr As String) As String
    Dim Time As String
    Dim Hour As String
    Dim Min As String
    
    Time = DelLast2Chr(sStr)
    Hour = DelLast2Chr(Time)
    Min = DelFirst2Chr(Time)
    
    CvtTmFormat = Hour & ":" & Min
    
End Function

Public Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
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

Public Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str
End Sub

Public Sub ChangeLine(iLineSpace As Integer)
    iCurY = iCurY + iLineSpace
    Printer.CurrentY = iCurY
    
End Sub

Public Sub PrtHeader()
   
    Dim Title As String
    Dim sWSNanme As String
    Dim iHeaderPosY As Integer
    
    '/* 보고서 제목
    Title = "미생물업무나열서"
    
    
    '        "미생물업무나열서"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    
    iposSEQ = iCm / 2                       '284
    iposWorkNo = iposSEQ + iCm / 1.2        '757
    iposPtName = iposWorkNo + 1.7 * iCm       '1891
    iposPtID = iposPtName + iCm + iCm / 1.6 '2894
    iposSAge = iposPtID + iCm + iCm / 6     '3744
    iposIO = iposSAge + iCm                 '4311
    iposSF = iposIO + 0.9 * iCm             '4765
    iposTestCD = iposSF + 0.8 * iCm
    iposRcv = iposTestCD + 4 * iCm
    iposSpccd = iposRcv + 2.2 * iCm
    'iposAnti = iposSpccd + 1 * iCm
    
    iCurY = iCurY + iCm * 1.5
    
        
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
  
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposWorkNo, "Work No", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "환자ID", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    Call WriteStr(iCurY, iposIO, "I/O", iCurY, 0)
    Call WriteStr(iCurY, iposSF, "S/F", iCurY, 0)
    Call WriteStr(iCurY, iposTestCD, "검사항목", iCurY, 0)
    Call WriteStr(iCurY, iposRcv, "검체도착시간", iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "검체", iCurY, iCm / 2)
'    Call WriteStr(iCurY, iposAnti, "항생제", iCurY, iCm / 2)
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
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
    Dim sSqlGetMedia As String
    Dim sSqlGetWorkDt As String
    Dim rsMediaCD As Recordset
    Dim rsWorkDt As Recordset
    Dim iXpos As Integer
    Dim i%
    
    sSqlGetMedia = " select cdval2 as MediaCD" & _
                    " from   " & T_LAB031 & _
                    " where  " & DBW("cdindex", LC2_SpcMedia, 2) & " and " & _
                                 DBW("cdval1", Trim(txtSpeGroupCD.Text), 2)
    Set rsMediaCD = New Recordset
    rsMediaCD.Open sSqlGetMedia, DBConn
    
    If rsMediaCD.EOF = True Then Exit Sub
    
    sSqlGetWorkDt = " select workdt , worktm" & _
                    " from   " & T_LAB401 & _
                    " where  " & DBW("wscd", Trim(txtSpeGroupCD.Text), 2) & " and " & _
                                 DBW("wsunit", Trim(txtWSUnit.Text), 2)
    Set rsWorkDt = New Recordset
    rsWorkDt.Open sSqlGetWorkDt, DBConn
    
    If rsWorkDt.EOF = True Then Exit Sub
      
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "작업일     : " & Trim(rsWorkDt.Fields("workdt").Value) & _
                  vbTab & CvtTmFormat(Trim(rsWorkDt.Fields("worktm").Value))
    
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ")
    Printer.Print "검체군명   : " & Trim(lblSpeGroupNm.Caption)
    
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ") + _
                       Printer.TextHeight("workArea : ")
    Printer.Print "배지코드   :"
    
    For i = 1 To rsMediaCD.RecordCount
        
        Printer.CurrentX = 3 * iCm + iXpos
        Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ") + _
                           Printer.TextHeight("workArea : ")
        Printer.Print Trim(rsMediaCD.Fields("MediaCD").Value) & "   "
        iXpos = Len(rsMediaCD.Fields("MediaCD").Value) * 0.3 * iCm
        rsMediaCD.MoveNext
    Next i
    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
    Set rsMediaCD = Nothing
    Set rsWorkDt = Nothing
End Sub

Private Sub cmdReport_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdExit.SetFocus
    End If
End Sub


Private Sub cmdWSCodeHelp_Click()

    Call txtSpeGroupCD_KeyPress(vbKeyDown)

End Sub

Private Sub cmdWsUnitHelp_Click()

    Call txtWSUnit_KeyPress(vbKeyDown)

End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        lstSpeGroup.Visible = False
        lstWSUnit.Visible = False
    End If

End Sub

Private Sub Form_Load()
    
    Call LoadLstSpeGroup
  
End Sub

Private Sub lstSpeGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then Call lstSpeGroup_KeyDown(vbKeyReturn, 0)
        
End Sub

Private Sub lstWSUnit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then Call lstWSUnit_keydown(vbKeyReturn, 0)

End Sub

Private Sub txtSpeGroupCD_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        Call txtSpeGroupCD_KeyPress(vbKeyDown)
    End If
End Sub
Private Function ChkSpeGroupExist() As Boolean
    Dim i%
    
    For i = 0 To lstSpeGroup.ListCount - 1
        If Trim(txtSpeGroupCD.Text) = Trim(Mid(lstSpeGroup.List(i), 1, _
                                      InStr(1, lstSpeGroup.List(i), vbTab) - 1)) Then
            ChkSpeGroupExist = True          ' Exist
            Exit Function
        End If
    Next i
    ChkSpeGroupExist = False                 ' Not Exist
End Function
Private Sub txtSpeGroupCD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    
    txtWSUnit.Text = ""       ' txtWsunit Clear

    
    If lstSpeGroup.ListCount = 0 Then Exit Sub
    Call DsplstSpeGroup
    
    If KeyAscii = vbKeyReturn And Len(txtSpeGroupCD.Text) > 1 And _
       lstSpeGroup.Text <> "" Then
       
        txtSpeGroupCD.Text = Trim(Mid(lstSpeGroup.Text, 1, _
                 InStr(1, lstSpeGroup.Text, vbTab) - 1))
        
        lblSpeGroupNm.Caption = Trim(Mid(lstSpeGroup.Text, _
                                InStr(1, lstSpeGroup.Text, vbTab) + 1, _
                                Len(lstSpeGroup.Text)))

        lstSpeGroup.Visible = False
        Call LoadLstWSUnit
    End If
    Call medCodeHelp(KeyAscii, lstSpeGroup, txtSpeGroupCD.Text, txtSpeGroupCD, txtWSUnit)

End Sub

Public Sub DsplstSpeGroup()
    lstSpeGroup.Top = Frame1.Top + txtSpeGroupCD.Top + txtSpeGroupCD.Height
    lstSpeGroup.Left = txtSpeGroupCD.Left
    lstSpeGroup.Visible = True
    lstSpeGroup.ZOrder 0
End Sub

Public Sub LoadLstSpeGroup()
    Dim rsSpeGroup As Recordset
    Dim sSqlGetSpeGroup As String
    Dim i%
    
    sSqlGetSpeGroup = " select cdval1, field1, field3" & _
                      " from   " & T_LAB032 & _
                      " where  " & DBW("cdindex", LC3_SGroup, 2) & _
                      " order by field3 "

    Set rsSpeGroup = New Recordset      ' Get 검체군코드및 명
    rsSpeGroup.Open sSqlGetSpeGroup, DBConn
    
    If rsSpeGroup.EOF = True Then
        MsgBox " 검체군 code가 존재하지 않습니다."
        Exit Sub
    End If

    With lstSpeGroup
        rsSpeGroup.MoveFirst
        For i = 0 To rsSpeGroup.RecordCount - 1
            .AddItem rsSpeGroup.Fields("cdval1").Value & vbTab & _
                     rsSpeGroup.Fields("field1").Value, i
            rsSpeGroup.MoveNext
        Next i
    End With
    Set rsSpeGroup = Nothing
End Sub

Private Sub lstSpeGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then    ' Enter key 입력시
        txtSpeGroupCD.Text = Trim(Mid(lstSpeGroup.Text, 1, _
                 InStr(1, lstSpeGroup.Text, vbTab) - 1))
        lblSpeGroupNm.Caption = Trim(Mid(lstSpeGroup.Text, _
                                     InStr(1, lstSpeGroup.Text, vbTab) + 1, _
                                     Len(lstSpeGroup.Text)))
        lstSpeGroup.Visible = False
        
        Call LoadLstWSUnit
        txtWSUnit.SetFocus
    End If
End Sub
Private Sub lstWSUnit_keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtWSUnit.Text = Trim(Mid(lstWSUnit.Text, 1, _
                 InStr(1, lstWSUnit.Text, vbTab) - 1))
        lstWSUnit.Visible = False
        cmdReport.SetFocus
    End If
End Sub
Private Sub LoadLstWSUnit()
    Dim rsWSUnit As Recordset
    Dim sSqlGetWSUnit As String
    Dim i%
    
    lstWSUnit.Clear
    
    sSqlGetWSUnit = " Select /*+ RULE **/ a.wsunit, a.workdt, a.worktm from " & T_LAB401 & " a " & _
                    " where  " & DBW("a.wscd", Trim(txtSpeGroupCD.Text), 2) & _
                    " and    exists (select * from " & T_LAB404 & " b, " & T_LAB402 & " c " & _
                    "        where c.wscd = a.wscd and c.wsunit = a.wsunit" & _
                    "        and b.workarea = c.workarea and b.accdt = c.accdt  " & _
                    "        and b.accseq = c.accseq and b.rsttype = c.scfg and " & DBW("c.genfg<", MWS_Final) & _
                    "        and " & DBW("b.stscd<", enStsCd.StsCd_LIS_FinRst) & ") " & _
                    " order by a.wsunit desc"
     
    'sSqlGetWSUnit = " select wsunit , workdt, worktm" & _
                    " from " & T_LAB401 & _
                    " where wscd = '" & Trim(txtSpeGroupCD.Text) & "'"
                    

    Set rsWSUnit = New Recordset     ' Get WSUnit code
    rsWSUnit.Open sSqlGetWSUnit, DBConn
    
    If rsWSUnit.EOF = True Then         ' wsunit code가 존재하지 않을 경우
        Exit Sub
    End If

    With lstWSUnit
        
        rsWSUnit.MoveFirst
        For i = 0 To rsWSUnit.RecordCount - 1
            .AddItem rsWSUnit.Fields("wsunit").Value & vbTab & _
                     rsWSUnit.Fields("workdt").Value & vbTab & _
                     CvtTmFormat(rsWSUnit.Fields("worktm").Value), i
        
            rsWSUnit.MoveNext
        Next i
    End With
    Set rsWSUnit = Nothing
End Sub

Private Sub txtWSUnit_keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call txtWSUnit_KeyPress(vbKeyDown)
    End If
End Sub

Private Sub txtWSUnit_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))

    If lstWSUnit.ListCount = 0 Then Exit Sub
    Call DsplstWSUnit
    
    If KeyAscii = vbKeyReturn Then
        txtWSUnit.Text = Trim(Mid(lstWSUnit.Text, 1, _
                 InStr(1, lstWSUnit.Text, vbTab) - 1))
               
        lstWSUnit.Visible = False
        cmdReport.SetFocus
    End If
    Call medCodeHelp(KeyAscii, lstWSUnit, txtWSUnit.Text, txtWSUnit, cmdReport)

End Sub

Public Sub DsplstWSUnit()
    lstWSUnit.Top = Frame1.Top + txtWSUnit.Top + txtWSUnit.Height
    lstWSUnit.Left = txtWSUnit.Left
    lstWSUnit.Visible = True
    lstWSUnit.ZOrder 0
End Sub
