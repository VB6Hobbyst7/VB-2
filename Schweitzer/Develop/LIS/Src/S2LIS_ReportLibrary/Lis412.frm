VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm412PWorkListG 
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
      TabIndex        =   15
      Top             =   3900
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "128"
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2850
      Left            =   90
      TabIndex        =   5
      Top             =   1005
      Width           =   10755
      Begin VB.CommandButton cmdWSCodeHelp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         Height          =   360
         Left            =   3795
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox txtWSCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   360
         HideSelection   =   0   'False
         Left            =   2205
         TabIndex        =   0
         Top             =   705
         Width           =   1590
      End
      Begin VB.TextBox txtStartWNum 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   315
         HideSelection   =   0   'False
         Left            =   2190
         TabIndex        =   2
         Top             =   2025
         Width           =   765
      End
      Begin VB.TextBox txtEndWNum 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00F1F5F4&
         Height          =   315
         HideSelection   =   0   'False
         Left            =   3570
         TabIndex        =   3
         Top             =   2025
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpWorkDate 
         Height          =   315
         Left            =   2190
         TabIndex        =   1
         Top             =   1545
         Width           =   2190
         _ExtentX        =   3863
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
         Format          =   73531395
         CurrentDate     =   36370
      End
      Begin MedControls1.LisLabel lblWSName 
         Height          =   345
         Left            =   4125
         TabIndex        =   9
         Top             =   705
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   609
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
         RightGab        =   0
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
         Caption         =   "WorkSheet Code"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   1260
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1530
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "작업일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   1
         Left            =   1260
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2010
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "작업순번"
         Appearance      =   0
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DBE6E6&
         Caption         =   "까지"
         Height          =   255
         Left            =   4455
         TabIndex        =   7
         Top             =   2100
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00DBE6E6&
         Caption         =   "부터"
         Height          =   255
         Left            =   3030
         TabIndex        =   6
         Top             =   2100
         Width           =   375
      End
   End
   Begin VB.ListBox lstWSCode 
      BackColor       =   &H00EBEBEB&
      Height          =   2400
      Left            =   1470
      TabIndex        =   4
      Top             =   4890
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "일반검사 Work List 출력"
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
      Left            =   885
      TabIndex        =   11
      Top             =   525
      Width           =   3735
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H0088A2B7&
      FillColor       =   &H00DEEEFE&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   165
      Shape           =   4  '둥근 사각형
      Top             =   435
      Width           =   5355
   End
End
Attribute VB_Name = "frm412PWorkListG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPageWidth As Integer
Dim iPageHeight As Integer
Dim iCurY As Integer


Dim iWidthTestCD As Integer
Dim iTestCount As Integer

    
Const iCm = 567
Const iLineHeight = 10

Dim iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposXTestCD%, iposSpccd%
Dim aWSTestCD() As String


Public Event FormClose()

Private Sub cmdExit_Click()
   Unload Me
'   Set frm412PWorkListG = Nothing

    RaiseEvent FormClose
End Sub

Public Sub Call_cmdReport_Click()
   Call cmdReport_Click
End Sub

Private Sub cmdReport_Click()
    
    Dim MyReport As New clsWorkListG
    If txtWSCode.Text <> "" And lblWSName.Caption <> "" And _
        txtStartWNum.Text <> "" And txtEndWNum.Text <> "" Then
        
        With MyReport
            .WorkCode = txtWSCode.Text
            .WorkName = lblWSName.Caption
            .WorkDate = Format(dtpWorkDate.Value, CS_DateDbFormat)
            .FromSeq = txtStartWNum.Text
            .ToSeq = txtEndWNum.Text
            Call .Print_Worksheet
        End With
        Set MyReport = Nothing
    End If
    
'    Exit Sub
'
'    'If Val(Mid$(txtWSCode.Text, Len(txtWSCode.Text) - 2, 3)) > 900 Then
'    If iTestCount > 10 Then
'        Call Reporting900
'    Else
'        Call Reporting
'    End If
'
End Sub

Public Sub Reporting900()

    Dim i%, j%, K%, iCurXTestCD%, iCurYTestCD%

    Dim sSqlLab301 As String
    Dim sSqlLab201 As String

    Dim ssqlHis001 As String
    
    Dim sWorkDate As String
    
    Dim rslab301 As Recordset
    Dim rslab201 As Recordset
    Dim rshis001 As Recordset

    Dim sWorkno As String
    Dim S_Age As String
    
    
    Dim sWorkarea As String
    Dim sAccDt As String
    Dim iAccSeq As Integer


    sWorkDate = Format(dtpWorkDate.Value, CS_DateDbFormat)
    
    '******* get workseq, labnum(workarea, accdt, accseq) ************************
    sSqlLab301 = " select workseq,workarea,accdt, accseq " & _
                 " from  " & T_LAB301 & _
                 " where " & DBW("workdt", sWorkDate, 2) & _
                 " and   " & DBW("workcd", Trim(txtWSCode.Text), 2) & _
                 " and   " & DBW("workseq >= ", Trim(txtStartWNum.Text)) & _
                 " and   " & DBW("workseq <= ", Trim(txtEndWNum.Text)) & _
                 " order by workseq "
    Set rslab301 = New Recordset
    rslab301.Open sSqlLab301
    
    If rslab301.EOF = True Then
        MsgBox " Worksheet 내역이 존재하지 않습니다.. "
        Set rslab301 = Nothing
        Exit Sub
    End If
    
    Call InitReport
    Call PrtHeader900
    Call prtPageNum
    Call prtTerm

    rslab301.MoveFirst
    
    For i = 1 To rslab301.RecordCount ' 301의 각 labnum 마다
        
        If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
            Printer.NewPage
            Call PrtHeader900
            Call prtPageNum
            Call prtTerm
        End If
      
        Call ChangeLine(iCm / 2)
        ' workseq, workno출력
        
        sWorkarea = rslab301.Fields("workarea").Value
        sAccDt = rslab301.Fields("accdt").Value
        iAccSeq = rslab301.Fields("accseq").Value
        
        
        '****** ptid, sex/age, deptcd, storecd, rcvdt, rcvtm, spccd 출력*******
        sSqlLab201 = " select ptid, sex, ageday,deptcd, storecd,rcvdt,statfg, " & _
                     "        rcvtm, spccd" & _
                     " from   " & T_LAB201 & _
                     " where  " & DBW("workarea", sWorkarea, 2) & _
                     " and    " & DBW("accdt", sAccDt, 2) & _
                     " and    " & DBW("accseq", iAccSeq, 2)
        Set rslab201 = New Recordset
        rslab201.Open sSqlLab201
        
        '응급여부 (**)로 표시..
        If Trim(rslab201.Fields("statfg").Value) = "1" Then
            Call WriteStr(iCurY, iposSEQ, rslab301.Fields("workseq").Value & "**", iCurY, 0)
        Else
            Call WriteStr(iCurY, iposSEQ, rslab301.Fields("workseq").Value, iCurY, 0)
        End If
        'Call WriteStr(iCurY, iposSEQ, rslab301.Fields("workseq").Value, iCurY, 0)
        sWorkno = DelFirst2Chr(rslab301.Fields("accdt").Value) & "-" & _
                 rslab301.Fields("accseq").Value
        Call WriteStr(iCurY, iposWorkNo, sWorkno, iCurY, 0)
        
        
        '****** ptnm,sex출력 ******************************************************
        ssqlHis001 = " select " & F_PTNM & " as ptnm, " & F_SEX & " as sex " & _
                     " from   " & T_HIS001 & _
                     " where  " & DBW(F_PTID, rslab201.Fields("ptid").Value, 2)
        Set rshis001 = Nothing
        Set rshis001 = New Recordset
        rshis001.Open ssqlHis001, dbconn
        
        Call WriteStr(iCurY, iposPtID, rslab201.Fields("ptid").Value, iCurY, 0)
        
        S_Age = rshis001.Fields("sex").Value & ""
        
        If IsNumeric(S_Age) Then
            S_Age = Choose((Val(S_Age) Mod 2) + 1, "F", "M")
        End If
        
        S_Age = S_Age & "/" & ("" & rslab201.Fields("ageday").Value \ 365 + 1)
        
        Call WriteStr(iCurY, iposSAge, S_Age, iCurY, 0)
        
        Call WriteStr(iCurY, iposIO, rslab201.Fields("deptcd").Value, iCurY, 0)
        Call WriteStr(iCurY, iposSF, rslab201.Fields("storecd").Value, iCurY, 0)
        Call WriteStr(iCurY, iposRcv, DelFirst2Chr(rslab201.Fields("rcvdt").Value), iCurY, 0)
        Call WriteStr(iCurY, iposRcv + iCm + 1.5, CvtTmFormat(rslab201.Fields("rcvtm").Value), iCurY, 0)
        Call WriteStr(iCurY, iposSpccd, rslab201.Fields("spccd").Value, iCurY, 0)
        
        Call WriteStr(iCurY, iposPtName, rshis001.Fields("ptnm").Value, iCurY, 0)
        
        '****** 해당 Test Code 괄호(    )  출력 *******************************
        
        Call DspBrackketTestCD(sWorkarea, sAccDt, iAccSeq)
        rslab301.MoveNext
    Next i
    Printer.EndDoc
    
    
    Set rslab301 = Nothing
    
    Set rslab201 = Nothing
    
    Set rshis001 = Nothing
End Sub

Private Sub cmdWSCodeHelp_Click()
    lstWSCode.ListIndex = medListFind(lstWSCode, txtWSCode.Text)
    DsplstWSCode
End Sub

Private Sub dtpWorkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtStartWNum.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        lstWSCode.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Call LoadLstWSCode
    dtpWorkDate.Value = Format(Now, "yyyy-mm-dd")
End Sub

Public Sub LoadLstWSCode()
    Dim rsWSCode As Recordset
    Dim sSqlGetWSCode As String
    Dim i%
    
    sSqlGetWSCode = " select  a.cdval1 as WorkCd, a.field1 as WorkNm, count(b.testcd) as TestCnt " & _
                    " from    " & T_LAB032 & " a, " & T_LAB008 & " b " & _
                    " where   " & DBW("a.cdindex", LC3_WorkSheetName, 2) & _
                    " and     b.workcd = a.cdval1 " & _
                    " and     " & DBW("a.field2", objsysinfo.BuildingCd, 2) & _
                    " group by a.cdval1, a.field1 "

    Set rsWSCode = New Recordset
    rsWSCode.Open sSqlGetWSCode, dbconn

    If rsWSCode.EOF = True Then
        MsgBox " worksheet code가 존재하지 않습니다."
        Exit Sub
    End If

    With lstWSCode
        rsWSCode.MoveFirst
        For i = 0 To rsWSCode.RecordCount - 1
            .AddItem rsWSCode.Fields("WorkCd").Value & vbTab & _
                     rsWSCode.Fields("WorkNm").Value & vbTab & _
                     rsWSCode.Fields("TestCnt").Value, i
            rsWSCode.MoveNext
        Next i
    End With
    Set rsWSCode = Nothing
End Sub



Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        txtWSCode.Text = Trim(Mid(lstWSCode.Text, 1, _
                 InStr(1, lstWSCode.Text, vbTab) - 1))
        lblWSName.Caption = medgetp(lstWSCode.Text, 2, vbTab)
        'lblWSName.Caption = Trim(Mid(lstWSCode.Text, _
                                     InStr(1, lstWSCode.Text, vbTab) + 1, _
                                     Len(lstWSCode.Text)))
        iTestCount = Val(medgetp(lstWSCode.Text, 3, vbTab))
        lstWSCode.Visible = False
        dtpWorkDate.SetFocus
    End If
End Sub

Private Sub lstWSCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then Call lstWSCode_KeyDown(vbKeyReturn, 0)

End Sub

Private Sub txtEndWNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdReport.SetFocus
    End If
End Sub


Private Sub txtStartWNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtEndWNum.SetFocus
    End If
End Sub

Private Sub txtWSCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call txtWSCode_KeyPress(vbKeyDown)
    End If
End Sub

Private Sub txtWSCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    
    If KeyAscii = vbKeyReturn And Len(txtWSCode.Text) > 1 And _
        lstWSCode.Text <> "" Then
         Call lstWSCode_KeyDown(vbKeyReturn, 0)
         dtpWorkDate.SetFocus
        Exit Sub
    End If
    
    Call DsplstWSCode
'    Call SearchWSCode(KeyAscii, lstWSCode, txtWSCode.Text)
    Call medCodeHelp(KeyAscii, lstWSCode, txtWSCode.Text, txtWSCode, dtpWorkDate)

End Sub

Public Sub DsplstWSCode()
    lstWSCode.Top = Frame1.Top + txtWSCode.Top + txtWSCode.Height
    lstWSCode.Left = txtWSCode.Left + Frame1.Left
    lstWSCode.Visible = True
    lstWSCode.ZOrder 0
End Sub

Public Sub Reporting()

    Dim i%, j%, K%, iCurXTestCD%, iCurYTestCD%

    Dim sSqlLab301 As String
    Dim sSqlLab201 As String

    Dim ssqlHis001 As String
    
    Dim sWorkDate As String
    
    Dim rslab301 As Recordset
    Dim rslab201 As Recordset
    Dim rshis001 As Recordset

    Dim sWorkno As String
    Dim S_Age As String
    
    
    Dim sWorkarea As String
    Dim sAccDt As String
    Dim iAccSeq As Integer
    Dim iHeaderPosY1 As Integer
    Dim iHeaderPosY2 As Integer


    sWorkDate = Format(dtpWorkDate.Value, CS_DateDbFormat)
    
    '******* get workseq, labnum(workarea, accdt, accseq) ************************
    sSqlLab301 = " select workseq,workarea,accdt, accseq " & _
                 " from   " & T_LAB301 & _
                 " where  " & DBW("workdt", sWorkDate, 2) & _
                 " and    " & DBW("workcd", Trim(txtWSCode.Text), 2) & _
                 " and    " & DBW("workseq >= ", Trim(txtStartWNum.Text)) & _
                 " and    " & DBW("workseq <= ", Trim(txtEndWNum.Text)) & _
                 " order by workseq "
    Set rslab301 = New Recordset
    rslab301.Open sSqlLab301
    
    If rslab301.EOF = True Then
        MsgBox " Worksheet 내역이 존재하지 않습니다.. ", vbExclamation
        Set rslab301 = Nothing
        Exit Sub
    End If
    
    Call InitReport
    Call PrtHeader
    Call prtPageNum
    Call prtTerm

    rslab301.MoveFirst
    
    For i = 1 To rslab301.RecordCount ' 301의 각 labnum 마다
        
        If iCurY > iPageHeight - 4 * iCm Then ' newPage일 경우
            Printer.NewPage
            Call PrtHeader
            Call prtPageNum
            Call prtTerm
        End If
      
        Call ChangeLine(iCm / 2)
         
         iHeaderPosY1 = iCurY
         iHeaderPosY2 = iCurY + iCm / 2
    
        ' workseq, workno출력
        Printer.FontSize = 9
        Call WriteStr(iHeaderPosY1, iposSEQ, rslab301.Fields("workseq").Value, iHeaderPosY1, 0)
        sWorkno = DelFirst2Chr(rslab301.Fields("accdt").Value) & "-" & _
                 rslab301.Fields("accseq").Value
        Call WriteStr(iHeaderPosY1, iposWorkNo, sWorkno, iHeaderPosY1, 0)
        
        sWorkarea = rslab301.Fields("workarea").Value
        sAccDt = rslab301.Fields("accdt").Value
        iAccSeq = rslab301.Fields("accseq").Value
        
        
        '****** ptid, sex/age, deptcd, storecd, rcvdt, rcvtm, spccd 출력*******
        sSqlLab201 = " select ptid, sex, ageday,deptcd, storecd,rcvdt,statfg, " & _
                     "        rcvtm, spccd" & _
                     " from   " & T_LAB201 & _
                     " where  " & DBW("workarea", sWorkarea, 2) & _
                     " and    " & DBW("accdt", sAccDt, 2) & _
                     " and    " & DBW("accseq", iAccSeq, 2)
        Set rslab201 = New Recordset
        rslab201.Open sSqlLab201, dbconn
        
        If Trim(rslab201.Fields("statfg").Value) = "1" Then _
            Call WriteStr(iHeaderPosY2, iposSEQ, "**", iHeaderPosY2, 0)
        
        '****** ptnm,sex출력 ******************************************************
        ssqlHis001 = " select " & F_PTNM & " as ptnm, " & F_SEX & " as sex " & _
                     " from  " & T_HIS001 & _
                     " where " & DBW(F_PTID, rslab201.Fields("ptid").Value, 2)
        Set rshis001 = New Recordset
        rshis001.Open ssqlHis001, dbconn
        
        Call WriteStr(iHeaderPosY2, iposPtID, rslab201.Fields("ptid").Value, iHeaderPosY2, 0)
        S_Age = rshis001.Fields("sex").Value & ""
        
        If IsNumeric(S_Age) Then
            S_Age = Choose((Val(S_Age) Mod 2) + 1, "F", "M")
        End If
        
        S_Age = S_Age & "/" & ("" & rslab201.Fields("ageday").Value \ 365 + 1)
        Call WriteStr(iHeaderPosY2, iposSAge, S_Age, iHeaderPosY2, 0)
        Call WriteStr(iHeaderPosY2, iposIO, rslab201.Fields("deptcd").Value, iHeaderPosY2, 0)
        Call WriteStr(iHeaderPosY1, iposSF, rslab201.Fields("storecd").Value, iHeaderPosY1, 0)
        Call WriteStr(iHeaderPosY1, iposRcv, DelFirst2Chr(rslab201.Fields("rcvdt").Value), iHeaderPosY1, 0)
        Call WriteStr(iHeaderPosY1, iposRcv + iCm + iCm / 6, CvtTmFormat(rslab201.Fields("rcvtm").Value), iHeaderPosY1, 0)
        Call WriteStr(iHeaderPosY1, iposSpccd, rslab201.Fields("spccd").Value, iHeaderPosY1, 0)
        
        Call WriteStr(iHeaderPosY2, iposPtName, rshis001.Fields("ptnm").Value, iHeaderPosY2, 0)
        
        '****** 해당 Test Code 괄호(    )  출력 *******************************
        Call DrawBrackets(sWorkarea, sAccDt, iAccSeq)
        rslab301.MoveNext
    Next i
    Printer.EndDoc
    
    Set rslab301 = Nothing
    Set rslab201 = Nothing
    Set rshis001 = Nothing
End Sub

Private Function CvtTmFormat(sStr As String) As String
    Dim Time As String
    Dim Hour As String
    Dim Min As String
    
    Time = DelLast2Chr(sStr)
    Hour = DelLast2Chr(Time)
    Min = DelFirst2Chr(Time)
    
    CvtTmFormat = Hour & ":" & Min
    
End Function

Private Function DelFirst2Chr(sStr As String) As String
    DelFirst2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function DelLast2Chr(sStr As String) As String
    DelLast2Chr = Trim(Mid(sStr, 1, Len(sStr) - 2))
End Function

Public Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
End Sub
Public Sub PrtHeader()
   
    Dim Title As String
    Dim sWSNanme As String
    Dim iHeaderPosY As Integer
    
    '/* 보고서 제목
    Title = "업무나열서"
    
    
    '        "업무나열서"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    'TITLE = "WorkSheet : " & lblWSName.Caption
    
    sWSNanme = "WorkSheet  :  " & lblWSName.Caption
    
    Call WriteStr(iCurY, iPageWidth / 2 - Printer.TextWidth(sWSNanme) / 2, _
                  sWSNanme, iCurY, iCm * 1.5)
'    Call prtTitle(TITLE, iCm / 4)
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 6                       '284
    iposWorkNo = iposSEQ + iCm / 1.2        '757
    iposRcv = iposWorkNo + 1.8 * iCm    'iposXTestCD + 8.5 * iCm        '10038.5
    iposSpccd = iposRcv + 2.2 * iCm         '11286
    iposSF = iposSpccd + 0.8 * iCm             '4765
    
    iposPtName = iposWorkNo       '1891
    iposPtID = iposPtName + iCm + iCm / 1.3 '2894
    iposSAge = iposPtID + iCm + iCm / 2     '3744
    'iposIO = iposSAge + iCm                 '4311
    iposIO = iposSAge + iCm
    
    iposXTestCD = iposSF + iCm / 1.5     'iposWorkNo
    
    iHeaderPosY = iCurY
    
    Call WriteStr(iHeaderPosY, iposSEQ, "SEQ", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposWorkNo, "Work No", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposRcv, "검체도착시간", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSpccd - iCm / 6, "검체", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSF - iCm / 6, "S/F", iHeaderPosY, 0)
    
    iHeaderPosY = iHeaderPosY + iCm / 2
    
    'Call WriteStr(iHeaderPosY, iposPtName, "환자성명", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposPtID, "환자ID", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposSAge, "S/Age", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposIO, "I/O", iHeaderPosY, 0)
    
    'Call WriteStr(iCurY, iposXTestCD , "검사종목", iCurY, 0)
    Call DspWSTestCD(iHeaderPosY - iCm / 2, iposXTestCD)
    Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "solid", 2, 0)
    
    '-------------------------------------------------------------------------
'    iposSEQ = iCm / 2                       '284
'    iposWorkNo = iposSEQ + iCm / 1.2        '757
'    iposPtName = iposWorkNo + 2 * iCm       '1891
'    iposPtID = iposPtName + iCm + iCm / 1.3 '2894
'    iposSAge = iposPtID + iCm + iCm / 2     '3744
'    iposIO = iposSAge + iCm                 '4311
'    iposSF = iposIO + 0.8 * iCm             '4765
'    iposRcv = iposSF + 0.8 * iCm  'iposXTestCD + 8.5 * iCm        '10038.5
'    iposSpccd = iposRcv + 2.2 * iCm         '11286
'
'    iposXTestCD = iposWorkNo
'
'    iHeaderPosY = iCurY
'
'    Call WriteStr(iHeaderPosY, iposSEQ, "SEQ", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposWorkNo, "Work No", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposPtName, "환자성명", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposPtID, "환자ID", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposSAge, "S/Age", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposIO, "I/O", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposSF, "S/F", iHeaderPosY, 0)
'    'Call WriteStr(iCurY, iposXTestCD , "검사종목", iCurY, 0)
'    Call WriteStr(iHeaderPosY, iposRcv, "검체도착시간", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposSpccd, "검체", iHeaderPosY, 0)
'    Call DspWSTestCD(iHeaderPosY + iCm / 2, iposXTestCD)
'    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, 0)
    
    
End Sub

Public Sub PrtHeader900()
   
    Dim Title As String
    Dim sWSNanme As String
    Dim iHeaderPosY As Integer
    
    '/* 보고서 제목
    Title = "업무나열서"
    
    
    '        "업무나열서"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    'TITLE = "WorkSheet : " & lblWSName.Caption
    
    sWSNanme = "WorkSheet  :  " & lblWSName.Caption
    
    Call WriteStr(iCurY, iPageWidth / 2 - Printer.TextWidth(sWSNanme) / 2, _
                  sWSNanme, iCurY, iCm * 1.5)
'    Call prtTitle(TITLE, iCm / 4)
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 2                       '284
    iposWorkNo = iposSEQ + iCm / 1.2        '757
    iposPtName = iposWorkNo + 2 * iCm       '1891
    iposPtID = iposPtName + iCm + iCm / 1.3 '2894
    iposSAge = iposPtID + iCm + iCm / 2     '3744
    iposIO = iposSAge + iCm                 '4311
    iposSF = iposIO + 0.8 * iCm             '4765
    iposRcv = iposSF + 0.8 * iCm  'iposXTestCD + 8.5 * iCm        '10038.5
    iposSpccd = iposRcv + 2.2 * iCm         '11286
    
    iposXTestCD = iposWorkNo
    
    iHeaderPosY = iCurY
    
    Call WriteStr(iHeaderPosY, iposSEQ, "SEQ", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposWorkNo, "Work No", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposPtName, "환자성명", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposPtID, "환자ID", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSAge, "S/Age", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposIO, "I/O", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSF, "S/F", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposRcv, "검체도착시간", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSpccd, "검체", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY + iCm / 2, iposXTestCD, "검사종목", iCurY, iCm / 2)
    'Call DspWSTestCD(iHeaderPosY + iCm / 2, iposXTestCD)
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, 0)
    
End Sub


Private Sub DspWSTestCD(ByVal iPosY As Integer, ByVal iPosX As Integer)
    Dim sSQL As String
    Dim rsWSTestCD As Recordset
    Dim i As Integer, iOldy As Integer, iOldX As Integer
    
    
    iOldy = iPosY
    iOldX = iPosX
    iWidthTestCD = 1.5 * iCm
    
    sSQL = " select testcd " & _
           " from   " & T_LAB008 & _
           " where  " & DBW("workcd", Trim(txtWSCode.Text), 2) & _
           " order by 1 asc "
            
    Set rsWSTestCD = New Recordset
    rsWSTestCD.Open sSQL, dbconn
    
    rsWSTestCD.MoveFirst
    ReDim aWSTestCD(1 To rsWSTestCD.RecordCount)
    
    For i = 1 To rsWSTestCD.RecordCount
        aWSTestCD(i) = rsWSTestCD.Fields("testcd").Value
        Call WriteStr(iPosY, iPosX, rsWSTestCD.Fields("testcd").Value, iCurY, 0)
        rsWSTestCD.MoveNext
        iPosX = iPosX + iWidthTestCD
        'If i Mod 10 = 0 Then
        '    iPosY = iPosY + iCm / 2
        '    iPosX = iOldX
        'End If
        If i Mod 20 = 0 Then
            Call DrawLine(iOldX, iPosY + iCm / 2.5, iposRcv - iCm / 3, _
                            iPosY + iCm / 2.5, "dot", 1, 1)
            iPosY = iPosY + iCm / 2
            iPosX = iOldX
        End If
    Next i
    
    'If rsWSTestCD.RecordCount Mod 16 > 0 Then
    '  iCurY = iPosY + iCm
    'Else
      iCurY = iPosY + iCm / 2
   'End If
    
    Set rsWSTestCD = Nothing
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
    Dim sSQLGetWAName As String
    Dim rsWAName As Recordset
    Dim sWorkDate As String
    
    
    sSQLGetWAName = " select distinct WS.workarea, WA.field1 as WAName" & _
                    " from  " & T_LAB032 & " WA, " & T_LAB008 & " WS " & _
                    " where   WA.cdval1 = WS.workarea" & _
                    " and   " & DBW("WA.cdindex", LC3_WorkArea, 2) & _
                    " and   " & DBW("WS.workcd", txtWSCode.Text, 2)

    Set rsWAName = New Recordset
    rsWAName.Open sSQLGetWAName, dbconn
    
    If rsWAName.EOF = True Then Exit Sub
        
    sWorkDate = Format(dtpWorkDate.Value, CS_DateLongFormat)
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "작 업 일  : " & sWorkDate
    
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ")
    Printer.Print "WorkArea  : " & rsWAName.Fields("workarea").Value & "   " & rsWAName.Fields("WAName").Value
                                    
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ") + _
                       Printer.TextHeight("workArea : ")
    Printer.Print "Work No   : " & txtStartWNum.Text & " - " & txtEndWNum.Text
    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
    Set rsWAName = Nothing
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

Private Sub DspBrackketTestCD(sWorkarea As String, sAccDt As String, iAccSeq As Integer)
    Dim sSqlLab302 As String
    Dim rslab302 As Recordset
    Dim aBracketsTestCD() As Integer
    Dim i%, j%
    Dim tmpiposXTestCD%
    Dim iXcnt%, iYcnt%, iOldy%
    Dim iWidthTestCD900%
    
    Call ChangeLine(iCm / 2)
    
    sSqlLab302 = " select testcd " & _
                 " from   " & T_LAB302 & _
                 " where  " & DBW("workarea", sWorkarea, 2) & _
                 " and    " & DBW("accdt", sAccDt, 2) & _
                 " and    " & DBW("accseq", iAccSeq, 2) & _
                 " order by 1 asc "
        
    Set rslab302 = New Recordset
    rslab302.Open sSqlLab302, dbconn
        
    rslab302.MoveFirst
    
    tmpiposXTestCD = iposXTestCD
    iOldy = iCurY
    
    iWidthTestCD900 = iCm * 2.3
    
    For i = 1 To rslab302.RecordCount
        
        iYcnt = i \ 9   ' 몫
        iXcnt = i Mod 8 '나머지

        iCurY = iOldy + iYcnt * (iCm / 2)

        If iXcnt = 0 Then
            tmpiposXTestCD = iposXTestCD + iWidthTestCD900 * 7
        Else
            tmpiposXTestCD = iposXTestCD + iWidthTestCD900 * (iXcnt - 1)
        End If

        Call WriteStr(iCurY, tmpiposXTestCD, Trim(rslab302.Fields("testcd").Value) & _
                     " (       )", iCurY, 0)
        rslab302.MoveNext
    Next i
        
    
    Call ChangeLine(iCm / 2)
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "dot", 2, 0)
    
    Set rslab302 = Nothing
End Sub

Private Sub DrawBrackets(sWorkarea As String, sAccDt As String, iAccSeq As Integer)
    
    Dim sSqlLab302 As String
    Dim rslab302 As Recordset
    Dim aBracketsTestCD() As Integer
    Dim i%, j%
    Dim tmpiposXTestCD%
    Dim iXcnt%, iYcnt%, iOldy%
    Dim strRemark As String
    
    'Call ChangeLine(iCm / 2)
    
    sSqlLab302 = " select a.testcd, b.mesg " & _
                 " from   " & T_LAB302 & " a, " & T_LAB102 & " b " & _
                 " where  " & DBW("a.workarea", sWorkarea, 2) & _
                 " and    " & DBW("a.accdt", sAccDt, 2) & _
                 " and    " & DBW("a.accseq", iAccSeq, 2) & _
                 " and    b.ptid = a.ptid " & _
                 " and    b.orddt = a.orddt " & _
                 " and    b.ordno = a.ordno " & _
                 " and    b.ordseq = a.ordseq "
                    '"     order by 1 asc "
        
    Set rslab302 = New Recordset
    rslab302.Open sSqlLab302, dbconn
        
    rslab302.MoveFirst
        
    strRemark = ""
    ReDim aBracketsTestCD(1 To rslab302.RecordCount)
    For i = 1 To rslab302.RecordCount
        For j = 1 To UBound(aWSTestCD)
            If Trim(rslab302.Fields("testcd").Value) = Trim(aWSTestCD(j)) Then
                aBracketsTestCD(i) = j
                If Trim(rslab302.Fields("mesg").Value) <> "" Then
                     strRemark = strRemark & Trim(rslab302.Fields("mesg").Value)
                     strRemark = Replace(strRemark, Chr(10), ",")
                     strRemark = Replace(strRemark, Chr(13), ",")
                End If
                Exit For
            End If
        Next j
        rslab302.MoveNext
    Next i
    

    tmpiposXTestCD = iposXTestCD
    iOldy = iCurY

    For i = 1 To UBound(aBracketsTestCD)
        'If iCurY > iPageHeight - 2 * iCm Then  ' newPage일 경우
        '    Printer.NewPage
        '    Call PrtHeader
        '    Call prtPageNum
        '    Call prtTerm
        'End If

        If aBracketsTestCD(i) <> 0 Then
        
            iYcnt = aBracketsTestCD(i) \ 20   ' 몫
            iXcnt = aBracketsTestCD(i) Mod 20 '나머지
        
            iCurY = iOldy + iYcnt * (iCm / 2)
        
            If iXcnt = 0 Then
                tmpiposXTestCD = iposXTestCD + iWidthTestCD * 4
            Else
                tmpiposXTestCD = iposXTestCD + iWidthTestCD * (iXcnt - 1)
            End If
        
            'If aBracketsTestCD(i) Mod 10 = 0 Then
            '   Call WriteStr(iCurY + iCm / 2, tmpiposXTestCD, "(       )", iCurY, 0)
            'Else
               Call WriteStr(iCurY, tmpiposXTestCD, "(       )", iCurY, 0)
            'End If
            
        End If
    Next i
    
    iOldy = iCurY
    
    Printer.FontSize = 8
    Call WriteStr(iCurY + iCm / 1.8, iposXTestCD, strRemark, iCurY, 0)
    
    iCurY = iOldy
    
    Printer.FontSize = 9
    Call ChangeLine(iCm)
    'Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "dot", 2, 0)
    Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "dash", 2, 0)
    
    Set rslab302 = Nothing
End Sub

Private Sub txtWSCode_LostFocus()
   If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
   If ActiveControl.Name = cmdReport.Name Then Exit Sub
   If ActiveControl.Name = cmdExit.Name Then Exit Sub
   If ActiveControl.Name = lstWSCode.Name Then Exit Sub
   Call txtWSCode_KeyPress(vbKeyReturn)
End Sub
