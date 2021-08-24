VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS965 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "헌혈자 DM발송"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   1650
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   7
      Text            =   "frmBBS965.frx":0000
      Top             =   2130
      Visible         =   0   'False
      Width           =   8955
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "헌혈자 DM발송"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   75
      TabIndex        =   1
      Top             =   285
      Width           =   10770
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9255
         Style           =   1  '그래픽
         TabIndex        =   2
         Tag             =   "124"
         Top             =   465
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpTo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   2790
         TabIndex        =   3
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   66846723
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   66846723
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   105
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   420
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
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
         Caption         =   "조회기간"
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2580
         TabIndex        =   5
         Tag             =   "40304"
         Top             =   480
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6720
      Left            =   75
      TabIndex        =   6
      Top             =   1545
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   11853
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   21
      MaxRows         =   27
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   13818331
      SpreadDesigner  =   "frmBBS965.frx":0054
      TextTip         =   4
   End
End
Attribute VB_Name = "frmBBS965"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tblCol
    tcChk = 1
    tcDonorNm
    tcDonationdt
    tcBldNum
    tcCompo
    
    tcABO
    TcAddr
    tcrmk
    tcZipCd
    tcDonorid
    
    tcDonorAccdt
    tcWorkarea
    tcAccdt
    tcAccseq
    tcBtest
    
    tcCtest
    tcALT
    tcMa
    tcAddr1
    tcAddr2
    
    tcRmkTxt
End Enum





Private Sub cmdClear_Click()
    dtpFrom.Value = GetSystemDate
    dtpTo.Value = GetSystemDate
    Call medClearTable(tblList)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean
    Dim ii As Integer
    
    
    With tblList
        .Row = Row
        .Col = tblCol.tcDonorNm
        If .Value = "" Then Exit Sub
        Select Case Col
            Case tblCol.tcrmk
               Ret = .GetCellPos(tblCol.tcDonationdt, Row, X, Y, Wdt, Hgt)
               If Row <> .DataRowCnt Then
                    Y = Y + Hgt
               Else
                    Y = Y ' + 200
               End If
               
               If .Height - Y < txtComment.Height Or Y < 0 Then
                      Ret = .GetCellPos(tblCol.tcDonationdt, Row, X, Y, Wdt, Hgt)
                      txtComment.Top = .Top + Y - txtComment.Height + 1065 + 950
                      txtComment.Left = .Left + X - 10

               Else
                  txtComment.Left = .Left + X - 10
                  txtComment.Top = .Top + Y
               End If
               .Col = tblCol.tcRmkTxt
               txtComment.Text = .Value
               txtComment.Tag = Row
               txtComment.Visible = True
               txtComment.SetFocus
        End Select
    
    End With

End Sub

Private Sub txtComment_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        With tblList
            .Row = txtComment.Tag
            .Col = tblCol.tcRmkTxt
            .Value = txtComment
            If .Value <> "" Then
                .Col = tblCol.tcrmk
                .Value = "Y": .ForeColor = vbRed: .FontBold = True
                .TypeHAlign = TypeHAlignCenter
            Else
                .Col = tblCol.tcrmk
                .Value = ""
            End If
        End With
        txtComment.Visible = False
    End If
End Sub
Private Sub Form_Load()
    dtpFrom.Value = GetSystemDate
    dtpTo.Value = GetSystemDate
    medClearTable tblList
End Sub

Private Sub cmdQuery_Click()
    Dim RS      As Recordset
    
    Dim SSQL    As String
    Dim sFDate  As String
    Dim sTDate  As String
    
    
    Call medClearTable(tblList)
    
    sFDate = Format(dtpFrom.Value, PRESENTDATE_FORMAT)
    sTDate = Format(dtpTo.Value, PRESENTDATE_FORMAT)
        
    Me.MousePointer = 11
    
    
    
    SSQL = "SELECT c.abo,c.rh,c.zipcd,c.addr1,c.addr2,c.donornm," & _
           " a.bldsrc,a.bldyy,a.bldno,a.compocd,a.donorid,a.donoraccdt,a.donationdt," & _
           " e.abbrnm as componm FROM " & _
           T_BBS006 & " e," & T_BBS401 & " d," & T_BBS601 & " c," & T_BBS602 & " a" & _
           " WHERE " & _
           DBW("a.donationdt>=", sFDate) & _
           " AND " & DBW("a.donationdt<=", sTDate) & _
           " AND a.donorid=c.donorid " & _
           " AND a.bldsrc=d.bldsrc AND a.bldyy=d.bldyy AND a.bldno=d.bldno AND a.compocd=d.compocd" & _
           " AND a.compocd=e.compocd " & _
           " AND " & DBW("d.stscd<>", BBSBloodStatus.stsEXPIRE) & _
           " order by a.donorid,a.donationdt"
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    If RS.EOF Then
        MsgBox "DM 발송내역이 없습니다.", vbInformation + vbOKOnly, "Info"
        Set RS = Nothing
        Me.MousePointer = 0
        Exit Sub
    End If

    
    With tblList
        .ReDraw = False
        Do Until RS.EOF
            If .DataRowCnt + 1 >= .MaxRows Then
                .MaxRows = .MaxRows + 1
            End If
            .Row = .DataRowCnt + 1
            .Col = tblCol.tcDonorNm:    .Value = RS.Fields("donornm").Value & ""
            .Col = tblCol.tcDonationdt: .Value = Format(RS.Fields("donationdt").Value & "", "####-##-##")
'            .Col = tblCol.tcWorkarea:   .Value = RS.Fields("workarea").Value & ""
'            .Col = tblCol.tcAccdt:      .Value = RS.Fields("accdt").Value & ""
'            .Col = tblCol.tcAccseq:     .Value = RS.Fields("accseq").Value & ""
            .Col = tblCol.TcAddr:       .Value = RS.Fields("addr1").Value & "" & RS.Fields("addr2").Value & ""
            .Col = tblCol.tcAddr1:      .Value = RS.Fields("addr1").Value & ""
            .Col = tblCol.tcAddr2:      .Value = RS.Fields("addr2").Value & ""
            .Col = tblCol.tcZipCd:      .Value = RS.Fields("zipcd").Value & ""
            .Col = tblCol.tcBldNum:     .Value = RS.Fields("bldsrc").Value & "" & "-" & _
                                                 RS.Fields("bldyy").Value & "" & "-" & _
                                                 Format(RS.Fields("bldno").Value & "", "00000#")
            .Col = tblCol.tcCompo:      .Value = RS.Fields("componm").Value & ""
            
            .Col = tblCol.tcDonorid:    .Value = RS.Fields("donorid").Value & ""
            .Col = tblCol.tcDonorAccdt: .Value = RS.Fields("donoraccdt").Value & ""
            .Col = tblCol.tcABO:        .Value = RS.Fields("abo").Value & "" & RS.Fields("rh").Value & ""

            
            RS.MoveNext
        Loop

        .ReDraw = True
    End With
           
    Me.MousePointer = 0
    Set RS = Nothing
End Sub

Private Function GetTestResullt(ByVal Workarea As String, ByVal accdt As String, ByVal accseq As String) As String
    Dim SSQL As String
    
    SSQL = " SELECT rstcd,'B60D9E' as div FROM " & T_LAB302 & _
               " WHERE " & _
                     DBW("workarea=", Workarea) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq) & _
           " AND " & DBW("testcd=", "B60D9E")
    SSQL = SSQL & " UNION ALL SELECT rstcd,'B530' as div FROM " & T_LAB302 & _
               " WHERE " & _
                     DBW("workarea=", Workarea) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq) & _
           " AND " & DBW("testcd=", "B530")
    SSQL = SSQL & "  UNION ALL  SELECT rstcd,'B258' as div FROM " & T_LAB302 & _
               " WHERE " & _
                     DBW("workarea=", Workarea) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq) & _
           " AND " & DBW("testcd=", "B258")
    SSQL = SSQL & "  UNION ALL  SELECT rstcd,'B453' as div FROM " & T_LAB302 & _
           " WHERE " & _
                     DBW("workarea=", Workarea) & _
           " AND " & DBW("accdt=", accdt) & _
           " AND " & DBW("accseq=", accseq) & _
           " AND " & DBW("testcd=", "B453")
    GetTestResullt = SSQL
End Function

Private Sub cmdPrint_Click()
    Dim RS          As Recordset
    Dim sRs         As Recordset
    
    Dim sDonorID    As String
    Dim sDonorAccDt As String
    Dim tmpResult   As String
    
    
    Dim Addr1       As String
    Dim Addr2       As String
    Dim ZipCd       As String
    Dim DonorNm     As String
    Dim BldNum      As String
    Dim DonationDt  As String
    Dim CompoNm     As String
    Dim BTest       As String
    Dim CTest       As String
    Dim ALT         As String
    Dim Madoc       As String
    Dim Rmk         As String
    
    Dim ii          As Integer
    
    Me.MousePointer = 11
    
    Call PrintIntionlize
    
    With tblList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = tblCol.tcChk
            If .Value = 0 Then
                .Col = tblCol.tcDonorid:    sDonorID = .Value
                .Col = tblCol.tcDonorAccdt: sDonorAccDt = .Value
                
                Set RS = GetWorkareaAccdtAccSeq(sDonorID, sDonorAccDt)
                If Not RS.EOF Then
                    Do Until RS.EOF
                        Set sRs = Nothing
                        Set sRs = New Recordset
                        RS.Open GetTestResullt(RS.Fields("workarea").Value & "", _
                                                             RS.Fields("accdt").Value & "", _
                                                             RS.Fields("accseq").Value & ""), dbconn
                        If Not sRs.EOF Then
                            Do Until sRs.EOF
                                
                                
                                Select Case sRs.Fields("div").Value & ""
                                    Case "B60D9E": .Col = tblCol.tcBtest
                                            tmpResult = UCase(sRs.Fields("rstcd").Value & "")
                                            If tmpResult = "-" Or tmpResult = "NEGATIVE" Then
                                                .Value = "음성"
                                            Else
                                                .Value = "양성"
                                            End If
                                    Case "B530": .Col = tblCol.tcCtest
                                            tmpResult = UCase(sRs.Fields("rstcd").Value & "")
                                            If tmpResult = "-" Or tmpResult = "NEGATIVE" Then
                                                .Value = "음성"
                                            Else
                                                .Value = "양성"
                                            End If
                                    Case "B258": .Col = tblCol.tcALT
                                            tmpResult = UCase(sRs.Fields("rstcd").Value & "")
                                            If tmpResult >= 3 And tmpResult <= 45 Then
                                                .Value = "음성"
                                            Else
                                                .Value = "양성"
                                            End If
                                    Case "B453": .Col = tblCol.tcMa
                                            tmpResult = UCase(sRs.Fields("rstcd").Value & "")
                                            If tmpResult <> "N" Then
                                                .Value = "음성"
                                            Else
                                                .Value = "양성"
                                            End If
                                End Select
                                
'                                If sRs.Fields("div").Value & "" = "B258" Then
'                                Else
'                                    tmpResult = medGetP(UCase(sRs.Fields("rstcd").Value & ""), 1, "/")
'
'                                    If tmpResult = "N" Or tmpResult = "NAGATIVE" Or tmpResult = "-" Then
'                                        .Value = "음성"
'                                    Else
'                                        .Value = "양성"
'                                    End If
'                                End If
                                sRs.MoveNext
                            Loop
                        End If
                        
                        RS.MoveNext
                    Loop
                End If
                .Col = tblCol.tcAddr1: Addr1 = .Value
                .Col = tblCol.tcAddr2: Addr2 = .Value
                .Col = tblCol.tcZipCd: ZipCd = .Value
                .Col = tblCol.tcDonorNm: DonorNm = .Value
                .Col = tblCol.tcBldNum: BldNum = .Value
                .Col = tblCol.tcDonationdt: DonationDt = .Value
                .Col = tblCol.tcCompo: CompoNm = .Value
                .Col = tblCol.tcBtest: BTest = .Value
                .Col = tblCol.tcCtest: CTest = .Value
                .Col = tblCol.tcALT: ALT = .Value
                .Col = tblCol.tcMa: Madoc = .Value
                .Col = tblCol.tcRmkTxt: Rmk = .Value
                
                Call PrintDM(Addr1, Addr2, ZipCd, DonorNm, _
                             BldNum, DonationDt, CompoNm, _
                             BTest, CTest, ALT, Madoc, _
                             Rmk)
            End If
        Next
        
    End With
    Me.MousePointer = 0

End Sub
Private Function GetWorkareaAccdtAccSeq(ByVal donorid As String, ByVal donoraccdt As String) As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT workarea,accdt,accseq FROM " & T_BBS605 & _
           " WHERE " & _
                     DBW("donorid=", donorid) & " AND " & DBW("donoraccdt=", donoraccdt)
    
    Set GetWorkareaAccdtAccSeq = New Recordset
    GetWorkareaAccdtAccSeq.Open SSQL, dbconn
End Function
Private Sub PrintIntionlize()
    Printer.Font = "굴림체"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
End Sub
Private Sub PrintDM(ByVal Addr1 As String, ByVal Addr2 As String, ByVal ZipCd As String, ByVal DonorNm As String, _
                    ByVal BldNum As String, ByVal DonationDt As String, ByVal CompoNm As String, _
                    ByVal BTest As String, ByVal CTest As String, ByVal ALT As String, ByVal Madoc As String, _
                    ByVal Rmk As String)
    Dim strTmp      As String
    Dim aryTmp()    As String
    
    Dim ii          As Long
    
    strTmp = Rmk
    
'    strTmp = "김정규 님께서 헌혈하신 혈액으로 B형 및 C형간염,간기능(ALT),매독검사에 대하여 검사한" & vbCrLf & _
           "결과모두 정상이였으며, ABO혈액형은 O형,RH식 혈액형은 positive로 판정되었습니다." & vbCrLf & _
           "소중한 생명을 살리는 헌혈에 참여하여 주신데 대하여 다시한번 감사드리며, 앞으로도 지속적" & vbCrLf & _
           "인 관심과 참여를 부탁드립니다."
    
    
    lngCurYPos = 40
    Call Print_Setting(Addr1, 85, 6, , "L")
    Call Print_Setting(Addr2, 85, 6, , "L")
    Call Print_Setting(ZipCd, 125, 6, , "L")
    Call Print_Setting(DonorNm & "  귀하", 105, 6, , "L")
    
    lngCurYPos = 92
    Call Print_Setting(BldNum, 20, 6, , "L", , False)
    Call Print_Setting(DonorNm, 60, 6, , "L", , False)
    Call Print_Setting(DonationDt, 100, 6, , "L", , False)
    Call Print_Setting(CompoNm, 140, 6, , "L", , False)
    
    lngCurYPos = 106
    Call Print_Setting(BTest, 56, 20, 15, "C", "C")
    Call Print_Setting(CTest, 56, 18, 15, "C", "C")
    Call Print_Setting(ALT, 56, 18, 15, "C", "C")
    Call Print_Setting(Madoc, 56, 20, 15, "C", "C")
    lngCurYPos = 190
    Call Print_Setting("헌혈에 참여하여 주셔서 감사합니다.", 18, 6, 15, "L")

    If strTmp <> "" Then
        aryTmp = Split(strTmp, vbCrLf)
        For ii = LBound(aryTmp) To UBound(aryTmp)
            Call Print_Setting(aryTmp(ii), 18, 6, 15, "L")
        Next
    End If
    
    Printer.EndDoc
    
End Sub

