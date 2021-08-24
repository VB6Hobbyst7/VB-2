VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS923 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "ÇåÇ÷ÀÚ´ëÀå"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS923.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "ÇåÇ÷ÀÚ ´ëÀå"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   825
      Left            =   75
      TabIndex        =   5
      Top             =   300
      Width           =   10770
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Á¶È¸(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   7
         Tag             =   "124"
         Top             =   210
         Width           =   1320
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
         Left            =   1170
         TabIndex        =   6
         Top             =   270
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   62193667
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Á¶È¸ ÀÏÀÚ"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Ãâ·Â(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Á¾·á(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7125
      Left            =   75
      TabIndex        =   0
      Top             =   1140
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   12568
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   13
      MaxRows         =   16
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS923.frx":076A
      TextTip         =   4
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   9000
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "µ¸¿òÃ¼"
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
      TabIndex        =   3
      Tag             =   "40304"
      Top             =   1020
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmBBS923"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblColumn
    tcChk = 1
    tcBldNo
    TcDonorDt
    tcABO
    TcNm
    tcSSN
    TcAge
    TcSex
    TcJob
    TcAddr
    TcRsn
    TcJudge
    tcrmk
End Enum

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdQuery_Click()
    Dim objPrt As New clsStatics
    Dim RS     As New Recordset
    Dim objPro As clsProgress
    Dim Fdt    As String
    Dim Tdt    As String
    Dim ii     As Integer
    
    
    Fdt = Format(dtpFrom.Value, "YYYYMM") & "01"
    Tdt = Format(dtpFrom.Value, "YYYYMM") & "31"
    
'    Call medClearTable(tblList)
    
    With tblList
        .MaxRows = 16
        For ii = 3 To 16
            .RowHeight(ii) = 19.77
        Next
    End With

    Set RS = objPrt.Get_DonorPaper(Fdt, Tdt)
    If Not RS.EOF Then
        Set objPro = New clsProgress
'        Set objPro.StatusBar = MainFrm.stsbar
        objPro.Container = MainFrm.stsbar
        objPro.Max = RS.RecordCount
        
        With tblList
            ii = 3
            .MaxRows = RS.RecordCount + 2
            Do Until RS.EOF
                .Row = ii
                .RowHeight(ii) = 20
                .Col = tblColumn.tcBldNo:   .Value = RS.Fields("bldsrc").Value & "" & "-" & _
                                                     RS.Fields("bldyy").Value & "" & "-" & _
                                                     Format(RS.Fields("bldno").Value & "", "000000")
                .Col = tblColumn.TcDonorDt: .Value = Format(RS.Fields("donationdt").Value & "", "####/##/##") & " " & RS.Fields("volumn").Value & "" & "cc"
                
                .Col = tblColumn.tcABO:     .Value = RS.Fields("abo").Value & "" & RS.Fields("rh").Value & ""
                .Col = tblColumn.TcNm:      .Value = RS.Fields("donornm").Value & ""
                
                .Col = tblColumn.tcSSN:     .Value = Mid(RS.Fields("ssn").Value & "", 1, 6)
                                            If Len(RS.Fields("ssn").Value & "") > 6 Then
                                                .Value = .Value & "-" & Mid(RS.Fields("ssn").Value & "", 6)
                                            End If
                .Col = tblColumn.TcSex:     .Value = RS.Fields("sex").Value & ""
                .Col = tblColumn.TcAge:     .Value = medFindAge(Mid(RS.Fields("dob").Value & "", 3), "Y")
                .Col = tblColumn.TcJob:     .Value = RS.Fields("field1").Value & ""
                .Col = tblColumn.TcAddr:    .Value = RS.Fields("addr1").Value & "" & vbTab & RS.Fields("addr2").Value & ""
                
                If Len(RS.Fields("exprsncd")) > 0 Then
                    .Col = tblColumn.TcRsn: .Value = objPrt.Get_RsnNm(RS.Fields("exprsncd").Value & "") & vbTab & Format(RS.Fields("realexpdt").Value & "", "####/##/##")
                End If
                .Col = tblColumn.TcJudge:   .Value = Format(RS.Fields("okdt3").Value & "", "####/##/##")
                objPro.Value = ii - 2
                ii = ii + 1
                RS.MoveNext
            Loop
            cmdPrint.Enabled = True
            
            If .MaxRows < 16 Then
                .MaxRows = 16
                For ii = 3 To 16
                    .RowHeight(ii) = 19.77
                Next
            End If
        End With
    Else
        cmdPrint.Enabled = False
        tblList.MaxRows = 2
    End If
    
    Set RS = Nothing
    Set objPro = Nothing
    Set objPrt = Nothing
    
End Sub

Private Sub Form_Load()
    dtpFrom = Format(GetSystemDate, "yyyy-MM")
    cmdPrint.Enabled = False
    tblList.MaxRows = 2
End Sub
Private Sub cmdPrint_Click()
    Printer.Font = "±¼¸²Ã¼"
    Printer.FontSize = 10
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORLandscape '/* Á¼°Ô
    Printer.ScaleMode = vbMillimeters
    Call VBPrinter
End Sub
Private Sub PrinterHeader()
    Dim lngX        As Long
    Dim lngY        As Long
    Dim lngcurY     As Long
    Dim lngLineY    As Long
    Dim ii          As Integer
    
    lngX = 4: lngY = 10: lngcurY = 6
    
    lngCurYPos = 10
    
    Printer.FontSize = 20:    Printer.FontBold = True
    Call Print_Setting("Çå Ç÷ ÀÚ ´ë Àå", 0, lngcurY * 2, Printer.ScaleWidth)
    
    Printer.FontSize = 8:    Printer.FontBold = False
    Call Print_Setting("Ç÷¾×¿ø ¸íÄª: ", 180, lngcurY, lngcurY, "L")
    Call Print_Setting("Ç÷¾×¿ø ¸íÄª: ", 180, lngcurY, lngcurY, "L")
    
    lngCurYPos = 45
    
    lngLineY = lngCurYPos
    Printer.DrawWidth = 4
    Printer.Line (lngX, lngCurYPos)-(284, 185), , B
    
    For ii = 1 To 13
        lngLineY = lngLineY + 10
        If ii = 1 Then
            Printer.Line (lngX, lngLineY)-(65, lngLineY)
            Printer.Line (75, lngLineY)-(260, lngLineY)
        Else
            Printer.Line (lngX, lngLineY)-(284, lngLineY)
        End If
    Next
    
    Printer.DrawWidth = 4
    Printer.Line (40, lngCurYPos)-(40, 185)
    Printer.Line (65, lngCurYPos)-(65, 185)
    Printer.Line (75, lngCurYPos)-(75, 185)
    Printer.Line (95, 55)-(95, 185)
    Printer.Line (130, 55)-(130, 185)
    Printer.Line (140, 55)-(140, 185)
    Printer.Line (155, 55)-(155, 185)
    Printer.Line (220, lngCurYPos)-(220, 185)
    Printer.Line (240, lngCurYPos)-(240, 185)
    Printer.Line (260, lngCurYPos)-(260, 185)
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbDashDotDot
    lngLineY = 60
    For ii = 1 To 12
        lngLineY = lngLineY + 10
        Printer.Line (40, lngLineY)-(65, lngLineY)
        Printer.Line (220, lngLineY)-(260, lngLineY)
    Next
    Call Print_Setting("ÇåÇ÷Áõ¼­¹øÈ£", lngX, lngY, 36, , , False)
    Call Print_Setting("ÇåÇ÷·®", 40, lngY, 25, , , False)
    Call Print_Setting("Ç÷¾×Çü", 65, (lngY) * 2, 10, , , False)
    Call Print_Setting("Çå Ç÷ ÀÚ ÀÇ ÀÎ Àû »ç Ç×", 75, lngY, 145, , , False)
    Call Print_Setting("Æó±â»çÀ¯", 220, lngY, 20, , , False)
    Call Print_Setting("°ø ±Þ Ã³", 240, lngY, 20, , , False)
    Call Print_Setting("ºñ  °í", 260, (lngY) * 2, 20, , , False)
    Call Print_Setting("", lngX, lngY)
    Call Print_Setting("Ç÷¾×¹øÈ£", lngX, lngY, 36, , , False)
    Call Print_Setting("ÇåÇ÷³â¿ùÀÏ", 40, lngY, 25, , , False)
    Call Print_Setting("¼º  ¸í", 75, lngY, 20, , , False)
    Call Print_Setting("ÁÖ¹Îµî·Ï¹øÈ£", 95, lngY - 5, 35)
    Call Print_Setting("¶Ç´Â »ý³â¿ùÀÏ", 95, lngY - 5, 35, , , False)
    Call Print_Setting("", lngX, -5)
    Call Print_Setting("¼ºº°", 130, lngY, 10, , , False)
    Call Print_Setting("Á÷¾÷", 140, lngY, 15, , , False)
    Call Print_Setting("ÁÖ              ¼Ò", 155, lngY, 65, , , False)
    Call Print_Setting("Æó±â³â¿ùÀÏ", 220, lngY, 20, , , False)
    Call Print_Setting("°ø±Þ³â¿ùÀÏ", 240, lngY, 20)
End Sub

Private Sub VBPrinter()
    Dim lngX        As Long
    Dim lngY        As Long
    Dim ii          As Integer
    Dim lngCnt      As Integer
    
    lngX = 4: lngY = 10:
    With tblList
        If .MaxRows < 3 Then Exit Sub
        Call PrinterHeader
        For ii = 3 To .MaxRows
            .Row = ii
            .Col = tblColumn.tcChk
            If .Value = 0 Then
                lngCnt = lngCnt + 1
                If lngCnt > 12 And lngCnt Mod 12 = 1 Then
                    Printer.NewPage
                    Call PrinterHeader
                End If
                
                .Col = tblColumn.tcBldNo:   Call Print_Setting(.Value, lngX, lngY, 36, , , False)
                .Col = tblColumn.TcDonorDt: Call Print_Setting(medGetP(.Value, 2, " "), 40, lngY - 5, 25)
                .Col = tblColumn.TcDonorDt: Call Print_Setting(medGetP(.Value, 1, " "), 40, lngY - 5, 25, , , False)
                                            Call Print_Setting("", lngX, -5)
                .Col = tblColumn.tcABO:     Call Print_Setting(.Value, 65, lngY, 10, , , False)
                .Col = tblColumn.TcNm:      Call Print_Setting(.Value, 75, lngY, 20, , , False)
                .Col = tblColumn.tcSSN:     Call Print_Setting(.Value, 95, lngY, 35, , , False)
                .Col = tblColumn.TcSex:     Call Print_Setting(.Value, 130, lngY, 10, , , False)
                .Col = tblColumn.TcJob:     Call Print_Setting(.Value, 140, lngY, 15, , , False)
                
                .Col = tblColumn.TcAddr:    Call Print_Setting(medGetP(.Value, 1, vbTab), 155, lngY - 5, 65)
                                            Call Print_Setting(medGetP(.Value, 2, vbTab), 155, lngY - 5, 65, , , False)
                                            Call Print_Setting("", lngX, -5)
                .Col = tblColumn.TcRsn:
                If .Value <> "" Then
                    Call Print_Setting(medGetP(.Value, 1, vbTab), 220, lngY - 5, 20)
                    Call Print_Setting(medGetP(.Value, 2, vbTab), 220, lngY - 5, 20, , False)
                    Call Print_Setting("", lngX, -lngY)
                End If
                                            Call Print_Setting(HOSPITAL_MAIN, 240, lngY - 5, 20)
                .Col = tblColumn.TcJudge:   Call Print_Setting(.Value, 240, lngY - 5, 20, , , False)
                                            Call Print_Setting("", lngX, -5)
                .Col = tblColumn.tcrmk:     Call Print_Setting(.Value, 260, lngY - 5, 20, , , False)
                                            Call Print_Setting("", lngX, lngY)
            End If
        
        Next
    End With
    
    Printer.EndDoc
End Sub

