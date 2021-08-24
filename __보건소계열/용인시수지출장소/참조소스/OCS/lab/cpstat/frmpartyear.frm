VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPartYear 
   Caption         =   "검사통계(년간)"
   ClientHeight    =   6765
   ClientLeft      =   150
   ClientTop       =   870
   ClientWidth     =   11475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11475
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
      _ExtentY        =   1296
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.ComboBox cmbYear 
         Height          =   300
         ItemData        =   "frmPartYear.frx":0000
         Left            =   1350
         List            =   "frmPartYear.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   135
         Width           =   1365
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   510
         Left            =   9450
         TabIndex        =   4
         Top             =   90
         Width           =   1680
         Caption         =   "출력"
         PicturePosition =   327683
         Size            =   "2963;900"
         Picture         =   "frmPartYear.frx":0004
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQryOk 
         Height          =   510
         Left            =   7785
         TabIndex        =   3
         Top             =   90
         Width           =   1680
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         Picture         =   "frmPartYear.frx":08DE
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "조회년도:"
         Height          =   240
         Left            =   315
         TabIndex        =   2
         Top             =   180
         Width           =   915
      End
   End
   Begin FPSpreadADO.fpSpread ssYearRep 
      Height          =   6630
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   11760
      _Version        =   196608
      _ExtentX        =   20743
      _ExtentY        =   11695
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   15
      ScrollBars      =   2
      SpreadDesigner  =   "frmPartYear.frx":11B8
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmPartYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    Dim strFont0        As String
    Dim strFont1        As String
    Dim strFont2        As String
    Dim strHead1        As String
    Dim strHead2        As String
    Dim strHead3        As String
    Dim strHead4        As String
    Dim strHead5        As String
    Dim sPortBar        As String
    
    sPortBar = ""
    
    For i = 1 To 50
        sPortBar = sPortBar & "━"
    Next
    
    
    
    If ssYearRep.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 년간 통계 보고서"
    strHead3 = "/f3" & "기간 : " & cmbYear.Text
    strHead5 = "/f4" & ""
    
    ssYearRep.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                     strFont2 + strHead3 + _
                                     strFont2 + strHead5
    ssYearRep.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(80) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                                     
    ssYearRep.PrintMarginLeft = 500
    ssYearRep.PrintMarginRight = 0
    ssYearRep.PrintMarginTop = 500
    ssYearRep.PrintMarginBottom = 500
    ssYearRep.PrintColHeaders = True
    ssYearRep.PrintRowHeaders = True
    ssYearRep.PrintBorder = True
    ssYearRep.PrintColor = False
    ssYearRep.PrintGrid = True
    ssYearRep.PrintShadows = True
    ssYearRep.PrintUseDataMax = False
    
    ssYearRep.Row = 1: ssYearRep.Row2 = ssYearRep.DataRowCnt
    ssYearRep.Col = 2: ssYearRep.Col2 = ssYearRep.DataColCnt
    ssYearRep.PrintType = SS_PRINT_CELL_RANGE
    ssYearRep.PrintOrientation = PrintOrientationPortrait
    ssYearRep.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdQryOk_Click()
    Dim sMonth(1 To 12) As String
    Dim sCode           As String
    Dim nRowTotal       As Single
    Dim nColcnt(15)     As Single
    Dim j               As Integer
    Dim LnSLipno        As Integer
    
    
    sMonth(1) = Left(Trim$(cmbYear.Text), 4) & "-" & "01"
    sMonth(2) = Left(Trim$(cmbYear.Text), 4) & "-" & "02"
    sMonth(3) = Left(Trim$(cmbYear.Text), 4) & "-" & "03"
    sMonth(4) = Left(Trim$(cmbYear.Text), 4) & "-" & "04"
    sMonth(5) = Left(Trim$(cmbYear.Text), 4) & "-" & "05"
    sMonth(6) = Left(Trim$(cmbYear.Text), 4) & "-" & "06"
    sMonth(7) = Left(Trim$(cmbYear.Text), 4) & "-" & "07"
    sMonth(8) = Left(Trim$(cmbYear.Text), 4) & "-" & "08"
    sMonth(9) = Left(Trim$(cmbYear.Text), 4) & "-" & "09"
    sMonth(10) = Left(Trim$(cmbYear.Text), 4) & "-" & "10"
    sMonth(11) = Left(Trim$(cmbYear.Text), 4) & "-" & "11"
    sMonth(12) = Left(Trim$(cmbYear.Text), 4) & "-" & "12"
    
    StrSql = ""
    StrSql = StrSql & "  SELECT SLipno1, Codenm, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(1) & "',   1, '0')) JAN, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(2) & "',   1, '0')) FEB, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(3) & "',   1, '0')) MAR, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(4) & "',   1, '0')) APR, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(5) & "',   1, '0')) MAY, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(6) & "',   1, '0')) JUN, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(7) & "',   1, '0')) JUL, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(8) & "',   1, '0')) AUG, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(9) & "',   1, '0')) SEP, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(10) & "',  1, '0')) OCT, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(11) & "',  1, '0')) NOV, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '" & sMonth(12) & "',  1, '0')) DEC, "
    StrSql = StrSql & "         SUM(DECODE(JeobsuMM,  '',               '', 1))      LineTotal"
    StrSql = StrSql & "  FROM ( SELECT TO_CHAR(a.COLLDate, 'yyyy-MM') JeobsuMM, "
    StrSql = StrSql & "                a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, a.SLipno1, b.Codenm"
    StrSql = StrSql & "         FROM   TWEXAM_Order   a,"
    StrSql = StrSql & "                TWEXAM_Specode b "
    StrSql = StrSql & "         WHERE  TO_CHAR(a.COLLDate,'yyyyMMdd') LIKE '" & Trim(cmbYear.Text) & "%'"
    StrSql = StrSql & "         AND    a.SLipno1  <  52"
    StrSql = StrSql & "         AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "         AND    a.SLipno1   = b.Codeky"
    StrSql = StrSql & "         AND    b.Codegu    = '12')"
    StrSql = StrSql & "  GROUP  BY SLipno1, Codenm"
    
    ssYearRep.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    ssYearRep.MaxRows = adoSet.RecordCount
    ssYearRep.RowHeight(-1) = 11.5
    
    Do Until adoSet.EOF
        ssYearRep.Row = ssYearRep.DataRowCnt + 1
        ssYearRep.Col = 1: ssYearRep.Text = adoSet.Fields("SLipno1").Value & ""
        ssYearRep.Col = 2: ssYearRep.Text = adoSet.Fields("Codenm").Value & ""
        ssYearRep.Col = 3:  ssYearRep.Text = adoSet.Fields("JAN").Value & ""
        ssYearRep.Col = 4:  ssYearRep.Text = adoSet.Fields("FEB").Value & ""
        ssYearRep.Col = 5:  ssYearRep.Text = adoSet.Fields("MAR").Value & ""
        ssYearRep.Col = 6:  ssYearRep.Text = adoSet.Fields("APR").Value & ""
        ssYearRep.Col = 7:  ssYearRep.Text = adoSet.Fields("MAY").Value & ""
        ssYearRep.Col = 8:  ssYearRep.Text = adoSet.Fields("JUN").Value & ""
        ssYearRep.Col = 9:  ssYearRep.Text = adoSet.Fields("JUL").Value & ""
        ssYearRep.Col = 10: ssYearRep.Text = adoSet.Fields("AUG").Value & ""
        ssYearRep.Col = 11: ssYearRep.Text = adoSet.Fields("SEP").Value & ""
        ssYearRep.Col = 12: ssYearRep.Text = adoSet.Fields("OCT").Value & ""
        ssYearRep.Col = 13: ssYearRep.Text = adoSet.Fields("NOV").Value & ""
        ssYearRep.Col = 14: ssYearRep.Text = adoSet.Fields("DEC").Value & ""
        ssYearRep.Col = 15: ssYearRep.Text = adoSet.Fields("LineTotal").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    For i = 3 To 15
        ssYearRep.Col = i
        nColcnt(i) = 0
        For j = 1 To ssYearRep.DataRowCnt
            ssYearRep.Row = j
            If ssYearRep.Text = "" Then ssYearRep.Text = 0
            nColcnt(i) = nColcnt(i) + CSng(ssYearRep.Text)
        Next
    Next

    ssYearRep.MaxRows = ssYearRep.MaxRows + 1
    ssYearRep.Row = ssYearRep.MaxRows
    
    
    For i = 3 To 15
        ssYearRep.Col = i
        ssYearRep.Row = ssYearRep.MaxRows
        ssYearRep.Text = nColcnt(i)
    Next
    
    ssYearRep.Row = ssYearRep.DataRowCnt
    ssYearRep.Col = 2
    ssYearRep.Text = "   총  합 계 "
    
    For i = 3 To 15
        ssYearRep.Col = i
        For j = 1 To ssYearRep.DataRowCnt
            ssYearRep.Row = j
            If Trim(ssYearRep.Text) = "0" Then ssYearRep.Text = ""
        Next
    Next
    
    ssYearRep.Row = ssYearRep.DataRowCnt
    ssYearRep.Row2 = ssYearRep.DataRowCnt
    ssYearRep.Col = 1: ssYearRep.Col2 = 15
    ssYearRep.BlockMode = True
    ssYearRep.CellBorderType = SS_BORDER_TYPE_TOP
    ssYearRep.CellBorderStyle = SS_BORDER_STYLE_SOLID
    ssYearRep.Action = SS_ACTION_SET_CELL_BORDER
    ssYearRep.BlockMode = False
    ssYearRep.ReDraw = True
    
    Exit Sub
    

End Sub

Private Sub Form_Load()
    
    For i = 1999 To 2010
        cmbYear.AddItem i
    Next

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
