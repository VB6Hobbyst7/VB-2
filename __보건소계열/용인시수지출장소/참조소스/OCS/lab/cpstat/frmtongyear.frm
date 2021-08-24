VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmTongYear 
   Caption         =   "년간 통계"
   ClientHeight    =   8130
   ClientLeft      =   195
   ClientTop       =   795
   ClientWidth     =   11805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   11805
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   135
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
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Text            =   "cmbSLip"
         Top             =   135
         Width           =   2265
      End
      Begin VB.ComboBox cmbYear 
         Height          =   300
         ItemData        =   "frmTongYear.frx":0000
         Left            =   1350
         List            =   "frmTongYear.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   135
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "검사종목"
         Height          =   195
         Left            =   3105
         TabIndex        =   7
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "조회년도:"
         Height          =   240
         Left            =   315
         TabIndex        =   5
         Top             =   180
         Width           =   915
      End
      Begin MSForms.CommandButton cmdQryOk 
         Height          =   510
         Left            =   7785
         TabIndex        =   4
         Top             =   90
         Width           =   1680
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         Picture         =   "frmTongYear.frx":0004
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   510
         Left            =   9495
         TabIndex        =   3
         Top             =   90
         Width           =   1680
         Caption         =   "출력"
         PicturePosition =   327683
         Size            =   "2963;900"
         Picture         =   "frmTongYear.frx":08DE
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread ssYearRep 
      Height          =   6630
      Left            =   45
      TabIndex        =   0
      Top             =   1035
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
      SpreadDesigner  =   "frmTongYear.frx":11B8
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmTongYear"
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
    
    Dim sMonth(1 To 12) As String * 2
    
    Screen.MousePointer = vbHourglass
    For i = 1 To 12
        sMonth(i) = Format(i, "00")
    Next
    
    GoSub Get_MainData
    GoSub Total_Calculate
    Screen.MousePointer = vbDefault
    Exit Sub
    
    
Get_MainData:
    StrSql = ""
    StrSql = StrSql & "        SELECT itemcd, itemname,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(1) & "',  1, '')) Jan,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(2) & "',  1, '')) Feb,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(3) & "',  1, '')) Mar,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(4) & "',  1, '')) Apr,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(5) & "',  1, '')) May,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(6) & "',  1, '')) Jun,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(7) & "',  1, '')) Jul,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(8) & "',  1, '')) Aug,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(9) & "',  1, '')) Sep,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(10) & "', 1, '')) Oct,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(11) & "', 1, '')) Nov,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '" & sMonth(12) & "', 1, '')) Dec,"
    StrSql = StrSql & "              SUM(Decode(JeobsuMM, '', '', 1)) LineTotal"
    StrSql = StrSql & "       FROM(  SELECT DISTINCT a.COLLDate,  a.COLLHH, a.COLLMM, a.Itemcd, b.Routinnm ItemName,"
    StrSql = StrSql & "                     TO_CHAR(a.COLLDate, 'MM') JeobsuMM"
    StrSql = StrSql & "              FROM   TWEXAM_Order   a,"
    StrSql = StrSql & "                     TWEXAM_Routine b"
    StrSql = StrSql & "              WHERE  TO_CHAR(a.COLLDate,'yyyyMMdd') LIKE '" & cmbYear.Text & "%'"
    StrSql = StrSql & "              AND    a.JeobsuYN  =  '*'"
    If cmbSLip.ListIndex > -1 Then
        StrSql = StrSql & "          AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    Else
        StrSql = StrSql & "          AND  ( a.SLipno1 BETWEEN 11 and 51 )"
    End If
    StrSql = StrSql & "              AND    a.ItemCd = b.RoutinCD"
    StrSql = StrSql & "                 Union ALL"
    StrSql = StrSql & "              SELECT a.COLLDate,  a.COLLHH, a.COLLMM, a.Itemcd, b.ItemNm ItemName,"
    StrSql = StrSql & "                     TO_CHAR(a.COLLDate, 'MM') JeobsuMM"
    StrSql = StrSql & "              FROM   TWEXAM_Order  a,"
    StrSql = StrSql & "                     TWEXAM_ItemML b"
    StrSql = StrSql & "              WHERE  TO_CHAR(a.COLLDate,'yyyyMMdd') LIKE '" & cmbYear.Text & "%'"
    StrSql = StrSql & "              AND    a.JeobsuYN  =  '*'"
    If cmbSLip.ListIndex > -1 Then
        StrSql = StrSql & "          AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    Else
        StrSql = StrSql & "          AND  ( a.SLipno1 BETWEEN 11 and 51 )"
    End If
    StrSql = StrSql & "              AND    a.ItemCd    = b.Codeky)"
    StrSql = StrSql & "       GROUP BY ITEMCD, ITEMNAME"
    
    ssYearRep.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    ssYearRep.MaxRows = adoSet.RecordCount
    ssYearRep.RowHeight(-1) = 11.5
    
    Do Until adoSet.EOF
        ssYearRep.Row = ssYearRep.DataRowCnt + 1
        ssYearRep.Col = 1:  ssYearRep.Text = adoSet.Fields("ItemCd").Value & ""
        ssYearRep.Col = 2:  ssYearRep.Text = adoSet.Fields("ItemName").Value & ""
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
    Return
    

Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    If ssYearRep.DataRowCnt = ssYearRep.MaxRows Then
        ssYearRep.MaxRows = ssYearRep.MaxRows + 2
    End If
    
    
    iLastSprRow = ssYearRep.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To ssYearRep.MaxCols
        nCalSum = 0
        For i = 1 To iLastSprRow
            ssYearRep.Row = i
            ssYearRep.Col = j
            If Trim(ssYearRep.Text) <> "" Then
               nCalSum = nCalSum + CSng(ssYearRep.Text)
            End If
        Next
        ssYearRep.Row = iLastSprRow + 2
        ssYearRep.Col = j
        If nCalSum > 0 Then
            ssYearRep.Text = nCalSum
        End If
    Next
    
    ssYearRep.Row = ssYearRep.DataRowCnt
    ssYearRep.Col = 2: ssYearRep.Text = "합계"
    Return

End Sub

Private Sub CommandButton1_Click()



End Sub

Private Sub Form_Load()
    
    For i = 1999 To 2010
        cmbYear.AddItem i
    Next
    
    GoSub Get_SLip
    Exit Sub
    
Get_SLip:
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Specode"
    StrSql = StrSql & " WHERE  CODEGU = '12'"
    StrSql = StrSql & " AND    Codeky < '52'"
    StrSql = StrSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    cmbSLip.Clear
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & _
                             adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
        
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
