VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPartMonth 
   Caption         =   "검사통계(월별)"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   780
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11625
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprMonth 
      Height          =   6405
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   11490
      _Version        =   196608
      _ExtentX        =   20267
      _ExtentY        =   11298
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   35
      ScrollBars      =   2
      SpreadDesigner  =   "frmPartMonth.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   960
      Left            =   45
      TabIndex        =   1
      Top             =   180
      Width           =   11490
      _Version        =   65536
      _ExtentX        =   20267
      _ExtentY        =   1693
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
         Left            =   1395
         TabIndex        =   4
         Top             =   225
         Width           =   1095
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   300
         ItemData        =   "frmPartMonth.frx":1DA1
         Left            =   2475
         List            =   "frmPartMonth.frx":1DC9
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txtLastDate 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   420
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   510
         Left            =   9540
         TabIndex        =   7
         Top             =   180
         Width           =   1545
         Caption         =   "출력"
         PicturePosition =   327683
         Size            =   "2725;900"
         Picture         =   "frmPartMonth.frx":1DFD
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "조회년,월"
         Height          =   240
         Left            =   450
         TabIndex        =   6
         Top             =   270
         Width           =   870
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   510
         Left            =   7920
         TabIndex        =   5
         Top             =   180
         Width           =   1635
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2884;900"
         Picture         =   "frmPartMonth.frx":26D7
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmPartMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbMonth_Click()
    Dim sYM         As String
    Dim sDate       As String
    
    sDate = cmbYear.Text & "-" & cmbMonth.Text & "-" & "01"
    
    GoSub Get_LastDate
    Exit Sub
    

Get_LastDate:
    StrSql = " SELECT TO_CHAR(LAST_DAY(TO_DATE('" & sDate & "','yyyy-MM-dd')),'dd') LstDay FROM DUAL"
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    txtLastDate.Text = adoSet.Fields("LstDay").Value & ""
    Call adoSetClose(adoSet)
    
    Return

End Sub

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
    
    For i = 1 To 80
        sPortBar = sPortBar & "━"
    Next
    
    
    If sprMonth.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 월별 통계 보고서"
    strHead3 = "/f3" & "기간 : " & cmbYear.Text & " 년 " & cmbMonth.Text & " 월"
    strHead5 = "/f4" & ""
    
    sprMonth.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                      strFont2 + strHead3 + _
                                      strFont2 + strHead5
    sprMonth.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(120) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                                      
    sprMonth.PrintMarginLeft = 500
    sprMonth.PrintMarginRight = 0
    sprMonth.PrintMarginTop = 500
    sprMonth.PrintMarginBottom = 500
    sprMonth.PrintColHeaders = True
    sprMonth.PrintRowHeaders = True
    sprMonth.PrintBorder = True
    sprMonth.PrintColor = False
    sprMonth.PrintGrid = True
    sprMonth.PrintShadows = True
    sprMonth.PrintUseDataMax = False
    
    sprMonth.Row = 1: sprMonth.Row2 = sprMonth.DataRowCnt
    sprMonth.Col = 2: sprMonth.Col2 = sprMonth.DataColCnt
    sprMonth.PrintType = SS_PRINT_CELL_RANGE
    sprMonth.PrintOrientation = PrintOrientationLandscape
    sprMonth.Action = SS_ACTION_PRINT


End Sub

Private Sub cmdQuery_Click()
    Dim sCompDay        As String
    Dim j               As Integer
    Dim sCode           As String
    Dim LnSLipno        As Integer
    Dim sCompYM         As String
    

    If cmbYear.ListIndex = -1 Or Trim(cmbYear.Text) = "" Then
        MsgBox "조회할 년도지정을 하십시오", vbInformation
        Exit Sub
    End If
    
    If cmbMonth.ListIndex = -1 Or Trim(cmbMonth.Text) = "" Then
        MsgBox "조회할 월 지정을 하십시오", vbInformation
        Exit Sub
    End If
    
    
    sCompYM = Left(cmbYear.Text, 4) & cmbMonth.Text
    
    Call SpreadSetClear(sprMonth)
    
    GoSub Spread_Cell_Set
    GoSub Get_Data_Process
    GoSub Total_Calculate
    
    Exit Sub
    


Spread_Cell_Set:
    sprMonth.MaxCols = Val(txtLastDate.Text) + 3
    sprMonth.ColWidth(1) = 6.75
    sprMonth.ColWidth(2) = 20
    
    For i = 3 To sprMonth.MaxCols - 1
        sprMonth.ColWidth(i) = 2.7
        sprMonth.Row = 0
        sprMonth.Col = i
        sprMonth.Text = i - 2
    Next
    
    sprMonth.ColWidth(sprMonth.MaxCols) = 4.5
    sprMonth.Row = 0
    sprMonth.Col = sprMonth.MaxCols
    sprMonth.Text = "Total"
    Return
    
    
Get_Data_Process:
    StrSql = ""
    StrSql = StrSql & " SELECT SLipno1, Codenm,"
    
    For i = 1 To 31
        sCompDay = Format(i, "00")
        StrSql = StrSql & "    SUM(DeCode(SubStr(TO_Char(COLLDate,'yyyyMMdd'),7,2), '" & sCompDay & "',1,'0')) " & "D" & sCompDay & ","
    Next
    StrSql = StrSql & "        SUM(DECODE(SUBSTR(TO_Char(COLLDate,'yyyyMMdd'),7,2), '',  '0',1))   LineTotal"
    StrSql = StrSql & "  FROM ( SELECT a.COLLDate, a.COLLHH, a.COLLMM, a.SLipno1, b.Codenm"
    StrSql = StrSql & "         FROM   TWEXAM_Order   a,"
    StrSql = StrSql & "                TWEXAM_Specode b "
    StrSql = StrSql & "         WHERE  TO_CHAR(a.COLLDate,'yyyyMMdd') LIKE '" & Trim(sCompYM) & "%'"
    StrSql = StrSql & "         AND    a.SLipno1  <  52"
    StrSql = StrSql & "         AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "         AND    a.SLipno1   = b.Codeky"
    StrSql = StrSql & "         AND    b.Codegu    = '12')"
    StrSql = StrSql & "  GROUP  BY SLipno1, Codenm"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprMonth.Row = sprMonth.DataRowCnt + 1
        sprMonth.Col = 1: sprMonth.Text = adoSet.Fields("SLipno1").Value & ""
        sprMonth.Col = 2: sprMonth.Text = adoSet.Fields("Codenm").Value & ""
        
        For j = 3 To sprMonth.MaxCols - 1
            sprMonth.Col = j: sprMonth.Text = adoSet.Fields("D" & Format(j - 2, "00")).Value & ""
            If sprMonth.Text = "0" Then sprMonth.Text = ""
        Next
        sprMonth.Col = sprMonth.MaxCols: sprMonth.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Return
    
    
Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    iLastSprRow = sprMonth.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprMonth.MaxCols
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprMonth.Row = i
            sprMonth.Col = j
            If Trim(sprMonth.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprMonth.Text)
            End If
        Next
        sprMonth.Row = iLastSprRow + 2
        sprMonth.Col = j
        sprMonth.Text = nCalSum
    Next
    
    sprMonth.Row = sprMonth.DataRowCnt
    sprMonth.Col = 2: sprMonth.Text = "합계"
    
    
    For i = 1 To sprMonth.DataRowCnt
        For j = 3 To sprMonth.DataColCnt
            sprMonth.Row = i
            sprMonth.Col = j
            If sprMonth.Text = "0" Then
                sprMonth.Text = ""
            End If
        Next
    Next
    Return

End Sub

Private Sub Form_Load()
    
    For i = 1999 To 2010
        cmbYear.AddItem i
    Next

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
    
End Sub
