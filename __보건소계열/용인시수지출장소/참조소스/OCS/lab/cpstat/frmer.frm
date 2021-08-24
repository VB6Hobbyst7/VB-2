VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmEr 
   Caption         =   "응급검사 통계"
   ClientHeight    =   7170
   ClientLeft      =   240
   ClientTop       =   2490
   ClientWidth     =   11490
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
   ScaleHeight     =   7170
   ScaleWidth      =   11490
   WindowState     =   2  '최대화
   Begin VB.ComboBox cmbYear1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   -17354
      TabIndex        =   11
      Top             =   -15884
      Width           =   1095
   End
   Begin VB.ComboBox cmbMonth 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmEr.frx":0000
      Left            =   -18164
      List            =   "frmEr.frx":0028
      Style           =   2  '드롭다운 목록
      TabIndex        =   10
      Top             =   -15884
      Width           =   825
   End
   Begin VB.TextBox txtLastDate 
      Appearance      =   0  '평면
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   300
      Left            =   -18614
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -15884
      Width           =   420
   End
   Begin VB.ComboBox cmbSLip1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   -22124
      TabIndex        =   8
      Text            =   "cmbSLip"
      Top             =   -15884
      Width           =   2265
   End
   Begin VB.ComboBox cmbYear 
      Height          =   300
      ItemData        =   "frmEr.frx":005C
      Left            =   1080
      List            =   "frmEr.frx":005E
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   675
      Width           =   1230
   End
   Begin VB.ComboBox cmbSLip 
      Height          =   300
      Left            =   3330
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   675
      Width           =   2265
   End
   Begin FPSpreadADO.fpSpread sprYear 
      Height          =   6225
      Left            =   180
      TabIndex        =   5
      Top             =   1260
      Width           =   11130
      _Version        =   196608
      _ExtentX        =   19632
      _ExtentY        =   10980
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
      MaxCols         =   15
      ScrollBars      =   2
      SpreadDesigner  =   "frmEr.frx":0060
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread sprMonth 
      Height          =   6315
      Left            =   -26174
      TabIndex        =   15
      Top             =   -22439
      Width           =   10950
      _Version        =   196608
      _ExtentX        =   19315
      _ExtentY        =   11139
      _StockProps     =   64
      Enabled         =   0   'False
      BackColorStyle  =   1
      ColsFrozen      =   2
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
      MaxCols         =   35
      SpreadDesigner  =   "frmEr.frx":3EF2
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   510
      Left            =   -25949
      TabIndex        =   0
      Top             =   -16049
      Width           =   1545
      VariousPropertyBits=   25
      Caption         =   "출력"
      PicturePosition =   327683
      Size            =   "2725;900"
      Picture         =   "frmEr.frx":5C9A
      FontName        =   "굴림체"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label4 
      Caption         =   "조회년,월"
      Enabled         =   0   'False
      Height          =   240
      Left            =   -16184
      TabIndex        =   14
      Top             =   -15869
      Width           =   870
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   510
      Left            =   -24419
      TabIndex        =   13
      Top             =   -16049
      Width           =   1635
      VariousPropertyBits=   25
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2884;900"
      FontName        =   "굴림체"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      Caption         =   "검사종목"
      Enabled         =   0   'False
      Height          =   195
      Left            =   -19784
      TabIndex        =   12
      Top             =   -15824
      Width           =   735
   End
   Begin MSForms.CommandButton cmdPr0 
      Height          =   465
      Left            =   7740
      TabIndex        =   7
      Top             =   675
      Width           =   1500
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "2646;820"
      Picture         =   "frmEr.frx":6574
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery0 
      Height          =   465
      Left            =   6255
      TabIndex        =   6
      Top             =   675
      Width           =   1500
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2646;820"
      Picture         =   "frmEr.frx":6E4E
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "조회년도:"
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "검사종목"
      Height          =   195
      Left            =   2475
      TabIndex        =   3
      Top             =   720
      Width           =   825
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmEr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbMonth_Click()
    Dim sYM         As String
    Dim sDate       As String
    
    sDate = cmbYear1.Text & "-" & cmbMonth.Text & "-" & "01"
    
    GoSub Get_LastDate
    Exit Sub
    
    



Get_LastDate:
    strSql = " SELECT TO_CHAR(LAST_DAY(TO_DATE('" & sDate & "','yyyy-MM-dd')),'dd') LstDay FROM DUAL"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    txtLastDate.Text = adoSet.Fields("LstDay").Value & ""
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub cmdPr0_Click()
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
    
    
    If sprYear.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 응급Order(년보) 통계"
    strHead3 = "/f3" & "조회년도 : " & cmbYear.Text
    strHead5 = "/f4" & "    " & cmbSLip.Text
    
    sprYear.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                     strFont2 + "/l" + strHead3 + _
                                     strFont2 + strHead5
    sprYear.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(80) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprYear.PrintMarginLeft = 0
    sprYear.PrintMarginRight = 0
    sprYear.PrintMarginTop = 0
    sprYear.PrintMarginBottom = 0
    sprYear.PrintColHeaders = True
    sprYear.PrintRowHeaders = False
    sprYear.PrintBorder = True
    sprYear.PrintColor = True
    sprYear.PrintGrid = True
    sprYear.PrintShadows = True
    sprYear.PrintUseDataMax = False
    
    sprYear.Row = 1: sprYear.Row2 = sprYear.DataRowCnt
    sprYear.Col = 2: sprYear.Col2 = sprYear.DataColCnt
    sprYear.PrintType = SS_PRINT_CELL_RANGE
    sprYear.PrintOrientation = PrintOrientationPortrait
    sprYear.Action = SS_ACTION_PRINT


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
    strHead2 = "/f2" & "/c" & "임상병리과 응급Order(일자별) 통계"
    strHead3 = "/f3" & "조회년월 : " & cmbYear1.Text & "년 " & cmbMonth.Text & "월 "
    strHead5 = "/f4" & "    " & cmbSLip1.Text
    
    sprMonth.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                     strFont2 + "/l" + strHead3 + _
                                     strFont2 + strHead5
    sprMonth.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(120) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprMonth.PrintMarginLeft = 0
    sprMonth.PrintMarginRight = 0
    sprMonth.PrintMarginTop = 0
    sprMonth.PrintMarginBottom = 0
    sprMonth.PrintColHeaders = True
    sprMonth.PrintRowHeaders = False
    sprMonth.PrintBorder = True
    sprMonth.PrintColor = True
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
    Dim sYYYYMM         As String
    
    
    
    sYYYYMM = Trim(cmbYear1.Text) & Trim(cmbMonth.Text)
    
    If cmbSLip1.ListIndex = -1 Or Trim(cmbSLip1.Text) = "" Then
        LnSLipno = 0
    Else
        LnSLipno = Val(Left(cmbSLip1.Text, 2))
    End If
    

    If cmbYear1.ListIndex = -1 Or Trim(cmbYear1.Text) = "" Then
        MsgBox "조회할 년도지정을 하십시오", vbInformation
        Exit Sub
    End If
    
    If cmbMonth.ListIndex = -1 Or Trim(cmbMonth.Text) = "" Then
        MsgBox "조회할 월 지정을 하십시오", vbInformation
        Exit Sub
    End If
    
    
    
    Call SpreadSetClear(sprMonth)
    
    GoSub Spread_Cell_Set
    GoSub Get_Data_Process
    
    GoSub Data_Spread_Calcurate
    
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
        
    strSql = ""
    strSql = strSql & " SELECT ItemCd, ItemName,"
    
    For i = 1 To 31
        sCompDay = Format(i, "00")
        strSql = strSql & "    SUM(DeCode(Day, '" & sCompDay & "',1,'')) " & "D" & sCompDay & ","
    Next
    strSql = strSql & "        SUM(DECODE(Day, '', '', 1)) Tot"
    
    strSql = strSql & " FROM(  SELECT TO_CHAR(a.COLLDate, 'dd') Day,"
    strSql = strSql & "               a.ItemCd, b.ItemNM ItemName,"
    strSql = strSql & "               TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "               a.JeobsuT1, a.JeobsuT2"
    strSql = strSql & "        FROM   TWEXAM_ORDER  a,"
    strSql = strSql & "               TWEXAM_ITEMML b"
    strSql = strSql & "        WHERE  TO_CHAR(a.COLLDate, 'YYYYMM') = '" & sYYYYMM & "'"
    If LnSLipno > 0 Then
        strSql = strSql & "    AND    a.SLipno1   = " & LnSLipno & ""
    End If
    strSql = strSql & "        AND   (a.DEPTCODE = 'ER' OR a.GBER = 'E')"
    strSql = strSql & "        AND    a.JEOBSUYN = '*'"
    strSql = strSql & "        AND    a.ItemCd   = Codeky"
    
    strSql = strSql & "        UNION ALL"
    strSql = strSql & "        SELECT DISTINCT TO_CHAR(a.COLLDate, 'dd') Day,"
    strSql = strSql & "               a.ItemCd, b.RoutinNm ItemName,"
    strSql = strSql & "               TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "               a.JeobsuT1, a.JeobsuT2"
    strSql = strSql & "        FROM   TWEXAM_ORDER   a,"
    strSql = strSql & "               TWEXAM_ROUTINE b"
    strSql = strSql & "        WHERE  TO_CHAR(a.COLLDate, 'YYYYMM') = '" & sYYYYMM & "'"
    If LnSLipno > 0 Then
        strSql = strSql & "    AND    a.SLipno1   = " & LnSLipno & ""
    End If
    strSql = strSql & "        AND   (a.DEPTCODE = 'ER' OR a.GBER = 'E')           "
    strSql = strSql & "        AND    a.JEOBSUYN = '*'"
    strSql = strSql & "        AND    a.ItemCd   = RoutinCd)"
    strSql = strSql & " GROUP  BY ITEMCD, ITEMNAME      "

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprMonth.Row = sprMonth.DataRowCnt + 1
        sprMonth.Col = 1: sprMonth.Text = adoSet.Fields("ItemCd").Value & ""
        sprMonth.Col = 2: sprMonth.Text = adoSet.Fields("ItemName").Value & ""
        
        For j = 3 To sprMonth.MaxCols - 1
            sprMonth.Col = j: sprMonth.Text = adoSet.Fields("D" & Format(j - 2, "00")).Value & ""
            If sprMonth.Text = "0" Then sprMonth.Text = ""
        Next
        sprMonth.Col = sprMonth.MaxCols: sprMonth.Text = adoSet.Fields("Tot").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
Data_Spread_Calcurate:
    Dim nLastLine       As Integer
    Dim nTmpSum         As Integer
    
    nLastLine = sprMonth.DataRowCnt
    
    For i = 3 To sprMonth.MaxCols
        sprMonth.Col = i
        For j = 1 To nLastLine
            sprMonth.Row = j
            nTmpSum = nTmpSum + Val(sprMonth.Text)
        Next
        sprMonth.Row = nLastLine + 2
        sprMonth.Col = 2: sprMonth.Text = "총합계"
        sprMonth.Col = i
        If nTmpSum > 0 Then
            sprMonth.Text = nTmpSum
        End If
        nTmpSum = 0
    Next
    
    
    Return

End Sub

Private Sub cmdQuery0_Click()
    Dim sSLip       As String * 2
    
    If cmbYear.ListIndex = -1 Then
        MsgBox "조회할 년도를 선택하세요!..", vbCritical
        Exit Sub
    End If
    
    
    sSLip = ""
    If cmbSLip.ListIndex > -1 Then
        sSLip = Left(cmbSLip, 2)
    End If
    
    
    Call SpreadSetClear(sprYear)
    GoSub Get_MainProc
    GoSub Total_Calculate
    
    Exit Sub


Get_MainProc:
    strSql = ""
    strSql = strSql & " SELECT ItemCd, ItemName,"
    strSql = strSql & "        SUM(Decode(Month, '01', 1, '')) Jan,"
    strSql = strSql & "        SUM(Decode(Month, '02', 1, '')) Feb,"
    strSql = strSql & "        SUM(Decode(Month, '03', 1, '')) Mar,"
    strSql = strSql & "        SUM(Decode(Month, '04', 1, '')) Apr,"
    strSql = strSql & "        SUM(Decode(Month, '05', 1, '')) May,"
    strSql = strSql & "        SUM(Decode(Month, '06', 1, '')) Jun,"
    strSql = strSql & "        SUM(Decode(Month, '07', 1, '')) Jul,"
    strSql = strSql & "        SUM(Decode(Month, '08', 1, '')) Aug,"
    strSql = strSql & "        SUM(Decode(Month, '09', 1, '')) Sep,"
    strSql = strSql & "        SUM(Decode(Month, '10', 1, '')) Oct,"
    strSql = strSql & "        SUM(Decode(Month, '11', 1, '')) Nov,"
    strSql = strSql & "        SUM(Decode(Month, '12', 1, '')) Dec,"
    strSql = strSql & "        SUM(Decode(Month, '00', '' , 1)) LineTotal"
    strSql = strSql & " FROM(  SELECT TO_CHAR(a.COLLDate, 'MM') Month,"
    strSql = strSql & "               a.ItemCd, b.ItemNM ItemName,"
    strSql = strSql & "               TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "               a.JeobsuT1, a.JeobsuT2"
    strSql = strSql & "        FROM   TWEXAM_ORDER  a,"
    strSql = strSql & "               TWEXAM_ITEMML b"
    strSql = strSql & "        WHERE  TO_CHAR(a.COLLDate, 'YYYY') = '" & Left(cmbYear.Text, 4) & "'"
    If Trim(sSLip) <> "" Then
        strSql = strSql & "    AND    a.SLipno1    = " & Val(sSLip)
    End If
    strSql = strSql & "        AND   (a.DEPTCODE = 'ER' OR a.GBER = 'E')"
    strSql = strSql & "        AND    a.JEOBSUYN = '*'"
    strSql = strSql & "        AND    a.ItemCd   = Codeky"
    strSql = strSql & "        UNION ALL"
    strSql = strSql & "        SELECT DISTINCT TO_CHAR(a.COLLDate, 'MM') Month,"
    strSql = strSql & "               a.ItemCd, b.RoutinNm ItemName,"
    strSql = strSql & "               TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "               a.JeobsuT1, a.JeobsuT2"
    strSql = strSql & "        FROM   TWEXAM_ORDER   a,"
    strSql = strSql & "               TWEXAM_ROUTINE b"
    strSql = strSql & "        WHERE  TO_CHAR(a.COLLDate, 'YYYY') = '" & Left(cmbYear.Text, 4) & "'"
    If Trim(sSLip) <> "" Then
        strSql = strSql & "    AND    a.SLipno1    = " & Val(sSLip)
    End If
    strSql = strSql & "        AND   (a.DEPTCODE = 'ER' OR a.GBER = 'E')           "
    strSql = strSql & "        AND    a.JEOBSUYN = '*'"
    strSql = strSql & "        AND    a.ItemCd   = RoutinCd)"
    strSql = strSql & " GROUP BY ITEMCD, ITEMNAME"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprYear.Row = sprYear.DataRowCnt + 1
        sprYear.Col = 1:  sprYear.Text = adoSet.Fields("ItemCd").Value & ""
        sprYear.Col = 2:  sprYear.Text = adoSet.Fields("ItemName").Value & ""
        sprYear.Col = 3:  sprYear.Text = adoSet.Fields("Jan").Value & ""
        sprYear.Col = 4:  sprYear.Text = adoSet.Fields("Feb").Value & ""
        sprYear.Col = 5:  sprYear.Text = adoSet.Fields("Mar").Value & ""
        sprYear.Col = 6:  sprYear.Text = adoSet.Fields("Apr").Value & ""
        sprYear.Col = 7:  sprYear.Text = adoSet.Fields("May").Value & ""
        sprYear.Col = 8:  sprYear.Text = adoSet.Fields("Jun").Value & ""
        sprYear.Col = 9:  sprYear.Text = adoSet.Fields("Jul").Value & ""
        sprYear.Col = 10: sprYear.Text = adoSet.Fields("Aug").Value & ""
        sprYear.Col = 11: sprYear.Text = adoSet.Fields("Sep").Value & ""
        sprYear.Col = 12: sprYear.Text = adoSet.Fields("Oct").Value & ""
        sprYear.Col = 13: sprYear.Text = adoSet.Fields("Nov").Value & ""
        sprYear.Col = 14: sprYear.Text = adoSet.Fields("Dec").Value & ""
        sprYear.Col = 15: sprYear.Text = adoSet.Fields("Linetotal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    If sprYear.DataRowCnt = sprYear.MaxRows Then
        sprYear.MaxRows = sprYear.MaxRows + 2
    End If
    
    
    iLastSprRow = sprYear.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprYear.MaxCols
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprYear.Row = i
            sprYear.Col = j
            If Trim(sprYear.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprYear.Text)
            End If
        Next
        sprYear.Row = iLastSprRow + 2
        sprYear.Col = j
        If nCalSum > 0 Then
            sprYear.Text = nCalSum
        End If
    Next
    
    sprYear.Row = sprYear.DataRowCnt
    sprYear.Col = 2: sprYear.Text = "합계"
    Return
    
End Sub

Private Sub Form_Load()
    
    
    For i = 1999 To 2010
        cmbYear.AddItem i
        cmbYear1.AddItem i
    Next
    
    GoSub Get_SLip
    Exit Sub
    
Get_SLip:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '12'"
    strSql = strSql & " AND    Codeky < '52'"
    strSql = strSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    cmbSLip.Clear
    cmbSLip1.Clear
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                             adoSet.Fields("Codenm").Value & ""
        cmbSLip1.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                              adoSet.Fields("Codenm").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
