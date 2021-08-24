VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmWeek 
   Caption         =   "요일별 통계"
   ClientHeight    =   7560
   ClientLeft      =   255
   ClientTop       =   1020
   ClientWidth     =   11535
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
   ScaleHeight     =   7560
   ScaleWidth      =   11535
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprWeek 
      Height          =   6765
      Left            =   315
      TabIndex        =   4
      Top             =   945
      Width           =   10950
      _Version        =   196608
      _ExtentX        =   19315
      _ExtentY        =   11933
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
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmWeek.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   315
      TabIndex        =   0
      Top             =   135
      Width           =   7305
      _Version        =   65536
      _ExtentX        =   12885
      _ExtentY        =   1138
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
         Left            =   4905
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   180
         Width           =   2310
      End
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   2520
         TabIndex        =   1
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36446
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36446
      End
      Begin VB.Label Label2 
         Caption         =   "검사종목"
         Height          =   240
         Left            =   4095
         TabIndex        =   6
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Width           =   780
      End
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   510
      Left            =   9405
      TabIndex        =   8
      Top             =   270
      Width           =   1590
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "2805;900"
      Picture         =   "frmWeek.frx":3DEE
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   510
      Left            =   7650
      TabIndex        =   7
      Top             =   270
      Width           =   1770
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "3122;900"
      Picture         =   "frmWeek.frx":46C8
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmWeek"
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
    
    
    If sprWeek.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "/c" & "임상병리과 요일별 검사건수 통계"
    strHead3 = "/f3" & "조회일자 : " & " (From: " & Format(dtFrDate.Value, "yyyy-MM-dd") & ") ~  (To: " & _
                                                    Format(dtToDate.Value, "yyyy-MM-dd") & ")"
    strHead5 = "/f4" & "    " & cmbSLip.Text
    
    sprWeek.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                     strFont2 + "/l" + strHead3 + _
                                     strFont2 + strHead5
    sprWeek.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(80) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprWeek.PrintMarginLeft = 0
    sprWeek.PrintMarginRight = 0
    sprWeek.PrintMarginTop = 0
    sprWeek.PrintMarginBottom = 0
    sprWeek.PrintColHeaders = True
    sprWeek.PrintRowHeaders = False
    sprWeek.PrintBorder = True
    sprWeek.PrintColor = False
    sprWeek.PrintGrid = True
    sprWeek.PrintShadows = True
    sprWeek.PrintUseDataMax = False
    
    sprWeek.Row = 1: sprWeek.Row2 = sprWeek.DataRowCnt
    sprWeek.Col = 2: sprWeek.Col2 = sprWeek.DataColCnt
    sprWeek.PrintType = SS_PRINT_CELL_RANGE
    sprWeek.PrintOrientation = PrintOrientationPortrait
    sprWeek.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    
    
    Call SpreadSetClear(sprWeek)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    GoSub Get_MainData
    GoSub Total_Calculate
    Exit Sub
    
    
Get_MainData:
    
    StrSql = ""
    StrSql = StrSql & " SELECT ItemCd, ItemName,"
    StrSql = StrSql & "        sum(Decode(Day, 'MONDAY',   1, ''))  MONDAY,"
    StrSql = StrSql & "        sum(Decode(Day, 'TUESDAY',  1, ''))  TUESDAY,"
    StrSql = StrSql & "        sum(Decode(Day, 'WEDNESDAY',1, ''))  WEDNESDAY,"
    StrSql = StrSql & "        sum(Decode(Day, 'THURSDAY', 1, ''))  THURSDAY,"
    StrSql = StrSql & "        sum(Decode(Day, 'FRIDAY',   1, ''))  FRIDAY,"
    StrSql = StrSql & "        sum(Decode(Day, 'SATURDAY', 1, ''))  SATURDAY,"
    StrSql = StrSql & "        sum(Decode(Day, 'SUNDAY',   1, ''))  SUNDAY,"
    StrSql = StrSql & "        sum(Decode(Day, '',        '',  1))  LineTotal"
    StrSql = StrSql & " FROM(  SELECT DISTINCT RTRIM(TO_CHAR(a.COLLDate, 'DAY')) Day,"
    StrSql = StrSql & "               TO_CHAR(a.COLLDate, 'yyyy-MM-dd') JeobsuDt,"
    StrSql = StrSql & "               a.COLLHH, a.COLLMM, a.SLipno1, a.ItemCd, b.Routinnm ItemName"
    StrSql = StrSql & "        FROM   TWEXAM_ORDER   a,"
    StrSql = StrSql & "               TWEXAM_Routine b"
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    If cmbSLip.ListIndex > -1 Then
        StrSql = StrSql & "    AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    Else
        StrSql = StrSql & "    AND    a.SLipno1   > 0 "
        StrSql = StrSql & "    AND    a.SLipno1   < 52"
    End If
    StrSql = StrSql & "        AND    a.itemcd    = b.Routincd"
    
    StrSql = StrSql & "   UNION ALL"
    StrSql = StrSql & "        SELECT RTRIM(TO_CHAR(a.COLLDate, 'DAY')) Day,"
    StrSql = StrSql & "               TO_CHAR(a.COLLDate, 'yyyy-MM-dd') JeobsuDt,"
    StrSql = StrSql & "               a.COLLHH, a.COLLMM, a.SLipno1, a.ItemCd, b.ItemNm ItemName"
    StrSql = StrSql & "        FROM   TWEXAM_ORDER  a,"
    StrSql = StrSql & "               TWEXAM_ITEMML b"
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    If cmbSLip.ListIndex > -1 Then
        StrSql = StrSql & "    AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    Else
        StrSql = StrSql & "    AND    a.SLipno1   > 0 "
        StrSql = StrSql & "    AND    a.SLipno1   < 52"
    End If
    StrSql = StrSql & "        AND    a.itemcd    = b.Codeky)"
    StrSql = StrSql & " GROUP  BY ITEMCD, ITEMNAME"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprWeek.Row = sprWeek.DataRowCnt + 1
        sprWeek.Col = 1:  sprWeek.Text = adoSet.Fields("ItemCd").Value & ""
        sprWeek.Col = 2:  sprWeek.Text = adoSet.Fields("ItemName").Value & ""
        sprWeek.Col = 3:  sprWeek.Text = adoSet.Fields("MONDAY").Value & ""
        sprWeek.Col = 4:  sprWeek.Text = adoSet.Fields("TUESDAY").Value & ""
        sprWeek.Col = 5:  sprWeek.Text = adoSet.Fields("WEDNESDAY").Value & ""
        sprWeek.Col = 6:  sprWeek.Text = adoSet.Fields("THURSDAY").Value & ""
        sprWeek.Col = 7:  sprWeek.Text = adoSet.Fields("FRIDAY").Value & ""
        sprWeek.Col = 8:  sprWeek.Text = adoSet.Fields("SATURDAY").Value & ""
        sprWeek.Col = 9:  sprWeek.Text = adoSet.Fields("SUNDAY").Value & ""
        sprWeek.Col = 10: sprWeek.Text = adoSet.Fields("LineTotal").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    

Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    iLastSprRow = sprWeek.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprWeek.MaxCols
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprWeek.Row = i
            sprWeek.Col = j
            If Trim(sprWeek.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprWeek.Text)
            End If
        Next
        sprWeek.Row = iLastSprRow + 2
        sprWeek.Col = j
        sprWeek.Text = nCalSum
    Next
    
    sprWeek.Row = sprWeek.DataRowCnt
    sprWeek.Col = 2: sprWeek.Text = "합계"
    Return
    
    
    
End Sub

Private Sub Form_Load()
    
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
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
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                             adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
