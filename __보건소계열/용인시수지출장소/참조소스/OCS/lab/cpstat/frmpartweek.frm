VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPartWeek 
   Caption         =   "검사통계(주간별)"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   945
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11640
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprWeek 
      Height          =   6180
      Left            =   135
      TabIndex        =   2
      Top             =   1125
      Width           =   11355
      _Version        =   196608
      _ExtentX        =   20029
      _ExtentY        =   10901
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
      MaxCols         =   10
      MaxRows         =   30
      ScrollBars      =   2
      SpreadDesigner  =   "frmPartWeek.frx":0000
      Appearance      =   1
   End
   Begin VB.TextBox txtEndDate 
      Enabled         =   0   'False
      Height          =   330
      Left            =   8820
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtStartDate 
      Enabled         =   0   'False
      Height          =   330
      Left            =   8820
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   1185
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   690
      Left            =   180
      TabIndex        =   3
      Top             =   135
      Width           =   8520
      _Version        =   65536
      _ExtentX        =   15028
      _ExtentY        =   1217
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
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   2655
         TabIndex        =   4
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36467
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   330
         Left            =   1080
         TabIndex        =   5
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36446
      End
      Begin MSForms.CommandButton cmdPr 
         Height          =   420
         Left            =   6075
         TabIndex        =   8
         Top             =   90
         Width           =   1410
         Caption         =   "출력확인"
         PicturePosition =   327683
         Size            =   "2487;741"
         Picture         =   "frmPartWeek.frx":0881
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   420
         Left            =   4320
         TabIndex        =   7
         Top             =   90
         Width           =   1680
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;741"
         Picture         =   "frmPartWeek.frx":115B
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmPartWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPr_Click()
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
    strHead3 = "/f3" & "조회일자 : " & " (From: " & Format(dtDate.Value, "yyyy-MM-dd") & ") ~  (To: " & _
                                                    Format(dtToDate.Value, "yyyy-MM-dd") & ")"
    strHead5 = "/f4" & "    "
    
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
    Dim sStartDate      As String
    Dim sEndDate        As String
    Dim sWeek(7)        As String
    
    sStartDate = txtStartDate.Text
    sEndDate = txtEndDate.Text
    
    Call SpreadSetClear(sprWeek)
    
    For i = 0 To 6
        sWeek(i) = Format(CDate(sStartDate) + i, "yyyy-MM-dd")
        sprWeek.Row = 0
        sprWeek.Col = i + 3: sprWeek.Text = Format(sWeek(i), "yyyy-MM-dd aaaa")
    Next
    
    sWeek(7) = "9999-99-99"
    
    GoSub Get_Data
    GoSub Total_Calculate
    Exit Sub
    
    
    
Get_Data:
    StrSql = ""
    StrSql = StrSql & " SELECT SLipno1, Codenm,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(0) & "', 1,  '0'))  sWeek0,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(1) & "', 1,  '0'))  sWeek1,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(2) & "', 1,  '0'))  sWeek2,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(3) & "', 1,  '0'))  sWeek3,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(4) & "', 1,  '0'))  sWeek4,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(5) & "', 1,  '0'))  sWeek5,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(6) & "', 1,  '0'))  sWeek6,"
    StrSql = StrSql & "        Sum(Decode(JeobsuDt, '" & sWeek(7) & "', '', 1 )) LineTotal "
    StrSql = StrSql & " FROM(  SELECT TO_CHAR(a.COLLDate,'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "               a.Jeobsut1, a.Jeobsut2, a.SLipno1, b.Codenm"
    StrSql = StrSql & "        FROM   TWEXAM_ORDER   a,"
    StrSql = StrSql & "               TWEXAM_SPECODE b"
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sStartDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sEndDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND   (a.SLipno1  BETWEEN 11 and 51)"
    StrSql = StrSql & "        AND    a.JeobsuYn  = '*'"
    StrSql = StrSql & "        AND    a.SLipno1   = b.Codeky"
    StrSql = StrSql & "        AND    b.Codegu    = '12')"
    StrSql = StrSql & " GROUP  BY SLipno1, Codenm"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        sprWeek.Row = sprWeek.DataRowCnt + 1
        sprWeek.Col = 1:  sprWeek.Text = adoSet.Fields("SLipno1").Value & ""
        sprWeek.Col = 2:  sprWeek.Text = adoSet.Fields("Codenm").Value & ""
        sprWeek.Col = 3:  sprWeek.Text = adoSet.Fields("sWeek0").Value & ""
        sprWeek.Col = 4:  sprWeek.Text = adoSet.Fields("sWeek1").Value & ""
        sprWeek.Col = 5:  sprWeek.Text = adoSet.Fields("sWeek2").Value & ""
        sprWeek.Col = 6:  sprWeek.Text = adoSet.Fields("sWeek3").Value & ""
        sprWeek.Col = 7:  sprWeek.Text = adoSet.Fields("sWeek4").Value & ""
        sprWeek.Col = 8:  sprWeek.Text = adoSet.Fields("sWeek5").Value & ""
        sprWeek.Col = 9:  sprWeek.Text = adoSet.Fields("sWeek6").Value & ""
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
    
    
    For i = 1 To sprWeek.DataRowCnt
        For j = 3 To sprWeek.DataColCnt
            sprWeek.Row = i
            sprWeek.Col = j
            If sprWeek.Text = "0" Then
                sprWeek.Text = ""
            End If
        Next
    Next
    Return

    
    
    
End Sub

Private Sub dtDate_Click()
    
    txtStartDate.Text = Format(dtDate.Value, "yyyy-MM-dd")
    txtEndDate.Text = Format(CDate(txtStartDate.Text) + 6, "yyyy-MM-dd")
    dtToDate.Value = txtEndDate.Text
    
End Sub

Private Sub dtDate_CloseUp()
    
    txtStartDate.Text = Format(dtDate.Value, "yyyy-MM-dd")
    txtEndDate.Text = Format(CDate(txtStartDate.Text) + 6, "yyyy-MM-dd")
    dtToDate.Value = txtEndDate.Text
    
End Sub

Private Sub Form_Load()
    
    dtDate.Value = Format(CDate(Dual_Date_Get("yyyy-MM-dd")) - 7, " yyyy-MM-dd")
    Call dtDate_CloseUp
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

