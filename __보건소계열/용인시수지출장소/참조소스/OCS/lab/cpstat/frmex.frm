VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmEx 
   Caption         =   "외부검사통계"
   ClientHeight    =   7740
   ClientLeft      =   105
   ClientTop       =   1080
   ClientWidth     =   11835
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
   ScaleHeight     =   7740
   ScaleWidth      =   11835
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprEx 
      Height          =   6405
      Left            =   135
      TabIndex        =   5
      Top             =   900
      Width           =   11580
      _Version        =   196608
      _ExtentX        =   20426
      _ExtentY        =   11298
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
      MaxCols         =   100
      SpreadDesigner  =   "frmEx.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   4560
      _Version        =   65536
      _ExtentX        =   8043
      _ExtentY        =   1058
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
         Left            =   2520
         TabIndex        =   1
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36446
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36446
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   180
         Width           =   780
      End
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   600
      Left            =   6435
      TabIndex        =   6
      Top             =   225
      Width           =   1770
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "3122;1058"
      Picture         =   "frmEx.frx":2A19
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   600
      Left            =   4770
      TabIndex        =   4
      Top             =   225
      Width           =   1680
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2963;1058"
      Picture         =   "frmEx.frx":32F3
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
Attribute VB_Name = "frmEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQueryo_Click()

End Sub

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
    
    For i = 1 To 80
        sPortBar = sPortBar & "━"
    Next
    
    
    If sprEx.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 외부의뢰검사(과별) 통계"
    strHead3 = "/f3" & "기간 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & ""
    
    sprEx.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                       strFont2 + "/l" + strHead3 + _
                                       strFont2 + strHead5
    sprEx.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(110) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprEx.PrintMarginLeft = 0
    sprEx.PrintMarginRight = 0
    sprEx.PrintMarginTop = 0
    sprEx.PrintMarginBottom = 0
    sprEx.PrintColHeaders = True
    sprEx.PrintRowHeaders = False
    sprEx.PrintBorder = True
    sprEx.PrintColor = True
    sprEx.PrintGrid = True
    sprEx.PrintShadows = True
    sprEx.PrintUseDataMax = False
    
    sprEx.Row = 1: sprEx.Row2 = sprEx.DataRowCnt
    sprEx.Col = 2: sprEx.Col2 = sprEx.DataColCnt
    sprEx.PrintType = SS_PRINT_CELL_RANGE
    sprEx.PrintOrientation = PrintOrientationLandscape
    sprEx.Action = SS_ACTION_PRINT


End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate       As String
    Dim sToDate       As String
    Dim sDeptC()      As String
    Dim sDeptC1()     As String
    Dim iCol          As Integer
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    
    GoSub Init_Spread
    GoSub DeptCode_Vinding
    GoSub Main_PRoc
    GoSub Total_Calculate
    
    sprEx.MaxRows = sprEx.DataRowCnt
    sprEx.MaxCols = sprEx.DataColCnt
    
    Exit Sub
    
Rem ------------------------------------------------------------------------------------------
    
Init_Spread:
    sprEx.ReDraw = False
    sprEx.MaxRows = 0
    sprEx.MaxRows = 300
    sprEx.RowHeight(-1) = 10.5
    sprEx.MaxCols = 100
    sprEx.ColWidth(-1) = 4.5
    sprEx.ColWidth(1) = 7
    sprEx.ColWidth(2) = 20
    sprEx.ReDraw = True
    
    Return
    
    
DeptCode_Vinding:
    StrSql = ""
    StrSql = StrSql & " SELECT b.DEPTCODE"
    StrSql = StrSql & " FROM   TWEXAM_GENERAL_SUB a,"
    StrSql = StrSql & "        TWEXAM_GENERAL     b "
    'StrSql = StrSql & "        TWEXAM_ITEMML      c "
    StrSql = StrSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    StrSql = StrSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    StrSql = StrSql & " AND    a.Codegu  = 'W'"
    
    'StrSql = StrSql & " AND    a.ItemCd = c.Codeky(+)"
    'StrSql = StrSql & " AND    c.GeomsaGb = 'W'"
    
    StrSql = StrSql & " AND    a.SLipno1  > 0 "
    StrSql = StrSql & " AND    a.SLipno1  < 52"
    StrSql = StrSql & " AND    a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & " AND    a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & " AND    a.SLipno2  = b.SLipno2(+)"
    StrSql = StrSql & " GROUP  BY b.DeptCode"
    StrSql = StrSql & " ORDER  BY b.DeptCode"
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    ReDim sDeptC(adoSet.RecordCount - 1)
    ReDim sDeptC1(adoSet.RecordCount - 1)
    
    i = 0
    Do Until adoSet.EOF
        sDeptC(i) = Trim(adoSet.Fields("Deptcode").Value & "")
        sDeptC1(i) = "a" & Trim(adoSet.Fields("Deptcode").Value & "")
        sprEx.Row = 0
        sprEx.Col = i + 3: sprEx.Text = sDeptC(i)
              
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

Main_PRoc:
    
    'kwak
    StrSql = ""
    StrSql = StrSql & " SELECT ItemCd, ItemName,"
    
    For i = 0 To UBound(sDeptC)
        StrSql = StrSql & "    SUM(Decode(RTRIM(DeptCode), '" & RTrim(sDeptC(i)) & "',1, '')) " & Trim(sDeptC1(i)) & ","
    Next
    StrSql = StrSql & "        COUNT(Decode(RTRIM(DeptCode), '99','',1)) LineTotal"
    
    StrSql = StrSql & " FROM(  SELECT TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    StrSql = StrSql & "               a.ItemCd, c.ItemNM ItemName, b.DeptCode"
    StrSql = StrSql & "        FROM   TWEXAM_GENERAL_SUB a,"
    StrSql = StrSql & "               TWEXAM_GENERAL     b, "
    StrSql = StrSql & "               TWEXAM_ITEMML      c"
    StrSql = StrSql & "        WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    StrSql = StrSql & "        AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    StrSql = StrSql & "        AND    a.Codegu   = 'W'"
    StrSql = StrSql & "        AND    a.ItemCd = c.Codeky(+)"
    StrSql = StrSql & "        AND    c.GeomsaGb = 'W'"
    
    StrSql = StrSql & "        AND    a.SLipno1  > 0 "
    StrSql = StrSql & "        AND    a.SLipno1  < 52"
    StrSql = StrSql & "        AND    a.JeobsuDt = b.JeobsuDt(+)"
    StrSql = StrSql & "        AND    a.SLipno1  = b.SLipno1(+)"
    StrSql = StrSql & "        AND    a.SLipno2  = b.SLipno2(+))"
    StrSql = StrSql & " GROUP  BY iTemCd, ItemName"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprEx.Row = sprEx.DataRowCnt + 1
        sprEx.Col = 1: sprEx.Text = adoSet.Fields("ItemCd").Value & ""
        sprEx.Col = 2: sprEx.Text = adoSet.Fields("ItemName").Value & ""
        
        iCol = 3
        For i = 0 To UBound(sDeptC)
            
            sprEx.Row = sprEx.DataRowCnt
            sprEx.Col = i + 3: sprEx.Text = adoSet.Fields(sDeptC1(i)).Value & ""
        Next
        sprEx.Col = UBound(sDeptC) + 4: sprEx.Text = adoSet.Fields("LineTotal").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    sprEx.MaxCols = sprEx.DataRowCnt
    sprEx.Col = sprEx.MaxCols
    sprEx.Row = 0: sprEx.Text = "합계"
    
    
    
    Return
    
    
    
Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    If sprEx.DataRowCnt = sprEx.MaxRows Then
        sprEx.MaxRows = sprEx.MaxRows + 2
    End If
    
    
    iLastSprRow = sprEx.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprEx.MaxCols
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprEx.Row = i
            sprEx.Col = j
            If Trim(sprEx.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprEx.Text)
            End If
        Next
        sprEx.Row = iLastSprRow + 2
        sprEx.Col = j
        sprEx.Text = nCalSum
    Next
    
    sprEx.Row = sprEx.DataRowCnt
    sprEx.Col = 2: sprEx.Text = "합계"
    Return


End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

    sprEx.Row = 0
    sprEx.Row2 = 0
    sprEx.Col = 3
    sprEx.Col2 = sprEx.MaxCols
    sprEx.BlockMode = True
    sprEx.Text = " "
    sprEx.BlockMode = False

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
