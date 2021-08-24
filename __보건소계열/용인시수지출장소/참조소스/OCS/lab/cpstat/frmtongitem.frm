VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmTongitem 
   Caption         =   "검사항목별 과별 통계"
   ClientHeight    =   4635
   ClientLeft      =   240
   ClientTop       =   2355
   ClientWidth     =   11325
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
   ScaleHeight     =   4635
   ScaleWidth      =   11325
   WindowState     =   2  '최대화
   Begin VB.ComboBox cmbSLip 
      Height          =   300
      Left            =   270
      Style           =   2  '드롭다운 목록
      TabIndex        =   7
      Top             =   855
      Width           =   2490
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   270
      TabIndex        =   1
      Top             =   135
      Width           =   5235
      _Version        =   65536
      _ExtentX        =   9234
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
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   2520
         TabIndex        =   2
         Top             =   180
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
         TabIndex        =   3
         Top             =   180
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
         TabIndex        =   4
         Top             =   225
         Width           =   780
      End
   End
   Begin FPSpreadADO.fpSpread sprDetail 
      Height          =   6090
      Left            =   3015
      TabIndex        =   0
      Top             =   1260
      Width           =   8610
      _Version        =   196608
      _ExtentX        =   15187
      _ExtentY        =   10742
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmTongitem.frx":0000
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread sprItemList 
      Height          =   6090
      Left            =   225
      TabIndex        =   6
      Top             =   1260
      Width           =   2760
      _Version        =   196608
      _ExtentX        =   4868
      _ExtentY        =   10742
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
      MaxCols         =   2
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmTongitem.frx":4CF0
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   465
      Left            =   9540
      TabIndex        =   5
      Top             =   675
      Width           =   2040
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "3598;820"
      Picture         =   "frmTongitem.frx":5818
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
Attribute VB_Name = "frmTongitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSLip_Click()
    If cmbSLip.ListIndex = -1 Then Exit Sub
    
    Call SpreadSetClear(sprItemList)
    
    StrSql = ""
    StrSql = StrSql & " SELECT CODEKY, ITEMNM"
    StrSql = StrSql & " FROM   TWEXAM_ITEMML"
    StrSql = StrSql & " WHERE  CODEKY  LIKE '" & Left(cmbSLip.Text, 2) & "%'"
    StrSql = StrSql & " ORDER  BY CODEKY"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprItemList.Row = sprItemList.DataRowCnt + 1
        sprItemList.Col = 1: sprItemList.Text = adoSet.Fields("Codeky").Value & ""
        sprItemList.Col = 2: sprItemList.Text = adoSet.Fields("ItemNM").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
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
    
    
    If sprDetail.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "검사항목별 통계"
    strHead3 = "/f3" & "/l" & "기간 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " / " & _
                                   Format(dtToDate.Value, "yyyy-MM-dd")
    sprItemList.Row = sprItemList.ActiveRow
    sprItemList.Col = 2
    
    strHead5 = "/f4" & "/l" & "검사항목: " & sprItemList.Text
    
    sprDetail.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                     strFont2 + strHead3 + "/n" + _
                                     strFont2 + strHead5
    sprDetail.PrintMarginLeft = 0
    sprDetail.PrintMarginRight = 0
    sprDetail.PrintMarginTop = 0
    sprDetail.PrintMarginBottom = 0
    sprDetail.PrintColHeaders = True
    sprDetail.PrintRowHeaders = True
    sprDetail.PrintBorder = True
    sprDetail.PrintColor = False
    sprDetail.PrintGrid = True
    sprDetail.PrintShadows = True
    sprDetail.PrintUseDataMax = False
    
    sprDetail.Row = 1: sprDetail.Row2 = sprDetail.DataRowCnt
    sprDetail.Col = 2: sprDetail.Col2 = sprDetail.DataColCnt
    sprDetail.PrintType = SS_PRINT_CELL_RANGE
    sprDetail.PrintOrientation = PrintOrientationLandscape
    sprDetail.Action = SS_ACTION_PRINT


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

Private Sub sprItemList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sDeptC()        As String
    Dim sItemCd         As String
    Dim sFrDate         As String
    Dim sToDate         As String
    
    
    If Row = 0 Then Exit Sub
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call SpreadSetClear(sprDetail)
    
    sprItemList.Row = Row
    sprItemList.Col = 1
    sItemCd = sprItemList.Text
    
    Screen.MousePointer = vbHourglass
    
    GoSub Get_DeptArray
    GoSub Main_Data_Select
    GoSub Calcurate_Spread
    
    Screen.MousePointer = vbDefault
    Exit Sub




Get_DeptArray:
    StrSql = ""
    StrSql = StrSql & " SELECT b.DeptCode"
    StrSql = StrSql & " FROM   TWEXAM_General_Sub a,"
    StrSql = StrSql & "        TWEXAM_General     b"
    StrSql = StrSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    StrSql = StrSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    StrSql = StrSql & " AND    a.ItemCD    = '" & sItemCd & "'"
    StrSql = StrSql & " AND    a.JeobsuDt  = b.JeobsuDt(+)"
    StrSql = StrSql & " AND    a.SLipno1   = b.SLipno1(+)"
    StrSql = StrSql & " AND    a.SLipno2   = b.SLipno2(+)"
    StrSql = StrSql & " GROUP BY b.DeptCode"
    
    If False = adoSetOpen(StrSql, adoSet) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ReDim sDeptC(adoSet.RecordCount - 1)
    i = 0
    Do Until adoSet.EOF
        sDeptC(i) = Trim(adoSet.Fields("DeptCode").Value & "")
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    sprDetail.Row = 0
    sprDetail.Col = 1: sprDetail.Text = "접수일자"
    sprDetail.ColWidth(1) = 10
    
    For i = 0 To UBound(sDeptC)
        sprDetail.Row = 0
        sprDetail.Col = i + 2
        sprDetail.Text = Trim(sDeptC(i))
        sprDetail.ColWidth(i + 2) = 5
    Next
    sprDetail.MaxCols = UBound(sDeptC) + 3
    sprDetail.Row = 0
    sprDetail.Col = sprDetail.MaxCols: sprDetail.Text = "합계"
    Return
    

Main_Data_Select:
    
    StrSql = ""
    StrSql = StrSql & " SELECT JEOBSUDT, "
    
    For i = 0 To UBound(sDeptC)
        StrSql = StrSql & "    SUM(Decode(RTRIM(DeptCode), '" & sDeptC(i) & "', CNT, '')) " & Trim(sDeptC(i)) & "a" & ","
    Next
    
    StrSql = StrSql & "        SUM(Decode(RTRIM(DeptCode), '', '', CNT ))  LineTotal"
    
    StrSql = StrSql & " FROM ( SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    StrSql = StrSql & "               B.DeptCode, COUNT(*) CNT"
    StrSql = StrSql & "        FROM   TWEXAM_General_Sub a,"
    StrSql = StrSql & "               TWEXAM_General     b"
    StrSql = StrSql & "        WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    StrSql = StrSql & "        AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    StrSql = StrSql & "        AND    a.iTemCD    = '" & sItemCd & "'"
    StrSql = StrSql & "        AND    a.JeobsuDt  = b.JeobsuDt(+)"
    StrSql = StrSql & "        AND    a.SLipno1   = b.SLipno1(+)"
    StrSql = StrSql & "        AND    a.SLipno2   = b.SLipno2(+)"
    StrSql = StrSql & "     GROUP BY a.JeobsuDt, b.DeptCode)"
    StrSql = StrSql & " GROUP BY JEOBSUDT"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprDetail.Row = sprDetail.DataRowCnt + 1
        sprDetail.Col = 1: sprDetail.Text = Format(adoSet.Fields("JeobsuDt").Value & "", "yyyy-MM-dd aaa")
        For i = 0 To UBound(sDeptC)
            sprDetail.Col = i + 2: sprDetail.Text = adoSet.Fields(sDeptC(i) & "a").Value & ""
        Next
        sprDetail.Col = sprDetail.MaxCols: sprDetail.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Return
    
    
Calcurate_Spread:
    Dim iCount      As String
    Dim iCol        As Integer
    Dim iRow        As Integer
    Dim iDataRow    As Integer
    
    sprDetail.Row = sprDetail.DataRowCnt + 2
    sprDetail.Col = 1: sprDetail.Text = "총합계"
    
    iDataRow = sprDetail.DataRowCnt
    
    For iCol = 2 To sprDetail.MaxCols
        sprDetail.Col = iCol
        For iRow = 1 To iDataRow
            sprDetail.Row = iRow
            iCount = Val(iCount) + Val(sprDetail.Text)
        Next
        sprDetail.Row = iDataRow
        sprDetail.Text = iCount
        iCount = 0
    Next
    
    
    Return
    
End Sub

