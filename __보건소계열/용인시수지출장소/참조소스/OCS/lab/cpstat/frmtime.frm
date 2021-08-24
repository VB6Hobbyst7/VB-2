VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmTime 
   Caption         =   "시간대별 항목별통계"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   1890
   ClientWidth     =   11775
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
   ScaleWidth      =   11775
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   135
      TabIndex        =   0
      Top             =   90
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
         TabIndex        =   1
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
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Width           =   780
      End
   End
   Begin FPSpreadADO.fpSpread sprDetail 
      Height          =   7035
      Left            =   135
      TabIndex        =   4
      Top             =   810
      Width           =   11490
      _Version        =   196608
      _ExtentX        =   20267
      _ExtentY        =   12409
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
      MaxCols         =   26
      ScrollBars      =   2
      SpreadDesigner  =   "frmTime.frx":0000
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   645
      Left            =   7380
      TabIndex        =   6
      Top             =   90
      Width           =   1815
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "3201;1138"
      Picture         =   "frmTime.frx":4357
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   645
      Left            =   5400
      TabIndex        =   5
      Top             =   90
      Width           =   1950
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "3440;1138"
      Picture         =   "frmTime.frx":4C31
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
Attribute VB_Name = "frmTime"
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
    
    For i = 1 To 80
        sPortBar = sPortBar & "━"
    Next
    
    DoEvents:
    sprDetail.ColWidth(-1) = 3.66
    sprDetail.ColWidth(1) = 18.13
    sprDetail.ColWidth(26) = 5.38
    
    
    
    If sprDetail.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 시간대별 통계 보고서"
    strHead3 = "/f3" & "기간 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & ""
    
    sprDetail.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                       strFont2 + "/l" + strHead3 + _
                                       strFont2 + strHead5
    sprDetail.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(110) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprDetail.PrintMarginLeft = 500
    sprDetail.PrintMarginRight = 0
    sprDetail.PrintMarginTop = 500
    sprDetail.PrintMarginBottom = 500
    sprDetail.PrintColHeaders = True
    sprDetail.PrintRowHeaders = False
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


    DoEvents:
    sprDetail.ColWidth(-1) = 2.63
    sprDetail.ColWidth(1) = 18.13
    sprDetail.ColWidth(26) = 4.88

End Sub

Private Sub cmdQuery_Click()

    Dim sFrDate     As String
    Dim sToDate     As String
    
    
    sprDetail.MaxRows = 0
    sprDetail.MaxRows = 100
    sprDetail.RowHeight(-1) = 10.23
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    
    GoSub GET_COUNT_Routine_DATA
    
    If sprDetail.DataRowCnt > 0 Then
        sprDetail.Row = sprDetail.DataRowCnt + 1
        sprDetail.Col = 1:
        sprDetail.Text = "********"
    End If
    
    GoSub GET_COUNT_Item_DATA
    
    GoSub Total_Calculate
        
    sprDetail.MaxRows = sprDetail.DataRowCnt
    
    Exit Sub
    


GET_COUNT_Routine_DATA:
    strSql = ""
    strSql = strSql & " SELECT ItemCd, ItemName,"
    
    For i = 0 To 23
        strSql = strSql & " SUM(Decode(JeobsuT1, " & i & ",1, '')) " & "T" & i & ","
    Next
    
    strSql = strSql & "     COUNT(Decode(JeobsuT1, 99, 0 , 1)) LineTotal"
    
    strSql = strSql & " FROM( SELECT  a.ItemCd, b.RoutinNM ItemName, "
    strSql = strSql & "               TO_NUMBER(a.COLLHH) JeobsuT1, a.COLLMM"
    strSql = strSql & "       FROM   TWEXAM_Order       a,"
    strSql = strSql & "              TWEXAM_Routine     b"
    strSql = strSql & "       WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & "       AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & "       AND    a.ItemCd    = b.RoutinCd"
    strSql = strSql & "       AND    a.SLipno1   < 52"
    strSql = strSql & "       AND    a.JeobsuYN  = '*'"
    strSql = strSql & "       GROUP BY ItemCd, b.RoutinNm, a.COLLDate, a.COLLHH, a.COLLMM)"
    strSql = strSql & " GROUP BY ItemCD, Itemname"
    strSql = strSql & " ORDER BY itemcd"

    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprDetail.Row = sprDetail.DataRowCnt + 1
        sprDetail.Col = 1: sprDetail.Text = adoSet.Fields("ItemName").Value & ""
        For i = 0 To 23
            sprDetail.Col = i + 2: sprDetail.Text = adoSet.Fields("T" & i).Value & ""
        Next
        sprDetail.Col = sprDetail.MaxCols: sprDetail.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

GET_COUNT_Item_DATA:
    strSql = ""
    strSql = strSql & " SELECT ItemCd, ItemName,"
    
    For i = 0 To 23
        strSql = strSql & " SUM(Decode(JeobsuT1, " & i & ",1, '')) " & "T" & i & ","
    Next
    
    strSql = strSql & "     COUNT(Decode(JeobsuT1, 99, 0 , 1)) LineTotal"
    strSql = strSql & " FROM( SELECT  a.ItemCd, b.ItemNM ItemName, "
    strSql = strSql & "               TO_NUMBER(a.COLLHH) JeobsuT1, a.COLLMM"
    strSql = strSql & "       FROM   TWEXAM_Order       a,"
    strSql = strSql & "              TWEXAM_ItemML      b"
    strSql = strSql & "       WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & "       AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & "       AND    a.ItemCd    = b.Codeky"
    strSql = strSql & "       AND    a.SLipno1  <  52"
    strSql = strSql & "       AND    a.JeobsuYN  = '*'"
    strSql = strSql & "       GROUP BY ItemCd, b.ItemNm, a.COLLDate, a.COLLHH, a.COLLMM)"
    strSql = strSql & " GROUP BY ItemCD, Itemname"
    strSql = strSql & " ORDER BY itemcd"

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    If adoSet.RecordCount < sprDetail.MaxRows + 3 Then
        sprDetail.MaxRows = sprDetail.MaxRows + 4
    End If
    
    Do Until adoSet.EOF
        sprDetail.Row = sprDetail.DataRowCnt + 1
        sprDetail.Col = 1: sprDetail.Text = adoSet.Fields("ItemName").Value & ""
        For i = 0 To 23
            sprDetail.Col = i + 2: sprDetail.Text = adoSet.Fields("T" & i).Value & ""
        Next
        sprDetail.Col = sprDetail.MaxCols: sprDetail.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return


Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    iLastSprRow = sprDetail.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 2 To sprDetail.DataColCnt
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprDetail.Row = i
            sprDetail.Col = j
            If Trim(sprDetail.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprDetail.Text)
            End If
        Next
        sprDetail.Row = iLastSprRow + 2
        sprDetail.Col = j
        sprDetail.Text = nCalSum
    Next
    
    sprDetail.Row = sprDetail.DataRowCnt
    sprDetail.Col = 1: sprDetail.Text = "합계"
    Return

End Sub

Private Sub Form_Load()
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
