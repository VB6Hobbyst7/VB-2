VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmJupsuList 
   Caption         =   "접수LIST"
   ClientHeight    =   8175
   ClientLeft      =   105
   ClientTop       =   645
   ClientWidth     =   11685
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
   ScaleHeight     =   8175
   ScaleWidth      =   11685
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJupsuList.frx":0000
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      ButtonWidth     =   1270
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit of 접수List"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbSLip 
      Height          =   300
      Left            =   2610
      Style           =   2  '드롭다운 목록
      TabIndex        =   10
      Top             =   450
      Width           =   2580
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "응급Order"
      Height          =   270
      Index           =   0
      Left            =   6390
      TabIndex        =   9
      Top             =   900
      Value           =   -1  'True
      Width           =   1140
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "응급제외"
      Height          =   270
      Index           =   1
      Left            =   7560
      TabIndex        =   8
      Top             =   900
      Width           =   1050
   End
   Begin FPSpreadADO.fpSpread sprPtList 
      Height          =   6225
      Left            =   315
      TabIndex        =   0
      Top             =   1305
      Width           =   11220
      _Version        =   196608
      _ExtentX        =   19791
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
      GridShowHoriz   =   0   'False
      MaxCols         =   14
      ScrollBars      =   2
      SpreadDesigner  =   "frmJupsuList.frx":0324
      Appearance      =   2
   End
   Begin Threed.SSPanel panelCheck 
      Height          =   330
      Left            =   6210
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   5280
      _Version        =   65536
      _ExtentX        =   9313
      _ExtentY        =   582
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
      Begin VB.CheckBox chkStatus 
         Caption         =   "결과완료"
         Height          =   225
         Index           =   3
         Left            =   3195
         TabIndex        =   5
         Tag             =   "C"
         Top             =   45
         Width           =   1050
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "부분결과"
         Height          =   225
         Index           =   2
         Left            =   2115
         TabIndex        =   4
         Tag             =   "P"
         Top             =   45
         Width           =   1050
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "미확인"
         Height          =   225
         Index           =   1
         Left            =   1170
         TabIndex        =   3
         Tag             =   "U"
         Top             =   45
         Width           =   915
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "접수중"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Tag             =   "R"
         Top             =   45
         Value           =   1  '확인
         Width           =   915
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   330
      Left            =   1575
      TabIndex        =   6
      Top             =   855
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "접수일자"
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
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   330
      Left            =   1575
      TabIndex        =   7
      Top             =   450
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "검사종목"
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
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   4500
      TabIndex        =   11
      Top             =   855
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   24576003
      CurrentDate     =   36444
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   2610
      TabIndex        =   12
      Top             =   855
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   24576003
      CurrentDate     =   36444
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   465
      Left            =   10080
      TabIndex        =   14
      Top             =   855
      Width           =   1410
      Caption         =   "출력"
      PicturePosition =   327683
      Size            =   "2487;820"
      Picture         =   "frmJupsuList.frx":4127
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   465
      Left            =   8685
      TabIndex        =   13
      Top             =   855
      Width           =   1410
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2487;820"
      Picture         =   "frmJupsuList.frx":4A01
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmJupsuList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdPrint_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    If sprPtList.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "접수내역 LIST"
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFrDate.Value, "yyyy-MM-dd hh:mm ampm") & " / " & _
                                               Format(dtToDate.Value, "yyyy-MM-dd hh:mm ampm")
    
    sprPtList.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                            strFont(1) + "/n" + sBarLine + strFont(1)
    sprPtList.PrintFooter = "/f2" & "/l" & sBarLine & _
                            "/n" & Space(80) & "Page : " & "/p" & " of " & sprPtList.PrintPageCount
    sprPtList.PrintMarginLeft = 0
    sprPtList.PrintMarginRight = 0
    sprPtList.PrintMarginTop = 0
    sprPtList.PrintMarginBottom = 0
    sprPtList.PrintColHeaders = True
    sprPtList.PrintRowHeaders = True
    sprPtList.PrintBorder = False
    sprPtList.PrintColor = False
    sprPtList.PrintGrid = True
    sprPtList.PrintShadows = True
    sprPtList.PrintUseDataMax = False
    sprPtList.Row = 1
    sprPtList.Row2 = sprPtList.DataRowCnt
    sprPtList.Col = 1
    sprPtList.Col2 = sprPtList.MaxCols
    sprPtList.PrintOrientation = 1
    sprPtList.PrintOrientation = PrintOrientationPortrait
    sprPtList.PrintType = PrintTypeCellRange
    sprPtList.Action = ActionPrint

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate1    As String
    Dim sFrDate     As String
    Dim sFrHH       As String
    Dim sFrMM       As String
    
    Dim sToDate1    As String
    Dim sToDate     As String
    Dim sToHH       As String
    Dim sToMM       As String
    
    Dim sCompTxt    As String
    Dim sWhere      As String
    
    
    If cmbSLip.ListIndex = -1 Then
        MsgBox "조회 SLip 을 선택하세요!......"
        Exit Sub
    End If
    
    sFrDate = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    sToDate = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    
    sFrDate1 = Format(dtFrDate.Value, "yyyy-MM-dd hh:mm")
    sFrDate = Left(sFrDate1, 10)
    sFrHH = Mid(sFrDate1, 12, 2)
    sFrMM = Right(sFrDate1, 2)
    
    sToDate1 = Format(dtToDate.Value, "yyyy-MM-dd hh:mm")
    sToDate = Left(sToDate1, 10)
    sToHH = Mid(sToDate1, 12, 2)
    sToMM = Right(sToDate1, 2)
    
    GoSub SPread_Set_Reset
    GoSub Select_Data:
    
    Exit Sub
    
    
    
    
SPread_Set_Reset:
    sprPtList.MaxRows = 0
    sprPtList.MaxRows = 500
    sprPtList.RowHeight(-1) = 10.91
    Return
    
    
Select_Data:
    strSql = ""
    strSql = strSql & " SELECT a.ROWID,"
    strSql = strSql & "        a.Ptno, c.Sname, c.Sex, c.ageYY, c.RoomCode, a.DeptCode, a.GBER,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        LTRIM(TO_CHAR(A.JeobsuT1, '00')) || ':' || LTRIM(TO_CHAR(A.JeobsuT2,'00')) jTime,"
    strSql = strSql & "        A.iTEMcd, B.iTEMNM itemName, D.NAME,"
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollDate,'yyyy-MM-dd')) || ' ' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollHH,  '00'))  || ':' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollMM,  '00'))  CollTime,"
    strSql = strSql & "        e.Status, e.SLipno1, e.SLipno2"
    strSql = strSql & " FROM   TWEXAM_Order   a,"
    strSql = strSql & "        TWEXAM_ITEMML  b,"
    strSql = strSql & "        TWEXAM_IDNOMST c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS     d,"
    strSql = strSql & "        TWEXAM_General e"
    strSql = strSql & " WHERE  LTRIM(TO_CHAR(a.CollDate, 'YYYY-MM-DD')) || ' ' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollHH, '00')) || ':' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollMM, '00'))   BETWEEN  '" & sFrDate1 & "'"
    strSql = strSql & "                                         AND      '" & sToDate1 & "'"
    strSql = strSql & " AND    A.JEOBSUYN   = '*'"
    strSql = strSql & " AND    A.GBCH       = 'Y'"
    strSql = strSql & " AND    a.SLipno1    = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    A.ITEMCD     = B.CODEKY"
    strSql = strSql & " AND    A.PTNO       = C.PTNO(+)"
    strSql = strSql & " AND    TO_NUMBER(A.COLLID)    = D.IDNUMBER(+)"
    strSql = strSql & " AND   (D.PROGRAMID = ' '  OR D.PROGRAMID IS NULL)"
    strSql = strSql & " AND    a.JeobsuDt  = e.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = e.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = e.Orderno(+)"
    'strSql = strSql & " AND    e.Status    = 'R'"       '
    If optWhere(0).Value = True Then strSql = strSql & " AND (a.GBER  = 'E'  OR RTRIM(a.DeptCode)  = 'ER')"
    If optWhere(1).Value = True Then
        strSql = strSql & " AND  RTRIM(a.DeptCode) != 'ER'"
        strSql = strSql & " AND (a.GBER != 'E' OR a.GBER IS NULL)"
    End If

    GoSub Where_Sql_Sum
    strSql = strSql & sWhere

    strSql = strSql & " UNION  ALL "
    strSql = strSql & " SELECT DISTINCT a.ROWID,"
    strSql = strSql & "        a.Ptno, c.Sname, c.Sex, c.ageYY, c.RoomCode,a.DeptCode, a.GBER,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        LTRIM(TO_CHAR(A.JeobsuT1, '00')) || ':' || LTRIM(TO_CHAR(A.JeobsuT2,'00')) jTime,"
    strSql = strSql & "        A.iTEMcd, B.RoutinNM itemName, D.NAME,"
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollDate,'yyyy-MM-dd')) || ' ' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollHH,  '00'))  || ':' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollMM,  '00'))  CollTime,"
    strSql = strSql & "        e.Status, e.SLipno1, e.SLipno2"
    strSql = strSql & " FROM   TWEXAM_Order   a,"
    strSql = strSql & "        TWEXAM_ROUTINE b,"
    strSql = strSql & "        TWEXAM_IDNOMST c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS     d,"
    strSql = strSql & "        TWEXAM_General e "
    strSql = strSql & " WHERE  LTRIM(TO_CHAR(a.CollDate, 'YYYY-MM-DD')) || ' ' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollHH, '00')) || ':' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.CollMM, '00'))   BETWEEN  '" & sFrDate1 & "'"
    strSql = strSql & "                                         AND      '" & sToDate1 & "'"
    strSql = strSql & " AND    A.JEOBSUYN  = '*'"
    strSql = strSql & " AND    A.GBCH      = 'Y'"
    strSql = strSql & " AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    A.ITEMCD    = B.ROUTINCD"
    strSql = strSql & " AND    A.PTNO      = C.PTNO(+)"
    strSql = strSql & " AND    TO_NUMBER(A.COLLID)    = D.IDNUMBER(+)"
    strSql = strSql & " AND   (D.PROGRAMID = ' '  OR D.PROGRAMID IS NULL)"
    strSql = strSql & " AND    a.JeobsuDt  = e.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = e.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = e.Orderno(+)"
    'strSql = strSql & " AND    e.Status    = 'R'"
    If optWhere(0).Value = True Then strSql = strSql & " AND (a.GBER  = 'E'  OR RTRIM(a.DeptCode)  = 'ER')"
    If optWhere(1).Value = True Then
        strSql = strSql & " AND  RTRIM(a.DeptCode) != 'ER'"
        strSql = strSql & " AND (a.GBER != 'E' OR a.GBER IS NULL)"
    End If
    
    GoSub Where_Sql_Sum
    strSql = strSql & sWhere
    strSql = strSql & " ORDER  BY JeobsuDt, SLipno1, SLipno2"

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprPtList.Row = sprPtList.DataRowCnt + 1
        
        If sCompTxt <> adoSet.Fields("Ptno").Value & "" Then
            sprPtList.Col = 1:  sprPtList.Text = adoSet.Fields("Ptno").Value & ""
            sprPtList.Col = 2:  sprPtList.Text = adoSet.Fields("Sname").Value & ""
            sprPtList.Col = 3:  sprPtList.Text = adoSet.Fields("Sex").Value & ""
            sprPtList.Col = 4:  sprPtList.Text = adoSet.Fields("AgeYY").Value & ""
            sprPtList.Col = 5:  sprPtList.Text = adoSet.Fields("RoomCode").Value & ""
            
            sprPtList.Col = 1:             sprPtList.Col2 = sprPtList.MaxCols
            sprPtList.Row = sprPtList.Row: sprPtList.Row2 = sprPtList.Row
            sprPtList.BlockMode = True
            sprPtList.CellBorderType = SS_BORDER_TYPE_TOP
            sprPtList.CellBorderStyle = CellBorderStyleSolid
            sprPtList.Action = ActionSetCellBorder
            sprPtList.BlockMode = False
        End If
        
        sprPtList.Col = 6:  sprPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprPtList.Col = 7:  sprPtList.Text = adoSet.Fields("SLipno2").Value & ""
        sprPtList.Col = 8:  sprPtList.Text = adoSet.Fields("ItemCD").Value & ""
        sprPtList.Col = 9:  sprPtList.Text = adoSet.Fields("itemName").Value & ""
        sprPtList.Col = 10: sprPtList.Text = adoSet.Fields("Colltime").Value & ""
        sprPtList.Col = 11: sprPtList.Text = adoSet.Fields("Name").Value & ""
        sprPtList.Col = 12: sprPtList.Text = adoSet.Fields("Deptcode").Value & ""
        sprPtList.Col = 13: sprPtList.Text = adoSet.Fields("GBER").Value & ""
        
        sprPtList.Col = 14
        Select Case adoSet.Fields("Status").Value & ""
            Case "C":  sprPtList.Text = "결과완료"
            Case "P":  sprPtList.Text = "부분결과"
            Case "R":  sprPtList.Text = "접수중"
            Case "U":  sprPtList.Text = "미확인"
            Case Else: sprPtList.Text = adoSet.Fields("Status").Value & ""
        End Select
        
        sCompTxt = adoSet.Fields("Ptno").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    


Where_Sql_Sum:
    sWhere = Set_CheckBox_SqlSum(chkStatus, "e.Status")
    
    Return
    
    
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd") & " 00:01"
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd") & " 23:59"
    
    GoSub Get_SLip
    Call SetComboBox(Me.cmbSLip, Format(GiExamNumb, "00"), 2)
    Exit Sub
    
    
Get_SLip:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '90'"
    strSql = strSql & " Order  By Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        Me.cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                                adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
