VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmRoutine 
   Caption         =   "Routine Code 관리"
   ClientHeight    =   7050
   ClientLeft      =   285
   ClientTop       =   1275
   ClientWidth     =   11415
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11415
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   3195
      Left            =   120
      TabIndex        =   8
      Top             =   180
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   5636
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtOrderCD 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1260
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtCodate 
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Top             =   1980
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36301
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   315
         Left            =   2940
         TabIndex        =   16
         Top             =   540
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "&H"
         BevelWidth      =   1
      End
      Begin VB.ComboBox cmbJangbi 
         Height          =   300
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   1980
         Width           =   2715
      End
      Begin VB.TextBox txtYageo 
         Height          =   315
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1620
         Width           =   1335
      End
      Begin VB.TextBox txtRoutineName 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   900
         Width           =   2715
      End
      Begin VB.TextBox txtSlipCode 
         Height          =   315
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   1
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox txtSlipKey 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   540
         Width           =   435
      End
      Begin VB.ComboBox cmbSlip 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   180
         Width           =   2715
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   420
         Left            =   4230
         TabIndex        =   28
         Top             =   180
         Width           =   1320
         Caption         =   "SLip조회"
         PicturePosition =   196613
         Size            =   "2328;741"
         Picture         =   "frmRoutine.frx":0000
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   555
         Left            =   4185
         TabIndex        =   25
         Top             =   2520
         Width           =   1410
         Caption         =   "출력"
         PicturePosition =   327683
         Size            =   "2487;979"
         Picture         =   "frmRoutine.frx":08DA
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label7 
         Caption         =   "OrderCode"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   1320
         Width           =   975
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   555
         Left            =   2895
         TabIndex        =   18
         Top             =   2520
         Width           =   1275
         Caption         =   "화면정리"
         PicturePosition =   327683
         Size            =   "2249;979"
         Picture         =   "frmRoutine.frx":11B4
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   555
         Left            =   1605
         TabIndex        =   17
         Top             =   2520
         Width           =   1275
         Caption         =   "삭제확인"
         PicturePosition =   327683
         Size            =   "2249;979"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsert 
         Height          =   555
         Left            =   315
         TabIndex        =   6
         Top             =   2520
         Width           =   1275
         Caption         =   "입력확인"
         PicturePosition =   327683
         Size            =   "2249;979"
         Picture         =   "frmRoutine.frx":2946
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label6 
         Caption         =   "검사장비"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "약어"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1686
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "RoutineName"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   982
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "RoutineCode"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "검사종류"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   240
         Width           =   1035
      End
   End
   Begin FPSpreadADO.fpSpread ssItem 
      Height          =   7770
      Left            =   5880
      TabIndex        =   15
      Top             =   180
      Width           =   5895
      _Version        =   196608
      _ExtentX        =   10398
      _ExtentY        =   13705
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
      MaxCols         =   3
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmRoutine.frx":4108
      UserResize      =   1
      VisibleCols     =   500
      VisibleRows     =   500
      Appearance      =   2
   End
   Begin FPSpreadADO.fpSpread ssRoutine 
      Height          =   4365
      Left            =   120
      TabIndex        =   7
      Top             =   3465
      Width           =   5700
      _Version        =   196608
      _ExtentX        =   10054
      _ExtentY        =   7699
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
      MaxCols         =   5
      MaxRows         =   50
      ScrollBars      =   2
      SpreadDesigner  =   "frmRoutine.frx":43F0
      UserResize      =   1
      VisibleCols     =   4
      VisibleRows     =   50
      Appearance      =   2
   End
   Begin Threed.SSPanel panelPrint 
      Height          =   7665
      Left            =   90
      TabIndex        =   21
      Top             =   180
      Visible         =   0   'False
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
      _ExtentY        =   13520
      _StockProps     =   15
      Caption         =   "Routine-Code Print Frame"
      ForeColor       =   0
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      FloodColor      =   16761024
      Alignment       =   0
      Begin FPSpreadADO.fpSpread ssPrRoutine 
         Height          =   5865
         Left            =   270
         TabIndex        =   22
         Top             =   945
         Width           =   10005
         _Version        =   196608
         _ExtentX        =   17648
         _ExtentY        =   10345
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmRoutine.frx":5713
         Appearance      =   1
      End
      Begin VB.TextBox txtSLipTitle 
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   495
         Width           =   2355
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SLip종류 :"
         Height          =   195
         Left            =   765
         TabIndex        =   27
         Top             =   540
         Width           =   960
      End
      Begin MSForms.CommandButton cmdExit 
         Height          =   465
         Left            =   8370
         TabIndex        =   24
         Top             =   405
         Width           =   1635
         Caption         =   "Exit"
         PicturePosition =   327683
         Size            =   "2884;820"
         Picture         =   "frmRoutine.frx":92DF
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrintOk 
         Height          =   465
         Left            =   6885
         TabIndex        =   23
         Top             =   405
         Width           =   1500
         Caption         =   "출력확인"
         PicturePosition =   327683
         Size            =   "2646;820"
         Picture         =   "frmRoutine.frx":9BB9
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub TRANS_CLEAR()
    
    txtSlipCode.Text = ""
    txtRoutineName.Text = ""
    txtYageo.Text = ""
    cmbJangbi.ListIndex = -1
    
    ssRoutine.Row = 1
    ssRoutine.Row2 = ssRoutine.DataRowCnt
    ssRoutine.Col = 1
    ssRoutine.Col2 = ssRoutine.DataColCnt
    ssRoutine.BlockMode = True
    ssRoutine.Action = SS_ACTION_CLEAR_TEXT
    ssRoutine.BlockMode = False
    
    
End Sub

Private Sub cmbSlip_Click()
    
    If cmbSlip.ListIndex = -1 Then Exit Sub
    
    DoEvents
    txtSlipKey.Text = Left(cmbSlip.Text, 2)
    GoSub Get_ITem_Data
    
    Call TRANS_CLEAR
    Exit Sub
    
'/__________________________________________________________
Get_ITem_Data:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_iTemML"
    strSql = strSql & " WHERE  Codeky LIKE '" & Trim(txtSlipKey.Text) & "%'"
    strSql = strSql & " ORDER  BY Codeky"
    
    ssItem.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssItem.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssItem.Row = ssItem.DataRowCnt + 1
        ssItem.Col = 1: ssItem.Text = Trim(adoSet.Fields("Codeky").Value & "")
        ssItem.Col = 2: ssItem.Text = RTrim(adoSet.Fields("Itemnm").Value & "")
        ssItem.Col = 3: ssItem.Text = Trim(adoSet.Fields("Sugacd").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub cmdClear_Click()

    Call ClearForm(frmRoutine)
    mdiMain.stbMain.Panels(1).Text = ""
    
End Sub

Private Sub cmdDelete_Click()
    Dim sRoutineCode        As String
    
    
    If Trim(txtSlipKey.Text) = "" Then
        MsgBox "삭제할 검사종류가 없습니다!..[확인바람]", vbQuestion
        Exit Sub
    End If
    If Trim(txtSlipCode.Text) = "" Then
        MsgBox "삭제할 RoutineCode 가 없습니다!..[확인바람]", vbQuestion
        Exit Sub
    End If
    
    sRoutineCode = Trim(txtSlipKey.Text) & Trim(txtSlipCode.Text)
    
    sMsg = Trim(sRoutineCode) & " [" & Trim(txtRoutineName.Text) & " ] " & vbCrLf & "를 삭제하시겠습니까?"
    If vbNo = MsgBox(sMsg, vbYesNo + vbQuestion + vbDefaultButton2, "삭제확인Box") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TWEXAM_ROUTINE"
    strSql = strSql & " WHERE  RoutinCD = '" & sRoutineCode & "'"
    
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = sRoutineCode & " 를 삭제하였습니다!."
        Call TRANS_CLEAR
    Else
        mdiMain.stbMain.Panels(1).Text = sRoutineCode & " 를 어떠한 오류로 인하여 삭제치 못하였습니다!"
    End If
    
    
End Sub

Private Sub cmdExit_Click()
    
    panelPrint.Visible = False
    
End Sub

Private Sub cmdHelp_Click()
    
    If txtSlipKey.Text = "" Then Exit Sub
    
    DoEvents
    hWndReturn = txtSlipCode.hwnd
    frmQryRoutine.Show vbModal
    If Trim(txtSlipCode.Text) <> "" Then
        DoEvents
        Call txtSlipCode_LostFocus
    End If
    
End Sub

Private Sub cmdInsert_Click()
    Dim sRoutineCd      As String
    Dim sDelFlag        As String * 1
    
    Dim sCodeky         As String
    Dim sItemName       As String
    Dim sCodate         As String
    Dim sSugaCD         As String
    Dim sJangbi         As String
    Dim sSeries         As String    '연속검사 Flag
    
    sDelFlag = ""
    If Trim(txtSlipKey.Text) = "" Then Exit Sub
    If Trim(txtSlipCode.Text) = "" Then Exit Sub
    
    sRoutineCd = Trim(txtSlipKey.Text) & Trim(txtSlipCode.Text)
    sCodate = Format(dtCodate.Value, "yyyy-MM-dd")
    If cmbJangbi.ListIndex > -1 Then
        sJangbi = Trim(Left(cmbJangbi.Text, 4))
    End If

    
    GoSub Delete_Routine_Sub
    
    For i = 1 To ssRoutine.DataRowCnt
        ssRoutine.Row = i
        ssRoutine.Col = 2: sCodeky = ssRoutine.Text
        ssRoutine.Col = 3: sItemName = ssRoutine.Text
        ssRoutine.Col = 4: sSugaCD = ssRoutine.Text
        ssRoutine.Col = 5
        If ssRoutine.Value = True Then
            sSeries = "1"
        Else
            sSeries = ""
        End If
        
        ssRoutine.Col = 1
        If ssRoutine.Value = False Then
            GoSub Insert_Routine_Sub
        End If
    Next
    GoSub RECALL_ROUTINE_CODE
    
    Exit Sub

'/______________________________________________________
Delete_Routine_Sub:
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_ROUTINE"
    strSql = strSql & " WHERE  RoutinCD = '" & sRoutineCd & "'"
    Call adoExec(strSql)
    Return
    
    

Insert_Routine_Sub:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_ROUTINE"
    strSql = strSql & "       (RoutinCD, RoutinNM, Codeky, Itemnm, Sugacd, "
    strSql = strSql & "        Codate,   OrderCd,  YakCD,  GbCode, Jangbi, OutCode, Series)"
    strSql = strSql & " VALUES('" & sRoutineCd & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtRoutineName.Text)) & "',"
    strSql = strSql & "        '" & sCodeky & "',"
    strSql = strSql & "        '" & sItemName & "',"
    strSql = strSql & "        '" & Trim(sSugaCD) & "',"
    strSql = strSql & "        TO_DATE('" & sCodate & "','YYYY-MM-DD'),"
    strSql = strSql & "        '" & Trim(txtOrderCD.Text) & "',"
    strSql = strSql & "        '" & Trim(txtYageo.Text) & "',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        '" & Trim(sJangbi) & "',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        '" & sSeries & "')"
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = sRoutineCd & " 가 입력되었습니다!"
    Else
        mdiMain.stbMain.Panels(1).Text = sRoutineCd & " - 입력 오류!"
        If vbYes = MsgBox("이상황에 끝내시겠습니까?", vbYesNo + vbCritical, "Procedure Stop") Then
            Exit Sub
        End If
    End If
    Return

RECALL_ROUTINE_CODE:
    ssRoutine.Row = 1
    ssRoutine.Row2 = ssRoutine.DataRowCnt
    ssRoutine.Col = 1
    ssRoutine.Col2 = ssRoutine.DataColCnt
    ssRoutine.BlockMode = True
    ssRoutine.Action = SS_ACTION_CLEAR_TEXT
    ssRoutine.BlockMode = False
    
    Call txtSlipCode_LostFocus
    
    Return

End Sub

Private Sub cmdPrint_Click()

    If Trim(txtSlipKey.Text) = "" Then Exit Sub
        
    txtSLipTitle.Text = Mid(cmbSlip.Text, 5, Len(cmbSlip.Text) - 4)
        
    DoEvents
    panelPrint.Top = 180
    panelPrint.Left = 90
    panelPrint.Visible = True
    panelPrint.ZOrder 0
    
    DoEvents
    GoSub Get_Data_RoutineData
    Exit Sub
    


Get_Data_RoutineData:
    Dim sTempText       As String
    
    Call Spread_Set_Clear(ssPrRoutine)
    
    strSql = ""
    strSql = strSql & " SELECT RoutinCd, RoutinNm, Codeky, Itemnm, SugaCd, YakCd "
    strSql = strSql & " FROM   TWEXAM_Routine "
    strSql = strSql & " WHERE  RoutinCD  LIKE '" & txtSlipKey.Text & "%'"
    strSql = strSql & " ORDER  BY RoutinCd, Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        ssPrRoutine.Row = ssPrRoutine.DataRowCnt + 1
        
        If sTempText <> adoSet.Fields("RoutinCD").Value & "" Then
            ssPrRoutine.Col = 0: ssPrRoutine.Col2 = ssPrRoutine.MaxCols
            ssPrRoutine.Row = ssPrRoutine.Row: ssPrRoutine.Row2 = ssPrRoutine.Row
            ssPrRoutine.BlockMode = True
            ssPrRoutine.CellBorderStyle = CellBorderStyleSolid
            ssPrRoutine.CellBorderType = SS_BORDER_TYPE_TOP
            ssPrRoutine.Action = ActionSetCellBorder
            ssPrRoutine.BlockMode = False
            
            ssPrRoutine.Col = 1: ssPrRoutine.Text = adoSet.Fields("RoutinCd").Value & ""
            ssPrRoutine.Col = 2: ssPrRoutine.Text = adoSet.Fields("RoutinNM").Value & ""
        End If
        ssPrRoutine.Col = 3: ssPrRoutine.Text = adoSet.Fields("Codeky").Value & ""
        ssPrRoutine.Col = 4: ssPrRoutine.Text = adoSet.Fields("ItemNM").Value & ""
        ssPrRoutine.Col = 5: ssPrRoutine.Text = adoSet.Fields("SugaCd").Value & ""
        ssPrRoutine.Col = 6: ssPrRoutine.Text = adoSet.Fields("YakCd").Value & ""
        
        sTempText = adoSet.Fields("RoutinCd").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub cmdPrintOk_Click()
    
    If ssPrRoutine.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("RoutineCode  Data 의 Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & Trim(txtSLipTitle.Text) & " RoutineCode Data "
    'strHead(1) = "/f2" & "Page : " & "/p" & " of " & ssPrRoutine.PrintPageCount & "/r"
    
    ssPrRoutine.PrintHeader = strFont(0) + strHead(0) + strFont(1) + strHead(1) + "/n" + strFont(1)
    ssPrRoutine.PrintFooter = "/f2" & "/c" & "Page : " & "/p" & " of " & ssPrRoutine.PrintPageCount - 1
    ssPrRoutine.PrintMarginLeft = 300
    ssPrRoutine.PrintMarginRight = 100
    ssPrRoutine.PrintMarginTop = 200
    ssPrRoutine.PrintMarginBottom = 300
    ssPrRoutine.PrintColHeaders = True
    ssPrRoutine.PrintRowHeaders = True
    ssPrRoutine.PrintBorder = True
    ssPrRoutine.PrintColor = False
    ssPrRoutine.PrintGrid = False
    ssPrRoutine.PrintShadows = True
    ssPrRoutine.PrintUseDataMax = False
    ssPrRoutine.Row = 1
    ssPrRoutine.Row2 = ssPrRoutine.DataRowCnt
    ssPrRoutine.Col = 1
    ssPrRoutine.Col2 = ssPrRoutine.MaxCols
    ssPrRoutine.PrintType = PrintTypeCellRange
    ssPrRoutine.PrintOrientation = PrintOrientationPortrait
    ssPrRoutine.Action = ActionPrint

End Sub

Private Sub CommandButton1_Click()
    Dim sSLip       As String
    
        
    sSLip = InputBox("조회할SLipNo를 지정하십시오!.....", "SLip조회")
    sSLip = Trim(sSLip)
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_iTemML"
    strSql = strSql & " WHERE  Codeky LIKE '" & sSLip & "%'"
    strSql = strSql & " ORDER  BY Codeky"
    
    ssItem.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssItem.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssItem.Row = ssItem.DataRowCnt + 1
        ssItem.Col = 1: ssItem.Text = Trim(adoSet.Fields("Codeky").Value & "")
        ssItem.Col = 2: ssItem.Text = RTrim(adoSet.Fields("Itemnm").Value & "")
        ssItem.Col = 3: ssItem.Text = Trim(adoSet.Fields("Sugacd").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub Form_Load()
    
    
    GoSub Get_Data_Specode12
    GoSub Get_Data_Specode21
    
    Exit Sub
    
Get_Data_Specode12:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " Order  By  Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSlip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return
    
Get_Data_Specode21:
    Dim sJang       As String * 4
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu = '21'"
    strSql = strSql & " ORDER  BY Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sJang = Trim(adoSet.Fields("Codeky").Value & "")
        cmbJangbi.AddItem sJang & ". " & Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub mnuQuit_Click()

    mdiMain.stbMain.Panels(1).Text = ""
    Unload Me
    
End Sub

Private Sub ssItem_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sMove(1 To 3)    As String
    
    If Row = 0 Then Exit Sub
    If Trim(txtSlipKey.Text) = "" Then Exit Sub
    If Trim(txtSlipCode.Text) = "" Then Exit Sub
    
    
    ssItem.Row = Row
    ssItem.Col = 1: sMove(1) = ssItem.Text
    ssItem.Col = 2: sMove(2) = ssItem.Text
    ssItem.Col = 3: sMove(3) = ssItem.Text
    GoSub LEFT_CHECK
    
    ssRoutine.Row = ssRoutine.DataRowCnt + 1
    ssRoutine.Col = 2: ssRoutine.Text = sMove(1)
    ssRoutine.Col = 3: ssRoutine.Text = sMove(2)
    ssRoutine.Col = 4: ssRoutine.Text = sMove(3)
    ssRoutine.Action = ActionActiveCell
    Exit Sub
    
LEFT_CHECK:
    For i = 1 To ssRoutine.DataRowCnt
        ssRoutine.Row = i
        ssRoutine.Col = 2
        If Trim(sMove(1)) = Trim(ssRoutine.Text) Then
            MsgBox "이미 같은 코드가 등록되어 있습니다!.", vbCritical
            Exit Sub
        End If
    Next
    Return
    
End Sub

Private Sub ssRoutine_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If Col = 5 Then
            mdiMain.stbMain.Panels(1).Text = "연속검사 여부Check Column"
        End If
    End If
    
End Sub

Private Sub txtSlipCode_LostFocus()
    Dim sRoutine        As String
    
    If Trim(txtSlipKey.Text) = "" Then Exit Sub
    If Trim(txtSlipCode.Text) = "" Then Exit Sub
    
    DoEvents
    sRoutine = Trim(txtSlipKey.Text) & Trim(txtSlipCode.Text)
    
    strSql = ""
    strSql = strSql & " SELECT a.*"
    strSql = strSql & " FROM   TWEXAM_ROUTINE a "
    strSql = strSql & " WHERE  a.RoutinCD  = '" & sRoutine & "'"
    strSql = strSql & " ORDER  BY a.Codeky    "
   
    GoSub Spread_Clear
    If False = adoSetOpen(strSql, adoSet) Then
        txtRoutineName.Text = ""
        txtYageo.Text = ""
        txtOrderCD.Text = ""
        cmbJangbi.ListIndex = -1
        dtCodate.Value = Dual_Date_Get("yyyy-MM-dd")
        Exit Sub
    End If
    
    
    ssRoutine.Row = 1
    ssRoutine.Col = 1
    ssRoutine.Action = SS_ACTION_ACTIVE_CELL
    
    Do Until adoSet.EOF
        ssRoutine.Row = ssRoutine.DataRowCnt + 1
        ssRoutine.Col = 2: ssRoutine.Text = adoSet.Fields("Codeky").Value & ""
        ssRoutine.Col = 3: ssRoutine.Text = adoSet.Fields("ItemNM").Value & ""
        ssRoutine.Col = 4: ssRoutine.Text = adoSet.Fields("SugaCD").Value & ""
        If adoSet.Fields("Series").Value & "" = "1" Then
            ssRoutine.Col = 5: ssRoutine.Value = True
        Else
        End If
        txtRoutineName.Text = Trim(adoSet.Fields("RoutinNM").Value & "")
        txtYageo.Text = Trim(adoSet.Fields("YakCD").Value & "")
        txtOrderCD.Text = Trim(adoSet.Fields("OrderCD").Value & "")
        
        For i = 0 To cmbJangbi.ListCount - 1
            If Left(cmbJangbi.List(i), 4) = Left(adoSet.Fields("Jangbi").Value & "", 4) Then
                cmbJangbi.ListIndex = i
                Exit For
            End If
        Next
        
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    Exit Sub
    
Spread_Clear:
    ssRoutine.Row = 1
    ssRoutine.Row2 = ssRoutine.DataRowCnt
    ssRoutine.Col = 1
    ssRoutine.Col2 = ssRoutine.DataColCnt
    ssRoutine.BlockMode = True
    ssRoutine.Action = SS_ACTION_CLEAR_TEXT
    ssRoutine.BlockMode = False
    Return
    
End Sub

