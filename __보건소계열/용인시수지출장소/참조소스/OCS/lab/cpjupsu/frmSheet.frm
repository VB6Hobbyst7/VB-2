VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSheet 
   Caption         =   "정규Order WorkSheet"
   ClientHeight    =   7740
   ClientLeft      =   150
   ClientTop       =   1110
   ClientWidth     =   11625
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11625
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   315
      TabIndex        =   1
      Top             =   90
      Width           =   10815
      _Version        =   65536
      _ExtentX        =   19076
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
      Begin MSComCtl2.DTPicker dtJeobsuDt 
         Height          =   330
         Left            =   2790
         TabIndex        =   7
         Top             =   45
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36524
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   330
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "정규Order WorkSheet 출력"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton optGbch 
         Caption         =   "확인Order"
         Height          =   180
         Index           =   0
         Left            =   4545
         TabIndex        =   3
         Tag             =   "Y"
         Top             =   135
         Width           =   1185
      End
      Begin VB.OptionButton optGbch 
         Caption         =   "미확인Order"
         Height          =   180
         Index           =   1
         Left            =   5760
         TabIndex        =   2
         Tag             =   "2"
         Top             =   135
         Value           =   -1  'True
         Width           =   1320
      End
      Begin MSForms.CommandButton cmdPr 
         Height          =   510
         Left            =   8955
         TabIndex        =   5
         Top             =   90
         Width           =   1680
         Caption         =   "출력"
         PicturePosition =   327683
         Size            =   "2963;900"
         Picture         =   "frmSheet.frx":0000
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQryOK 
         Height          =   510
         Left            =   7380
         TabIndex        =   4
         Top             =   90
         Width           =   1590
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2805;900"
         Picture         =   "frmSheet.frx":08DA
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread sprSheet 
      Height          =   6405
      Left            =   315
      TabIndex        =   0
      Top             =   1035
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "frmSheet.frx":11BC
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPr_Click()
    Dim strFont1        As String
    Dim strFont2        As String
    Dim strHead1        As String
    Dim strHead2        As String
    Dim strHead3        As String
    Dim iThisPage       As Integer
    Dim sFooter         As String
    Dim sPortBar        As String
    
    If sprSheet.DataRowCnt < 1 Then
        MsgBox "Printing 할 Data 가 없습니다!.확인하세요 ,,,,"
        Exit Sub
    End If
    
    sPortBar = ""
    
    For i = 1 To 110
        sPortBar = sPortBar & "━"
    Next
    
    strFont1 = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont2 = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs0"
    strHead1 = "/f1" & "/c" & "정규채혈 OrderList"
    strHead2 = "/f2" & "/l" & "Date :" & Dual_Date_Get("yyyy-MM-dd")
    
    sprSheet.PrintHeader = strFont1 + strHead1 + "/n/n" + strFont2 + strHead2 + "/n" + _
                           strFont2 + "/l" + sPortBar
    
    sprSheet.PrintFooter = strFont2 + "/l" + sPortBar & "/n" & _
                           Space(110) & "출력일자: " & Format(Dual_Date_Get("yyyy-MM-dd"), "yyyy-MM-dd aaaa")

    sprSheet.PrintMarginLeft = 0
    sprSheet.PrintMarginRight = 0
    sprSheet.PrintMarginTop = 100
    sprSheet.PrintMarginBottom = 100
    sprSheet.PrintColHeaders = True
    sprSheet.PrintRowHeaders = True
    sprSheet.PrintBorder = False
    sprSheet.PrintColor = False
    sprSheet.PrintGrid = True
    sprSheet.PrintShadows = True
    sprSheet.PrintUseDataMax = False
    sprSheet.Row = 1
    sprSheet.Col = 1
    sprSheet.Row2 = sprSheet.DataRowCnt
    sprSheet.Col2 = sprSheet.MaxCols
    sprSheet.PrintType = SS_PRINT_CELL_RANGE
    sprSheet.PrintOrientation = PrintOrientationLandscape
    sprSheet.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdQryOK_Click()
    Dim sFrJeobsuDt         As String
    Dim sToJeobsuDt         As String
    Dim strJeobsuDt         As String
    Dim sComWard            As String
    Dim sComRoom            As String
    Dim sComSLipno1         As String
    Dim sComDeptDr          As String
    Dim sComPtno            As String
    
    
    'strJeobsuDt = Format(frmIPDMain.dtJeobsuDt.Value, "yyyy-MM-dd")
    
    strJeobsuDt = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
    
    DoEvents: Screen.MousePointer = vbHourglass
    Call Spread_Set_Clear(sprSheet)
    GoSub Get_Order_MainProcess
    DoEvents: Screen.MousePointer = vbDefault
    
    Exit Sub
    

Get_Order_MainProcess:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */  "
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID OrderRowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt1,"
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate1,  "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt1, "
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate1,"
    strSql = strSql & "        a.DeptCode DeptCode1, a.SLipno1 SLno,"
    strSql = strSql & "        a.Ptno Ptno1,"
    strSql = strSql & "        b.Sname, c.Codenm SLname,"
    strSql = strSql & "        d.Codenm Samplename, e.Drname, g.WardCode, a.RoomCode RoomCode1,"
    strSql = strSql & "        f.ITemnm ItemNM, 'i' RoutineGb"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML  f, "
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     g "
    strSql = strSql & " WHERE  a.JeobsuDt    = TO_DATE('" & strJeobsuDt & "','YYYY-MM-DD')    "
'C    strSql = strSql & " AND    a.SLipno1   < 52"
    strSql = strSql & " AND    a.SLipno1   < 90 "
    strSql = strSql & " AND    a.Gbio      = 'I'"         '입원환자  만 ...
    strSql = strSql & " AND    a.OrderGB  IN ('X','Y','Z',' ')"  '정규Order 만...

    If optGbch(0).Value = True Then
        strSql = strSql & " AND    a.GBCh      = '" & optGbch(0).Tag & "'"
    End If
    If optGbch(1).Value = True Then
        strSql = strSql & " AND    a.GbCh      = '" & optGbch(1).Tag & "'"
    End If
    
    strSql = strSql & " AND    a.JeobsuYn   = '*'"
    'strsql = strsql & " AND    a.EntTime   = 1
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)"
    strSql = strSql & " AND    a.ItemCd    = f.Codeky"
    
    If frmIPDMain.cmbWard.ListIndex > -1 Then
        strSql = strSql & " AND    g.WardCode  = '" & Left(frmIPDMain.cmbWard.Text, 4) & "'"
    End If
    
    strSql = strSql & " AND    a.RoomCode  = g.RoomCode(+)"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " UNION ALL    "
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */  "
    strSql = strSql & " SELECT DISTINCT a.*, a.RowID OrderRowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt1,"
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate1,  "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt1, "
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate1,"
    strSql = strSql & "        a.DeptCode DeptCode1, a.SLipno1 SLno,"
    strSql = strSql & "        a.Ptno Ptno1,"
    strSql = strSql & "        b.Sname, c.Codenm SLname,"
    strSql = strSql & "        d.Codenm Samplename, e.Drname, g.WardCode, a.RoomCode RoomCode1,"
    strSql = strSql & "        f.RoutinNM ItemNM, 'r' RoutineGb"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine f, "
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     g  "
    strSql = strSql & " WHERE  a.JeobsuDt    = TO_DATE('" & strJeobsuDt & "','YYYY-MM-DD')    "
'C    strSql = strSql & " AND    a.SLipno1   < 52"
    strSql = strSql & " AND    a.SLipno1   < 90 "
    strSql = strSql & " AND    a.Gbio      = 'I'"
    strSql = strSql & " AND    a.OrderGB  IN ('X','Y','Z',' ')"  '정규Order 만...
    If optGbch(0).Value = True Then
        strSql = strSql & " AND    a.GBCh      = '" & optGbch(0).Tag & "'"
    End If
    If optGbch(1).Value = True Then
        strSql = strSql & " AND    a.GbCh      = '" & optGbch(1).Tag & "'"
    End If
    

    strSql = strSql & " AND    a.JeobsuYn   = '*'"
    'strsql = strsql & " AND    a.EntTime   = 1
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)"
    strSql = strSql & " AND    a.ItemCd    = f.RoutinCD"
    
    If frmIPDMain.cmbWard.ListIndex > -1 Then
        strSql = strSql & " AND    g.WardCode  = '" & Left(frmIPDMain.cmbWard.Text, 4) & "'"
    End If
    
    strSql = strSql & " AND    a.RoomCode  = g.RoomCode(+)"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " ORDER  BY  RoomCode1, Ptno1, SLno, Jeobsudt1, DeptCode1"
    
    sprSheet.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    sprSheet.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        sprSheet.Row = sprSheet.DataRowCnt + 1
        
        If sComWard <> adoSet.Fields("WardCode").Value & "" Then
            sprSheet.Col = 1:  sprSheet.Text = adoSet.Fields("WardCode").Value & ""
            
            sprSheet.Col = 1:  sprSheet.Col2 = sprSheet.MaxCols
            sprSheet.Row = sprSheet.Row:  sprSheet.Row2 = sprSheet.Row
            sprSheet.BlockMode = True
            sprSheet.CellBorderType = SS_BORDER_TYPE_TOP
            sprSheet.CellBorderStyle = CellBorderStyleSolid
            sprSheet.Action = ActionSetCellBorder
            sprSheet.BlockMode = False
        
        End If
        
        If sComRoom <> adoSet.Fields("RoomCode").Value & "" Then
            sprSheet.Col = 2:  sprSheet.Text = adoSet.Fields("RoomCode").Value & ""
            
            sprSheet.Col = 2:  sprSheet.Col2 = sprSheet.MaxCols
            sprSheet.Row = sprSheet.Row:  sprSheet.Row2 = sprSheet.Row
            sprSheet.BlockMode = True
            sprSheet.CellBorderType = SS_BORDER_TYPE_TOP
            sprSheet.CellBorderStyle = CellBorderStyleSolid
            sprSheet.Action = ActionSetCellBorder
            sprSheet.BlockMode = False
        End If
        
        
        If sComPtno <> adoSet.Fields("Ptno").Value & "" Then
            sprSheet.Col = 3:  sprSheet.Text = adoSet.Fields("Ptno").Value & ""
            sprSheet.Col = 4:  sprSheet.Text = adoSet.Fields("Sname").Value & ""
            sprSheet.Col = 5:  sprSheet.Text = adoSet.Fields("Sex").Value & ""
            sprSheet.Col = 6:  sprSheet.Text = adoSet.Fields("AgeYY").Value & ""
            sprSheet.Col = 7:  sprSheet.Text = adoSet.Fields("SLno").Value & ""
            sprSheet.Col = 8:  sprSheet.Text = adoSet.Fields("Samplename").Value & ""
            sprSheet.Col = 10: sprSheet.Text = adoSet.Fields("DeptCode1").Value & ""
            sprSheet.Col = 11: sprSheet.Text = adoSet.Fields("Drname").Value & ""
            
            sprSheet.Col = 3:  sprSheet.Col2 = sprSheet.MaxCols
            sprSheet.Row = sprSheet.Row:   sprSheet.Row2 = sprSheet.Row
            sprSheet.BlockMode = True
            sprSheet.CellBorderType = SS_BORDER_TYPE_TOP
            sprSheet.CellBorderColor = RGB(192, 192, 192)
            sprSheet.CellBorderStyle = CellBorderStyleSolid
            
            sprSheet.Action = ActionSetCellBorder
            sprSheet.BlockMode = False
            
        End If
        
        If sComSLipno1 <> adoSet.Fields("SLipno1").Value & "" & adoSet.Fields("SampleName").Value & "" Then
            sprSheet.Col = 7:  sprSheet.Text = adoSet.Fields("SLno").Value & ""
            sprSheet.Col = 8:  sprSheet.Text = adoSet.Fields("Samplename").Value & ""
        End If
        
        
        sprSheet.Col = 9:  sprSheet.Text = adoSet.Fields("ItemNM").Value & ""
        
        sprSheet.Col = 12: sprSheet.Text = adoSet.Fields("CmDoctor").Value & ""
        sprSheet.Col = 13: sprSheet.Text = adoSet.Fields("GbER").Value & ""
        
        
        sComWard = adoSet.Fields("WardCode").Value & ""
        sComRoom = adoSet.Fields("RoomCode").Value & ""
        sComPtno = adoSet.Fields("Ptno").Value & ""
        sComSLipno1 = adoSet.Fields("SLipno1").Value & "" & adoSet.Fields("SampleName").Value & ""

        
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    Return
    
    
Spread_sprSheet_Clear:
    sprSheet.ReDraw = False
    sprSheet.MaxRows = 0
    sprSheet.MaxRows = 100
    sprSheet.RowHeight(-1) = 11.5
    sprSheet.ReDraw = True
    
    
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then Me.Controls(i).Text = ""
    Next
    
    
    
    Return

End Sub

Private Sub Form_Load()
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    dtJeobsuDt.Value = Format(frmIPDMain.dtJeobsuDt.Value, "yyyy-MM-dd")
    
    Call cmdQryOK_Click
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
