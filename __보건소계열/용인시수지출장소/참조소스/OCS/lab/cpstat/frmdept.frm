VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmDept 
   Caption         =   "진료과별 통계"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   1110
   ClientWidth     =   11745
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
   ScaleHeight     =   6615
   ScaleWidth      =   11745
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   9465
      _Version        =   65536
      _ExtentX        =   16695
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
      Begin VB.OptionButton Option3 
         Caption         =   "입원"
         Height          =   330
         Left            =   7155
         TabIndex        =   12
         Top             =   135
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Caption         =   "외래"
         Height          =   330
         Left            =   6300
         TabIndex        =   11
         Top             =   135
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "외래.입원"
         Height          =   330
         Left            =   4995
         TabIndex        =   10
         Top             =   135
         Value           =   -1  'True
         Width           =   1140
      End
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
   Begin Threed.SSPanel panelo 
      Height          =   7440
      Left            =   45
      TabIndex        =   7
      Top             =   1215
      Visible         =   0   'False
      Width           =   11850
      _Version        =   65536
      _ExtentX        =   20902
      _ExtentY        =   13123
      _StockProps     =   15
      Caption         =   "외래Data"
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
      Alignment       =   0
      Begin FPSpreadADO.fpSpread sprDepto 
         Height          =   6630
         Left            =   360
         TabIndex        =   16
         Top             =   630
         Width           =   11175
         _Version        =   196608
         _ExtentX        =   19711
         _ExtentY        =   11695
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
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
         MaxCols         =   100
         MaxRows         =   497
         SpreadDesigner  =   "frmDept.frx":0000
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdPro 
         Height          =   465
         Left            =   9990
         TabIndex        =   18
         Top             =   135
         Width           =   1500
         Caption         =   "출력확인"
         PicturePosition =   327683
         Size            =   "2646;820"
         Picture         =   "frmDept.frx":2E42
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQueryo 
         Height          =   465
         Left            =   8505
         TabIndex        =   17
         Top             =   135
         Width           =   1455
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2566;820"
         Picture         =   "frmDept.frx":371C
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel paneli 
      Height          =   7260
      Left            =   45
      TabIndex        =   6
      Top             =   1035
      Visible         =   0   'False
      Width           =   11850
      _Version        =   65536
      _ExtentX        =   20902
      _ExtentY        =   12806
      _StockProps     =   15
      Caption         =   "입원Data"
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
      Alignment       =   0
      Begin FPSpreadADO.fpSpread sprDepti 
         Height          =   6630
         Left            =   270
         TabIndex        =   13
         Top             =   585
         Width           =   11175
         _Version        =   196608
         _ExtentX        =   19711
         _ExtentY        =   11695
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
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
         MaxCols         =   100
         MaxRows         =   497
         SpreadDesigner  =   "frmDept.frx":3FF6
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdQueryi 
         Height          =   465
         Left            =   8505
         TabIndex        =   15
         Top             =   90
         Width           =   1455
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2566;820"
         Picture         =   "frmDept.frx":6E38
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPri 
         Height          =   465
         Left            =   9990
         TabIndex        =   14
         Top             =   90
         Width           =   1500
         Caption         =   "출력확인"
         PicturePosition =   327683
         Size            =   "2646;820"
         Picture         =   "frmDept.frx":7712
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel panelio 
      Height          =   7440
      Left            =   45
      TabIndex        =   4
      Top             =   810
      Width           =   11850
      _Version        =   65536
      _ExtentX        =   20902
      _ExtentY        =   13123
      _StockProps     =   15
      Caption         =   "외래.입원"
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
      Alignment       =   0
      Begin FPSpreadADO.fpSpread sprDept 
         Height          =   6630
         Left            =   405
         TabIndex        =   5
         Top             =   630
         Width           =   11175
         _Version        =   196608
         _ExtentX        =   19711
         _ExtentY        =   11695
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
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
         MaxCols         =   100
         MaxRows         =   498
         SpreadDesigner  =   "frmDept.frx":7FEC
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   465
         Left            =   8505
         TabIndex        =   9
         Top             =   135
         Width           =   1455
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2566;820"
         Picture         =   "frmDept.frx":B52F
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPr 
         Height          =   465
         Left            =   9990
         TabIndex        =   8
         Top             =   135
         Width           =   1500
         Caption         =   "출력확인"
         PicturePosition =   327683
         Size            =   "2646;820"
         Picture         =   "frmDept.frx":BE09
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
Attribute VB_Name = "frmDept"
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
    
    
    If sprDept.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 진료과별 통계(입.외)"
    strHead3 = "/f3" & "기간 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & ""
    
    sprDept.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                       strFont2 + "/l" + strHead3 + _
                                       strFont2 + strHead5
    sprDept.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(110) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprDept.PrintMarginLeft = 0
    sprDept.PrintMarginRight = 0
    sprDept.PrintMarginTop = 0
    sprDept.PrintMarginBottom = 0
    sprDept.PrintColHeaders = True
    sprDept.PrintRowHeaders = False
    sprDept.PrintBorder = True
    sprDept.PrintColor = True
    sprDept.PrintGrid = True
    sprDept.PrintShadows = True
    sprDept.PrintUseDataMax = False
    
    sprDept.Row = 1: sprDept.Row2 = sprDept.DataRowCnt
    sprDept.Col = 2: sprDept.Col2 = sprDept.DataColCnt
    sprDept.PrintType = SS_PRINT_CELL_RANGE
    sprDept.PrintOrientation = PrintOrientationLandscape
    sprDept.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdPri_Click()
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
    
    
    If sprDepti.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 진료과별 통계(입원)"
    strHead3 = "/f3" & "기간 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & ""
    
    sprDepti.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                       strFont2 + "/l" + strHead3 + _
                                       strFont2 + strHead5
    sprDepti.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(110) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprDepti.PrintMarginLeft = 0
    sprDepti.PrintMarginRight = 0
    sprDepti.PrintMarginTop = 0
    sprDepti.PrintMarginBottom = 0
    sprDepti.PrintColHeaders = True
    sprDepti.PrintRowHeaders = False
    sprDepti.PrintBorder = True
    sprDepti.PrintColor = True
    sprDepti.PrintGrid = True
    sprDepti.PrintShadows = True
    sprDepti.PrintUseDataMax = False
    
    sprDepti.Row = 1: sprDepti.Row2 = sprDepti.DataRowCnt
    sprDepti.Col = 2: sprDepti.Col2 = sprDepti.DataColCnt
    sprDepti.PrintType = SS_PRINT_CELL_RANGE
    sprDepti.PrintOrientation = PrintOrientationLandscape
    sprDepti.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdPro_Click()
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
    
    
    If sprDepto.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "임상병리과 진료과별 통계(외래)"
    strHead3 = "/f3" & "기간 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & ""
    
    sprDepto.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                       strFont2 + "/l" + strHead3 + _
                                       strFont2 + strHead5
    sprDepto.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(110) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprDepto.PrintMarginLeft = 0
    sprDepto.PrintMarginRight = 0
    sprDepto.PrintMarginTop = 0
    sprDepto.PrintMarginBottom = 0
    sprDepto.PrintColHeaders = True
    sprDepto.PrintRowHeaders = False
    sprDepto.PrintBorder = True
    sprDepto.PrintColor = True
    sprDepto.PrintGrid = True
    sprDepto.PrintShadows = True
    sprDepto.PrintUseDataMax = False
    
    sprDepto.Row = 1: sprDepto.Row2 = sprDepto.DataRowCnt
    sprDepto.Col = 2: sprDepto.Col2 = sprDepto.DataColCnt
    sprDepto.PrintType = SS_PRINT_CELL_RANGE
    sprDepto.PrintOrientation = PrintOrientationLandscape
    sprDepto.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdQuery_Click()
    Dim sDept()     As String
    Dim sFrDate     As String
    Dim sToDate     As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call SpreadSetClear(sprDept)
    
    GoSub Get_DeptCode_Array
    GoSub Get_Main_Proc
    GoSub Total_Calculate
    Exit Sub
    
    
    
Get_DeptCode_Array:
    StrSql = ""
    StrSql = StrSql & " SELECT DeptCode"
    StrSql = StrSql & " FROM   TWEXAM_Order"
    StrSql = StrSql & " WHERE  COLLDate  >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    COLLDate  <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    SLipno1   >  0"
    StrSql = StrSql & " AND    JeobsuYN   = '*'"
    StrSql = StrSql & " AND    SLipno1   <  52"
    StrSql = StrSql & " GROUP  BY DeptCode"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    If adoSet.RecordCount = 0 Then
        Exit Sub
    End If
    
    ReDim sDept(adoSet.RecordCount - 1)
    i = 0: iCol = 3
    Do Until adoSet.EOF
        sprDept.Row = iRow
        sDept(i) = Trim(adoSet.Fields("DeptCode").Value) & ""
        
        adoSet.MoveNext: i = i + 1: iCol = iCol + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

Get_Main_Proc:
    StrSql = ""
    StrSql = StrSql & " SELECT itemCd, itemName,"
    
    For i = 0 To UBound(sDept)
        StrSql = StrSql & "    SUM(Decode(DeptCode, '" & RTrim(sDept(i)) & "_o',1, '')) " & Trim(sDept(i)) & "_o,"
        StrSql = StrSql & "    SUM(Decode(DeptCode, '" & RTrim(sDept(i)) & "_i',1, '')) " & Trim(sDept(i)) & "_i,"
    Next
        
    StrSql = StrSql & "        COUNT(Decode(DeptCode, '99','',1)) LineTotal"
    
    StrSql = StrSql & " FROM(  SELECT  DISTINCT a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, "
    StrSql = StrSql & "                a.ItemCd, b.RoutinNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode) || '_o' Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_Routine     b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'O'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.RoutinCd "
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 "

  '  strSql = strSql & "        GROUP BY ItemCd, b.RoutinNm, a.GBIO, a.DeptCode, c.DeptNamek"
    StrSql = StrSql & "     UNION ALL"
    StrSql = StrSql & "        SELECT  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2,"
    StrSql = StrSql & "                a.ItemCd, b.ITEMNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode) || '_o' Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_iTEMML      b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'O'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.Codeky"
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 "
    
    StrSql = StrSql & "     UNION ALL"
    StrSql = StrSql & "        SELECT  DISTINCT a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, "
    StrSql = StrSql & "                a.ItemCd, b.RoutinNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode) || '_i' Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_Routine     b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'I'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.RoutinCd "
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 "
    
   ' strSql = strSql & "        GROUP BY ItemCd, b.RoutinNm, a.GBIO, a.DeptCode, c.DeptNamek"
    StrSql = StrSql & "     UNION ALL"
    StrSql = StrSql & "        SELECT  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, "
    StrSql = StrSql & "                a.ItemCd, b.ITEMNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode) || '_i' Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_iTEMML      b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'I'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.Codeky"
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52)"
    
    'strSql = strSql & "        GROUP BY ItemCd, b.ITEMNM, a.GBIO, a.DeptCode, c.DeptNamek) "
    
    StrSql = StrSql & " GROUP  BY itemCd, itemName"
    
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Dim sDepto  As String
    Dim sDepti  As String
    
    sprDept.MaxCols = 100
    sprDept.ColWidth(-1) = 5
    sprDept.ColWidth(1) = 7.38
    sprDept.ColWidth(2) = 19.38
    
    For i = 3 To sprDept.MaxCols - 1
        sprDept.Row = 1
        If i Mod 2 = 1 Then
            sprDept.Col = i: sprDept.Text = "o"
        Else
            sprDept.Col = i: sprDept.Text = "i"
        End If
    Next
    
    sprDept.ReDraw = False
    Do Until adoSet.EOF
        sprDept.Row = sprDept.DataRowCnt + 1
        sprDept.Col = 1: sprDept.Text = adoSet.Fields("ItemCd").Value & ""
        sprDept.Col = 2: sprDept.Text = adoSet.Fields("ItemName").Value & ""
        
        iCol = 3
        For i = 0 To UBound(sDept)
            sDepto = Trim(sDept(i)) & "_o"
            sDepti = Trim(sDept(i)) & "_i"
            
            sprDept.Col = iCol: sprDept.Text = adoSet.Fields(sDepto).Value & ""
            sprDept.Row = 0: sprDept.Text = sDept(i)
            iCol = iCol + 1
            sprDept.Row = sprDept.DataRowCnt
            
            sprDept.Col = iCol: sprDept.Text = adoSet.Fields(sDepti).Value & ""
            sprDept.Row = 0: sprDept.Text = " "
            sprDept.Row = sprDept.DataRowCnt
            iCol = iCol + 1
        Next
        
        sprDept.Col = (UBound(sDept) + 3) * 2 - 1: sprDept.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    sprDept.ReDraw = True
    
    sprDept.MaxCols = (UBound(sDept) + 3) * 2 - 1
    sprDept.Col = sprDept.MaxCols
    sprDept.Row = 0: sprDept.Text = "합계"
    sprDept.Row = 1: sprDept.Text = ""
    
    Call adoSetClose(adoSet)
    
    Return
    
Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    iLastSprRow = sprDept.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprDept.DataColCnt
        nCalSum = 0
        For i = 2 To iLastSprRow
            sprDept.Row = i
            sprDept.Col = j
            If Trim(sprDept.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprDept.Text)
            End If
        Next
        sprDept.Row = iLastSprRow + 2
        sprDept.Col = j
        sprDept.Text = nCalSum
    Next
    
    sprDept.Row = sprDept.DataRowCnt
    sprDept.Col = 2: sprDept.Text = "합계"
    Return
    
    
End Sub



Private Sub cmdQueryi_Click()
    Dim sDept()     As String
    Dim sFrDate     As String
    Dim sToDate     As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call SpreadSetClear(sprDepti)
    
    GoSub Get_DeptCode_Array
    GoSub Get_Main_Proc
    GoSub Total_Calculate
    Exit Sub
    
    
    
Get_DeptCode_Array:
    StrSql = ""
    StrSql = StrSql & " SELECT DeptCode"
    StrSql = StrSql & " FROM   TWEXAM_Order"
    StrSql = StrSql & " WHERE  COLLDate  >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    COLLDate  <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    SLipno1   >  0"
    StrSql = StrSql & " AND    JeobsuYN  = '*'"
    StrSql = StrSql & " AND    SLipno1   <  52"
    StrSql = StrSql & " AND    GBIO       = 'I'"
    StrSql = StrSql & " GROUP  BY DeptCode"
    
    If False = adoSetOpen(StrSql, adoSet) Then
        MsgBox "Data 가 없습니다!..........", vbCritical
        Exit Sub
    End If
    
    If adoSet.RecordCount = 0 Then
        MsgBox "Data 가 없습니다!..........", vbCritical
        Exit Sub
    End If
    
    ReDim sDept(adoSet.RecordCount - 1)
    i = 0
    Do Until adoSet.EOF
        sDept(i) = Trim(adoSet.Fields("DeptCode").Value) & ""
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

Get_Main_Proc:
    StrSql = ""
    StrSql = StrSql & " SELECT itemCd, itemName,"
    
    For i = 0 To UBound(sDept)
        StrSql = StrSql & "    SUM(Decode(DeptCode, '" & RTrim(sDept(i)) & "',1, '')) " & Trim(sDept(i)) & ","
    Next
        
    StrSql = StrSql & "        COUNT(Decode(DeptCode, '99','',1)) LineTotal"
    
    StrSql = StrSql & " FROM(  SELECT  DISTINCT a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, "
    StrSql = StrSql & "                a.ItemCd, b.RoutinNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode)  Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_Routine     b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'I'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.RoutinCd "
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 "
    StrSql = StrSql & "     UNION ALL"
    StrSql = StrSql & "        SELECT  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2,"
    StrSql = StrSql & "                a.ItemCd, b.ITEMNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode)  Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_iTEMML      b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'I'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.Codeky"
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 )"
    'strSql = strSql & "        GROUP BY ItemCd, b.ITEMNM, a.GBIO, a.DeptCode, c.DeptNamek) "
    
    StrSql = StrSql & " GROUP  BY itemCd, itemName"
    
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    sprDepti.MaxCols = 100
    sprDepti.ColWidth(-1) = 5
    sprDepti.ColWidth(1) = 7.38
    sprDepti.ColWidth(2) = 19.38
    
    
    sprDepti.ReDraw = False
    Do Until adoSet.EOF
        sprDepti.Row = sprDepti.DataRowCnt + 1
        sprDepti.Col = 1: sprDepti.Text = adoSet.Fields("ItemCd").Value & ""
        sprDepti.Col = 2: sprDepti.Text = adoSet.Fields("ItemName").Value & ""
        
        iCol = 3
        For i = 0 To UBound(sDept)
            sprDepti.Col = i + 3: sprDepti.Text = adoSet.Fields(sDept(i)).Value & ""
            sprDepti.Row = 0: sprDepti.Text = sDept(i)
            iCol = iCol + 1
            sprDepti.Row = sprDepti.DataRowCnt
        Next
        sprDepti.Col = UBound(sDept) + 4: sprDepti.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    sprDepti.ReDraw = True
    
    sprDepti.MaxCols = UBound(sDept) + 4
    sprDepti.Col = sprDepti.MaxCols
    sprDepti.Row = 0: sprDepti.Text = "합계"
    
    
    Call adoSetClose(adoSet)
    
    Return
    
Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    iLastSprRow = sprDepti.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprDepti.DataColCnt
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprDepti.Row = i
            sprDepti.Col = j
            If Trim(sprDepti.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprDepti.Text)
            End If
        Next
        sprDepti.Row = iLastSprRow + 2
        sprDepti.Col = j
        sprDepti.Text = nCalSum
    Next
    
    sprDepti.Row = sprDepti.DataRowCnt
    sprDepti.Col = 2: sprDepti.Text = "합계"
    Return

End Sub

Private Sub cmdQueryo_Click()
    Dim sDept()     As String
    Dim sFrDate     As String
    Dim sToDate     As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call SpreadSetClear(sprDepto)
    
    GoSub Get_DeptCode_Array
    GoSub Get_Main_Proc
    GoSub Total_Calculate
    Exit Sub
    
    
    
Get_DeptCode_Array:
    StrSql = ""
    StrSql = StrSql & " SELECT DeptCode"
    StrSql = StrSql & " FROM   TWEXAM_Order"
    StrSql = StrSql & " WHERE  COLLDate  >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    COLLDate  <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    SLipno1   >  0"
    StrSql = StrSql & " AND    SLipno1   <  52"
    StrSql = StrSql & " AND    GBIO       = 'O'"
    StrSql = StrSql & " AND    JeobsuYN   = '*'"
    StrSql = StrSql & " GROUP  BY DeptCode"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
        
    
    ReDim sDept(adoSet.RecordCount - 1)
    i = 0
    Do Until adoSet.EOF
        sDept(i) = Trim(adoSet.Fields("DeptCode").Value) & ""
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

Get_Main_Proc:
    StrSql = ""
    StrSql = StrSql & " SELECT itemCd, itemName,"
    
    For i = 0 To UBound(sDept)
        StrSql = StrSql & "    SUM(Decode(DeptCode, '" & RTrim(sDept(i)) & "',1, '')) " & Trim(sDept(i)) & ","
    Next
        
    StrSql = StrSql & "        COUNT(Decode(DeptCode, '99','',1)) LineTotal"
    
    StrSql = StrSql & " FROM(  SELECT  DISTINCT a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, "
    StrSql = StrSql & "                a.ItemCd, b.RoutinNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode)  Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_Routine     b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'O'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.RoutinCd "
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 "
    StrSql = StrSql & "     UNION ALL"
    StrSql = StrSql & "        SELECT  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2,"
    StrSql = StrSql & "                a.ItemCd, b.ITEMNM ItemName, a.GbIO, "
    StrSql = StrSql & "                rtrim(a.DeptCode)  Deptcode, c.DeptNameK "
    StrSql = StrSql & "        FROM   TWEXAM_Order       a, "
    StrSql = StrSql & "               TWEXAM_iTEMML      b, "
    StrSql = StrSql & "               TW_MIS_PMPA.TWBAS_Dept         c "
    StrSql = StrSql & "        WHERE  a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.GbIO      = 'O'"
    StrSql = StrSql & "        AND    a.JeobsuYN  = '*'"
    StrSql = StrSql & "        AND    a.ItemCd    = b.Codeky"
    StrSql = StrSql & "        AND    a.DeptCode  = c.DeptCode(+) "
    StrSql = StrSql & "        AND    a.SLipno1   > 0 "
    StrSql = StrSql & "        AND    a.SLipno1   < 52 )"
    'strSql = strSql & "        GROUP BY ItemCd, b.ITEMNM, a.GBIO, a.DeptCode, c.DeptNamek) "
    
    StrSql = StrSql & " GROUP  BY itemCd, itemName"
    
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    sprDepto.MaxCols = 100
    sprDepto.ColWidth(-1) = 5
    sprDepto.ColWidth(1) = 7.38
    sprDepto.ColWidth(2) = 19.38
    
    
    sprDepto.ReDraw = False
    Do Until adoSet.EOF
        sprDepto.Row = sprDepto.DataRowCnt + 1
        sprDepto.Col = 1: sprDepto.Text = adoSet.Fields("ItemCd").Value & ""
        sprDepto.Col = 2: sprDepto.Text = adoSet.Fields("ItemName").Value & ""
        
        iCol = 3
        For i = 0 To UBound(sDept)
            sprDepto.Col = i + 3: sprDepto.Text = adoSet.Fields(sDept(i)).Value & ""
            sprDepto.Row = 0: sprDepto.Text = sDept(i)
            iCol = iCol + 1
            sprDepto.Row = sprDepto.DataRowCnt
        Next
        sprDepto.Col = UBound(sDept) + 4: sprDepto.Text = adoSet.Fields("LineTotal").Value & ""
        adoSet.MoveNext
    Loop
    sprDepto.ReDraw = True
    
    sprDepto.MaxCols = UBound(sDept) + 4
    sprDepto.Col = sprDepto.MaxCols
    sprDepto.Row = 0: sprDepto.Text = "합계"

    
    Call adoSetClose(adoSet)
    
    Return
    
Total_Calculate:
    Dim iLastSprRow  As Integer
    Dim nCalSum      As Single
    Dim nDataRow     As Integer
    
    
    iLastSprRow = sprDepto.DataRowCnt
    nDataRow = iLastSprRow - 1
    
    For j = 3 To sprDepto.MaxCols
        nCalSum = 0
        For i = 1 To iLastSprRow
            sprDepto.Row = i
            sprDepto.Col = j
            If Trim(sprDepto.Text) <> "" Then
               nCalSum = nCalSum + CSng(sprDepto.Text)
            End If
        Next
        sprDepto.Row = iLastSprRow + 2
        sprDepto.Col = j
        sprDepto.Text = nCalSum
    Next
    
    sprDepto.Row = sprDepto.DataRowCnt
    sprDepto.Col = 2: sprDepto.Text = "합계"
    Return

End Sub



Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub Option1_Click()
    
    paneli.Visible = False
    panelo.Visible = False
    
    panelio.Top = 765
    panelio.Left = 90
    panelio.Visible = True
    panelio.ZOrder 0
    
    
End Sub

Private Sub Option2_Click()
    
    panelio.Visible = False
    paneli.Visible = False
    
    panelo.Top = 765
    panelo.Left = 90
    panelo.Visible = True
    panelo.ZOrder 0

End Sub

Private Sub Option3_Click()
    
    panelio.Visible = False
    panelo.Visible = False
    
    paneli.Top = 765
    paneli.Left = 90
    paneli.Visible = True
    paneli.ZOrder 0

End Sub
