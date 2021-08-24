VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmWhonet1 
   Caption         =   "항균제별 Sens 통계"
   ClientHeight    =   7455
   ClientLeft      =   255
   ClientTop       =   1065
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11550
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprWho 
      Height          =   6810
      Left            =   4275
      TabIndex        =   4
      Top             =   450
      Width           =   6675
      _Version        =   196608
      _ExtentX        =   11774
      _ExtentY        =   12012
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
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
      MaxCols         =   5
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmWhonet1.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1140
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   2011
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
         Left            =   2475
         TabIndex        =   1
         Top             =   450
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
         Left            =   990
         TabIndex        =   2
         Top             =   450
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36446
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자: From / To"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   180
         Width           =   1770
      End
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   465
      Left            =   2025
      TabIndex        =   6
      Top             =   1845
      Width           =   1410
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "2487;820"
      Picture         =   "frmWhonet1.frx":1013
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   465
      Left            =   270
      TabIndex        =   5
      Top             =   1845
      Width           =   1680
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2963;820"
      Picture         =   "frmWhonet1.frx":18ED
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmWhonet1"
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
    
    
    If sprWho.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""12"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "/l" & "Antimicrobial susceptibility (%) of P.aeruginosa"
    strHead3 = "/f3" & "조회조건 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & _
                                       Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & "    "
    
    sprWho.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                     strFont2 + "/l" + strHead3 + _
                                     strFont2 + strHead5
    sprWho.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(80) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
    sprWho.PrintMarginLeft = 0
    sprWho.PrintMarginRight = 0
    sprWho.PrintMarginTop = 0
    sprWho.PrintMarginBottom = 0
    sprWho.PrintColHeaders = True
    sprWho.PrintRowHeaders = False
    sprWho.PrintBorder = True
    sprWho.PrintColor = True
    sprWho.PrintGrid = True
    sprWho.PrintShadows = True
    sprWho.PrintUseDataMax = False
    
    sprWho.Row = 1: sprWho.Row2 = sprWho.DataRowCnt
    sprWho.Col = 1: sprWho.Col2 = sprWho.DataColCnt
    sprWho.PrintType = SS_PRINT_CELL_RANGE
    sprWho.PrintOrientation = PrintOrientationPortrait
    sprWho.Action = SS_ACTION_PRINT

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate     As String
    Dim sToDate     As String
    
    Dim iSS      As Integer
    Dim iII      As Integer
    Dim iRR      As Integer
    Dim iTotal   As Integer
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call SpreadSetClear(sprWho)
    sprWho.Row = 1
    sprWho.Col = 1: sprWho.Text = ""
    sprWho.Col = 2: sprWho.Text = "Antimicrobial"
    sprWho.Col = 3: sprWho.Text = "R"
    sprWho.Col = 4: sprWho.Text = "I"
    sprWho.Col = 5: sprWho.Text = "S"
    
    GoSub Get_Data
    
    sprWho.Row = sprWho.DataRowCnt + 1
    sprWho.Col = 1
    sprWho.Action = ActionActiveCell
    
    Exit Sub
    
    
    
Get_Data:
    StrSql = ""
    StrSql = StrSql & " SELECT YAKCOD, Codenm,"
    StrSql = StrSql & "        SUM(DECODE(SENS, 'S', 1, '')) S,"
    StrSql = StrSql & "        SUM(DECODE(SENS, 'I', 1, '')) I,"
    StrSql = StrSql & "        SUM(DECODE(SENS, 'R', 1, '')) R,"
    StrSql = StrSql & "        SUM(DECODE(SENS, '', '', 1 )) LineTotal"
    StrSql = StrSql & " FROM(  SELECT RTRIM(a.YAKCOD) YAKCOD, b.Codenm,a.SENS"
    StrSql = StrSql & "        FROM   TWEXAM_SENS     a,"
    StrSql = StrSql & "               TWEXAM_ANTILIST b"
    StrSql = StrSql & "        WHERE  a.JEOBSUDT >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.JEOBSUDT <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & "        AND    a.Yakcod    = b.Codeky)"
    StrSql = StrSql & " GROUP BY YAKCOD, Codenm"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprWho.Row = sprWho.DataRowCnt + 1
        sprWho.Col = 1:  sprWho.Text = adoSet.Fields("YakCod").Value & ""
        sprWho.Col = 2:  sprWho.Text = adoSet.Fields("Codenm").Value & ""
        iSS = Val(adoSet.Fields("S").Value & "")
        iII = Val(adoSet.Fields("I").Value & "")
        iRR = Val(adoSet.Fields("R").Value & "")
        iTotal = Val(adoSet.Fields("LineTotal").Value & "")
        
        If iTotal > 0 Then
            sprWho.Col = 3:  If iSS > 0 Then sprWho.Text = Format(Round((iSS / iTotal) * 100), "###.0")
            sprWho.Col = 4:  If iII > 0 Then sprWho.Text = Format(Round((iII / iTotal) * 100), "###.0")
            sprWho.Col = 5:  If iRR > 0 Then sprWho.Text = Format(Round((iRR / iTotal) * 100), "###.0")
        End If
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
