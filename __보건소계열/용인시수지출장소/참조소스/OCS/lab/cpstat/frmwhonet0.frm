VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmWhonet0 
   Caption         =   "일자별균주 백분율"
   ClientHeight    =   7080
   ClientLeft      =   225
   ClientTop       =   1260
   ClientWidth     =   11595
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
   ScaleHeight     =   7080
   ScaleWidth      =   11595
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprWho 
      Height          =   5910
      Left            =   135
      TabIndex        =   4
      Top             =   810
      Width           =   11130
      _Version        =   196608
      _ExtentX        =   19632
      _ExtentY        =   10425
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
      MaxCols         =   14
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmWhonet0.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
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
         Format          =   24444931
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
         Format          =   24444931
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
      Height          =   510
      Left            =   8055
      TabIndex        =   6
      Top             =   180
      Width           =   1590
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "2805;900"
      Picture         =   "frmWhonet0.frx":133E
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   510
      Left            =   6300
      TabIndex        =   5
      Top             =   180
      Width           =   1770
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "3122;900"
      Picture         =   "frmWhonet0.frx":1C18
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
Attribute VB_Name = "frmWhonet0"
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
    
    
    If sprWho.DataRowCnt = 0 Then
        MsgBox "출력할 Data 가 없습니다!...", vbInformation
        Exit Sub
    End If
    
    strFont0 = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont1 = "/fn""바탕체"" /fz""12"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "/l" & "Antimicrobial resistance(%) of gram_negative bacili frequently isolated"
    strHead3 = "/f3" & "조회조건 : " & Format(dtFrDate.Value, "yyyy-MM-dd") & " ~ " & _
                                       Format(dtToDate.Value, "yyyy-MM-dd")
    strHead5 = "/f4" & "    "
    
    sprWho.PrintHeader = strHead1 + strFont1 + strHead2 + strFont0 + "/n/n" + _
                                    strFont2 + "/l" + strHead3 + _
                                    strFont2 + strHead5
    sprWho.PrintFooter = strFont2 + "/l" + sPortBar + "/n" + _
                            Space(120) + "발행일자:" + Dual_Date_Get("yyyy-MM-dd")
                            
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
    sprWho.PrintOrientation = PrintOrientationLandscape
    sprWho.Action = SS_ACTION_PRINT


End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sOrgCode        As String
    Dim sOrgReturnCount As String
    
    
    Call SpreadSetClear(sprWho)
    GoSub SpreadTitle_Set_sprWho
    
    sFrDate = Format(dtFrDate.Value, "YYYY-MM-DD")
    sToDate = Format(dtToDate.Value, "YYYY-MM-DD")
    
    For i = 2 To sprWho.MaxCols
        sprWho.Row = 0
        sprWho.Col = i: sOrgCode = sprWho.Text
        GoSub TOTAL_COUNT_ORGLIST
        
        sprWho.Row = 1
        sprWho.Col = i: sprWho.Text = sOrgCode & vbCrLf & "(" & sOrgReturnCount & ")"
        
        sprWho.Row = 2
        sprWho.Col = i: sprWho.Text = sOrgReturnCount
    Next
    GoSub DeCode_Group_AntiList
    GoSub Data_Percentage_Calc
    
    sprWho.Row = sprWho.DataRowCnt + 1
    sprWho.Col = 1
    sprWho.Action = ActionActiveCell
    
    Exit Sub
    


SpreadTitle_Set_sprWho:
    Dim sTmpText        As String
    
    sprWho.Row = 1
    sprWho.Col = 1: sprWho.Text = "Antimicrobial"
    
    For i = 2 To sprWho.MaxCols
        sprWho.Row = 0
        sprWho.Col = i: sTmpText = sprWho.Text
        
        sprWho.Row = 1
        sprWho.Col = i: sprWho.Text = sTmpText
    Next
    
    Return

TOTAL_COUNT_ORGLIST:
    sOrgReturnCount = ""
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_General_Sub"
    strSql = strSql & " WHERE  JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  = 42"
    strSql = strSql & " AND    VERIFY    = 'Y'"
    strSql = strSql & " AND   ( RTRIM(RCode1) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode2) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode3) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode4) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode5) = '" & LCase(sOrgCode) & "')"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    sOrgReturnCount = adoSet.RecordCount
    Call adoSetClose(adoSet)
    
    Return


DeCode_Group_AntiList:
    strSql = ""
    strSql = strSql & " SELECT AntiName, "
    strSql = strSql & "        SUM(Decode(OraCode, 'eco', cnt, '')) ECO, "
    strSql = strSql & "        SUM(Decode(OraCode, 'cfr', cnt, '')) CFR, "
    strSql = strSql & "        SUM(Decode(OraCode, 'kpn', cnt, '')) KPN, "
    strSql = strSql & "        SUM(Decode(OraCode, 'kox', cnt, '')) KOX, "
    strSql = strSql & "        SUM(Decode(OraCode, 'ecl', cnt, '')) ECL, "
    strSql = strSql & "        SUM(Decode(OraCode, 'eae', cnt, '')) EAE, "
    strSql = strSql & "        SUM(Decode(OraCode, 'sma', cnt, '')) SAM, "
    strSql = strSql & "        SUM(Decode(OraCode, 'pmi', cnt, '')) PMI, "
    strSql = strSql & "        SUM(Decode(OraCode, 'pvu', cnt, '')) PVU, "
    strSql = strSql & "        SUM(Decode(OraCode, 'mmo', cnt, '')) MMO, "
    strSql = strSql & "        SUM(Decode(OraCode, 'pst', cnt, '')) PST, "
    strSql = strSql & "        SUM(Decode(OraCode, 'aba', cnt, '')) ABA, "
    strSql = strSql & "        SUM(Decode(OraCode, 'smp', cnt, '')) SMP "
    strSql = strSql & " FROM(  SELECT RTRIM(a.OraCod) OraCode, a.yakcod, b.codenm AntiName,  count(*) cnt "
    strSql = strSql & "        FROM   TWEXAM_SENS     a, "
    strSql = strSql & "               TWEXAM_ANTILIST b "
    strSql = strSql & "        WHERE  a.JeobsuDt >=  TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    strSql = strSql & "        AND    a.JeobsuDt <=  TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    strSql = strSql & "        AND    a.yakcod    = b.codeky(+) "
    strSql = strSql & "        Group  By OraCod, yakcod, b.Codenm) "
    strSql = strSql & " GROUP  BY AntiName"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprWho.Row = sprWho.DataRowCnt + 1
        sprWho.Col = 1:  sprWho.Text = adoSet.Fields("ANTINAME").Value & ""
        sprWho.Col = 2:  sprWho.Text = adoSet.Fields("ECO").Value & ""
        sprWho.Col = 3:  sprWho.Text = adoSet.Fields("CFR").Value & ""
        sprWho.Col = 4:  sprWho.Text = adoSet.Fields("KPN").Value & ""
        sprWho.Col = 5:  sprWho.Text = adoSet.Fields("KOX").Value & ""
        sprWho.Col = 6:  sprWho.Text = adoSet.Fields("ECL").Value & ""
        sprWho.Col = 7:  sprWho.Text = adoSet.Fields("EAE").Value & ""
        sprWho.Col = 8:  sprWho.Text = adoSet.Fields("SAM").Value & ""
        sprWho.Col = 9:  sprWho.Text = adoSet.Fields("PMI").Value & ""
        sprWho.Col = 10: sprWho.Text = adoSet.Fields("PVU").Value & ""
        sprWho.Col = 11: sprWho.Text = adoSet.Fields("MMO").Value & ""
        sprWho.Col = 12: sprWho.Text = adoSet.Fields("ABA").Value & ""
        sprWho.Col = 13: sprWho.Text = adoSet.Fields("SMP").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
Data_Percentage_Calc:
    Dim iTotalCnt       As Integer
    
    For i = 2 To sprWho.MaxCols
        For j = 3 To sprWho.DataRowCnt
            sprWho.Row = 2: sprWho.Col = i
            iTotalCnt = Val(sprWho.Text)
            
            sprWho.Row = j: sprWho.Col = i
            If Val(sprWho.Text) > 0 Then
                If iTotalCnt > 0 Then
                    sprWho.Text = Format((Val(sprWho.Text) / iTotalCnt) * 100, "##.0")
                End If
            End If
            iTotalCnt = 0
        Next
    Next
    Return
    
    
End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

