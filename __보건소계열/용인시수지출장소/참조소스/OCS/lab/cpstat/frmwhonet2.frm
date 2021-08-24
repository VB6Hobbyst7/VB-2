VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmWhonet2 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "항균제 백분율"
   ClientHeight    =   6120
   ClientLeft      =   1260
   ClientTop       =   1365
   ClientWidth     =   8865
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
   ScaleHeight     =   6120
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprWho 
      Height          =   6630
      Left            =   5085
      TabIndex        =   0
      Top             =   675
      Width           =   6135
      _Version        =   196608
      _ExtentX        =   10821
      _ExtentY        =   11695
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
      SpreadDesigner  =   "frmWhonet2.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   675
      Width           =   4515
      _Version        =   65536
      _ExtentX        =   7964
      _ExtentY        =   1931
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
         Left            =   2430
         TabIndex        =   2
         Top             =   450
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
         Left            =   990
         TabIndex        =   3
         Top             =   450
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36446
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자: From / To"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   1770
      End
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   510
      Left            =   2160
      TabIndex        =   6
      Top             =   1935
      Width           =   1725
      Caption         =   "출력확인"
      PicturePosition =   327683
      Size            =   "3043;900"
      Picture         =   "frmWhonet2.frx":0FB1
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   510
      Left            =   360
      TabIndex        =   5
      Top             =   1935
      Width           =   1770
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "3122;900"
      Picture         =   "frmWhonet2.frx":188B
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
Attribute VB_Name = "frmWhonet2"
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
    strFont1 = "/fn""바탕체"" /fz""10"" /fb1 /fi0 /fu1 /fk0 /fs1"
    strFont2 = "/fn""바탕체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    
    strHead1 = "/f1" & Space(23)
    strHead2 = "/f2" & "/l" & "Antimicrobial resistance (%) of Staphylococcus and Enterococcus"
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
    
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_General_Sub"
    StrSql = StrSql & " WHERE  JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    SLipno1  = 42"
    StrSql = StrSql & " AND    VERIFY    = 'Y'"
    StrSql = StrSql & " AND   ( RTRIM(RCode1) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode2) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode3) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode4) = '" & LCase(sOrgCode) & "'" & " Or " & _
                               "RTRIM(RCode5) = '" & LCase(sOrgCode) & "')"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    sOrgReturnCount = adoSet.RecordCount
    Call adoSetClose(adoSet)
    
    Return


DeCode_Group_AntiList:
    StrSql = ""
    StrSql = StrSql & " SELECT AntiName, "
    StrSql = StrSql & "        SUM(Decode(OraCode, 'sau', cnt, '')) SAU, "
    StrSql = StrSql & "        SUM(Decode(OraCode, 'cns', cnt, '')) CNS, "
    StrSql = StrSql & "        SUM(Decode(OraCode, 'efa', cnt, '')) EFA, "
    StrSql = StrSql & "        SUM(Decode(OraCode, 'efm', cnt, '')) EFM "
    StrSql = StrSql & " FROM(  SELECT RTRIM(a.OraCod) OraCode, a.yakcod, b.codenm AntiName,  count(*) cnt "
    StrSql = StrSql & "        FROM   TWEXAM_SENS     a, "
    StrSql = StrSql & "               TWEXAM_ANTILIST b "
    StrSql = StrSql & "        WHERE  a.JeobsuDt >=  TO_DATE('" & sFrDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.JeobsuDt <=  TO_DATE('" & sToDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "        AND    a.yakcod    = b.codeky(+) "
    StrSql = StrSql & "        Group  By OraCod, yakcod, b.Codenm) "
    StrSql = StrSql & " GROUP  BY AntiName"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprWho.Row = sprWho.DataRowCnt + 1
        sprWho.Col = 1:  sprWho.Text = adoSet.Fields("ANTINAME").Value & ""
        sprWho.Col = 2:  sprWho.Text = adoSet.Fields("SAU").Value & ""
        sprWho.Col = 3:  sprWho.Text = adoSet.Fields("CNS").Value & ""
        sprWho.Col = 4:  sprWho.Text = adoSet.Fields("EFA").Value & ""
        sprWho.Col = 5:  sprWho.Text = adoSet.Fields("EFM").Value & ""
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
                    sprWho.Text = Format((Val(sprWho.Text) / iTotalCnt) * 100, "###.0")
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
