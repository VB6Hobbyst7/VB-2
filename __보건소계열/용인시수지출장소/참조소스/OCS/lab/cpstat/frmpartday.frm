VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPartDay 
   Caption         =   "검사통계(일자별)"
   ClientHeight    =   7245
   ClientLeft      =   315
   ClientTop       =   1215
   ClientWidth     =   11130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   11130
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   135
      TabIndex        =   0
      Top             =   585
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
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
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   330
         Left            =   1080
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
      Begin MSForms.CommandButton cmdQuery 
         Height          =   420
         Left            =   2655
         TabIndex        =   3
         Top             =   90
         Width           =   1680
         Caption         =   "조회확인"
         Size            =   "2963;741"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   780
      End
   End
   Begin FPSpreadADO.fpSpread sprSLip 
      Height          =   6495
      Left            =   4905
      TabIndex        =   4
      Top             =   540
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9551
      _ExtentY        =   11456
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
      MaxCols         =   4
      MaxRows         =   40
      ScrollBars      =   2
      SpreadDesigner  =   "frmPartDay.frx":0000
      Appearance      =   2
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmPartDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuery_Click()
    Dim sCount      As Integer
    Dim sDate       As String
    Dim iSLno       As Integer
    
    sDate = Format(dtDate.Value, "yyyy-MM-dd")

    Call SpreadSetClear(sprSLip)
    GoSub Get_SLIPcount
    GoSub Get_PtCount
    Exit Sub



Get_SLIPcount:
    StrSql = ""
    StrSql = StrSql & "  SELECT a.SLIPNO1, b.Codenm,  COUNT(*) Count"
    StrSql = StrSql & "  FROM   TWEXAM_ORDER   a,"
    StrSql = StrSql & "         TWEXAM_SPECODE b"
    StrSql = StrSql & "  WHERE  a.COLLDate = TO_DATE('" & sDate & "','YYYY-MM-DD') "
    StrSql = StrSql & "  AND    a.JeobsuYn = '*'"
    StrSql = StrSql & "  AND    a.SLipno1  > 0 "
    StrSql = StrSql & "  AND    a.SLipno1  < 52"
    StrSql = StrSql & "  AND    a.SLipno1  = b.Codeky"
    StrSql = StrSql & "  AND    b.Codegu   = '12'"
    StrSql = StrSql & "  GROUP BY SLIPNO1, b.Codenm"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprSLip.Row = sprSLip.DataRowCnt + 1
        sprSLip.Col = 1: sprSLip.Text = adoSet.Fields("SLipno1").Value & ""
                         iSLno = Val(adoSet.Fields("SLipno1").Value & "")
        sprSLip.Col = 2: sprSLip.Text = adoSet.Fields("Codenm").Value & ""
        sprSLip.Col = 3: GoSub Get_PtCount
                         sprSLip.Text = sCount
        sprSLip.Col = 4: sprSLip.Text = adoSet.Fields("Count").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
Get_PtCount:
    Dim adoPtCnt        As ADODB.Recordset
    
    sCount = 0
    StrSql = ""
    StrSql = StrSql & " SELECT ptno"
    StrSql = StrSql & " FROM   TWEXAM_Order"
    StrSql = StrSql & " WHERE  COLLDate = TO_DATE('" & sDate & "','YYYY-MM-DD') "
    StrSql = StrSql & " AND    JeobsuYN = '*'"
    StrSql = StrSql & " AND    SLipno1  = " & iSLno
    StrSql = StrSql & " GROUP  BY Ptno"
    If False = adoSetOpen(StrSql, adoPtCnt) Then Return
    sCount = adoPtCnt.RecordCount
    Call adoSetClose(adoPtCnt)
    
    Return

End Sub

Private Sub Form_Load()
    
    dtDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
