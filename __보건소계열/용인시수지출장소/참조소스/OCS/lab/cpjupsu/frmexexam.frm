VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Begin VB.Form frmExExam 
   Caption         =   "외부의뢰 관리"
   ClientHeight    =   4950
   ClientLeft      =   330
   ClientTop       =   2160
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   11415
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread ssResult 
      Height          =   6015
      Left            =   90
      TabIndex        =   0
      Top             =   1215
      Width           =   11775
      _Version        =   196608
      _ExtentX        =   20770
      _ExtentY        =   10610
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421376
      MaxCols         =   12
      MaxRows         =   600
      Position        =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmExExam.frx":0000
      UserResize      =   0
      VisibleCols     =   10
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   450
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   794
      _StockProps     =   15
      Caption         =   "외 부 의 뢰 관 리 "
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Font3D          =   4
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1035
      Left            =   4095
      TabIndex        =   4
      Top             =   30
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1826
      _StockProps     =   14
      Caption         =   "조  건"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   330
         Left            =   855
         TabIndex        =   14
         Top             =   585
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   24444931
         CurrentDate     =   36381
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   330
         Left            =   855
         TabIndex        =   13
         Top             =   225
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm "
         Format          =   24444931
         CurrentDate     =   36381.0000115741
      End
      Begin Threed.SSOption optComplete 
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   " 완  료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optJeobsu 
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   465
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   " 의뢰중"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optAll 
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   210
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   " 접수중"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   285
         TabIndex        =   9
         Top             =   630
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   285
         TabIndex        =   8
         Top             =   285
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdInquiry 
      Height          =   840
      Left            =   8280
      TabIndex        =   7
      Top             =   90
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "조 회"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "frmExExam.frx":1FF2
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   840
      Left            =   10680
      TabIndex        =   6
      Top             =   90
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "이전화면"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "frmExExam.frx":230C
   End
   Begin Threed.SSCommand cmdResult 
      Height          =   840
      Left            =   9480
      TabIndex        =   5
      Top             =   90
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1482
      _StockProps     =   78
      Caption         =   "등 록"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "frmExExam.frx":2626
   End
   Begin Spin.SpinButton spinFromDate 
      Height          =   360
      Left            =   12000
      TabIndex        =   3
      Top             =   3660
      Width           =   195
      _Version        =   65536
      _ExtentX        =   344
      _ExtentY        =   635
      _StockProps     =   73
      ShadowThickness =   1
      TdThickness     =   1
   End
   Begin Spin.SpinButton spinToDate 
      Height          =   360
      Left            =   12000
      TabIndex        =   2
      Top             =   4080
      Width           =   195
      _Version        =   65536
      _ExtentX        =   344
      _ExtentY        =   635
      _StockProps     =   73
      ShadowThickness =   1
      TdThickness     =   1
   End
   Begin MSForms.CommandButton cmdSelect 
      Height          =   465
      Left            =   90
      TabIndex        =   16
      Top             =   585
      Width           =   1185
      Caption         =   "▼전체선택"
      Size            =   "2090;820"
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   465
      Left            =   2070
      TabIndex        =   15
      Top             =   585
      Width           =   1860
      Caption         =   "출력"
      PicturePosition =   327683
      Size            =   "3281;820"
      Picture         =   "frmExExam.frx":2940
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmExExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Dim LsCodeKy(0 To 100)     As String * 6
 Dim LsGeomchcd(0 To 100)   As String * 6
 Dim LsGeomsast(0 To 500)   As String * 6
 Dim LsJeobsuJA(0 To 100)   As String * 6
 Dim LsGbibo(0 To 10)       As String * 1
 Dim LsSex(0 To 2)          As String * 1
 Dim LsJinrye(0 To 200)     As String * 4
 Dim LsDoctor(0 To 500)     As String * 6
 Dim LsGeomsaJA(0 To 100)   As String * 6
  
 Dim LsJeobsuTm             As String
 Dim LsGeomsaDt             As String
 Dim LsCodegu               As String
 Dim LsDatech               As String
 Dim LsGbjupsu              As String
 Dim LsStatus               As String * 1
 
 Dim i                      As Integer
 Dim LiRowCnt               As Integer
 Dim LiDbseq                As Integer
 Dim LiChkFlg               As Integer
 Dim LiRd1Flg               As Integer      'Routine   Db read 이상확인용
 Dim LiGenFlg               As Integer      '접수/결과 Db read 이상확인용
 Dim LiGenSubFlg            As Integer      '접수/결과 Db read 이상확인용 SUb
 Dim LiIdNoFlg              As Integer      '환자 Db Read Check
 Dim LsRet                  As Integer      'Message Box display
 Dim LexItemPrRow()         As String       '외부의뢰검사 WorkSheet 발행 Item Check(General_Sub)

Private Sub cmdChoise_Click()

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdInquiry_Click()
    
    Dim LsPtNo       As String * 8
    Dim LsStatus     As String * 1
    Dim LsCodeKy     As String
    Dim i            As Integer
    Dim LiReccnt     As Integer
    Dim sFrJeobsu    As String
    Dim sToJeobsu    As String
    Dim sFrYYmmdd   As String
    Dim sToYYmmdd   As String
    
    Call Spread_Set_Clear(ssResult)
    
    ssResult.Col = 9: ssResult.Row = 1
    ssResult.Col2 = 9: ssResult.Row2 = ssResult.MaxRows
    ssResult.BlockMode = True
    ssResult.BackColor = RGB(250, 250, 225)
    ssResult.Lock = False
    ssResult.BlockMode = False
    
    sFrYYmmdd = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    sToYYmmdd = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    sFrJeobsu = Format(dtFromJeobsu.Value, "yyyy-MM-dd HH:mm")
    sToJeobsu = Format(dtToJeobsu.Value, "yyyy-MM-dd HH:mm")
    
    '
    ' 전체 검색 문장 (default)
    '
    
    If optAll = True Then
        cmdResult.Enabled = True
        strSql = ""
        strSql = strSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
        strSql = strSql & "        g.Result1, g.Result4, g.Result5, g.Verify, h.GeomchCd,"
        strSql = strSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname, p.Sname,"
        strSql = strSql & "        TO_CHAR(h.GBDate,'yyyy-MM-dd hh24:mi') GBDate"
        strSql = strSql & "  FROM  TWEXAM_General_Sub g,"
        strSql = strSql & "        TWEXAM_General     h,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      i,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode     j,"
        strSql = strSql & "        TWEXAM_IDNOMST     p"
        strSql = strSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrYYmmdd & "','YYYY-MM-DD') "
        strSql = strSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToYYmmdd & "','YYYY-MM-DD') "
        strSql = strSql & "  AND   g.Verify    = 'N'                          "
        strSql = strSql & "  AND   i.GeomsaGb  = 'W'                          "
        strSql = strSql & "  AND   g.ItemCd    = I.CodeKy                     "
        strSql = strSql & "  AND  (g.Result4   = ' ' or G.Result4  IS NULL )  "
        strSql = strSql & "  AND   g.JeobsuDt  = h.JeobsuDt(+)"
        strSql = strSql & "  AND   g.SLipno1   = h.SLipno1(+)"
        strSql = strSql & "  AND   g.SLipno2   = h.SLipno2(+)"
        strSql = strSql & "  AND   h.GBdate   >=  TO_DATE('" & sFrJeobsu & "','yyyy-MM-dd hh24:mi')"
        strSql = strSql & "  AND   h.GBDate   <=  TO_DATE('" & sToJeobsu & "','yyyy-MM-dd hh24:mi')"
        strSql = strSql & "  AND   j.Codegu    = '12'"
        strSql = strSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
        strSql = strSql & "  AND   g.Ptno      = p.Ptno(+)"
        strSql = strSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    ElseIf optJeobsu = True Then
        cmdResult.Enabled = False
        strSql = ""
        strSql = strSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
        strSql = strSql & "        g.Result1, g.Result4, g.Result5, g.Verify, h.GeomchCd,"
        strSql = strSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname, p.Sname,"
        strSql = strSql & "        TO_CHAR(h.GBDate,'yyyy-MM-dd hh24:mi') GBDate"
        strSql = strSql & "  FROM  TWEXAM_General_Sub g,"
        strSql = strSql & "        TWEXAM_General     h,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      i,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode     j,"
        strSql = strSql & "        TWEXAM_IDNOMST     p"
        strSql = strSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrYYmmdd & "','YYYY-MM-DD') "
        strSql = strSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToYYmmdd & "','YYYY-MM-DD') "
        strSql = strSql & "  AND   g.Verify    = 'N'                          "
        strSql = strSql & "  AND   i.GeomsaGb  = 'W'                          "
        strSql = strSql & "  AND   g.ItemCd    = I.CodeKy                     "
        strSql = strSql & "  AND   g.Result4   = '1'                          "    'Print Check
        strSql = strSql & "  AND   j.Codegu    = '12'"
        strSql = strSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
        strSql = strSql & "  AND   g.JeobsuDt  = h.JeobsuDt(+)"
        strSql = strSql & "  AND   g.SLipno1   = h.SLipno1(+)"
        strSql = strSql & "  AND   g.SLipno2   = h.SLipno2(+)"
        strSql = strSql & "  AND   h.GBdate   >=  TO_DATE('" & sFrJeobsu & "','yyyy-MM-dd hh24:mi')"
        strSql = strSql & "  AND   h.GBDate   <=  TO_DATE('" & sToJeobsu & "','yyyy-MM-dd hh24:mi')"
        strSql = strSql & "  AND   g.Ptno      = p.Ptno(+)"
        strSql = strSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    ElseIf optComplete = True Then
        cmdResult.Enabled = False
        strSql = ""
        strSql = strSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
        strSql = strSql & "        g.Result1, g.Result4, g.Result5, g.Verify, h.GeomchCd,"
        strSql = strSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname, p.Sname,"
        strSql = strSql & "        TO_CHAR(h.GBDate,'yyyy-MM-dd hh24:mi') GBDate"
        strSql = strSql & "  FROM  TWEXAM_General_Sub g,"
        strSql = strSql & "        TWEXAM_General     h,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      i,"
        strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode     j,"
        strSql = strSql & "        TWEXAM_IDNOMST     p"
        strSql = strSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrYYmmdd & "','YYYY-MM-DD') "
        strSql = strSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToYYmmdd & "','YYYY-MM-DD') "
        strSql = strSql & "  AND   i.GeomsaGb  = 'W'"
        strSql = strSql & "  AND   g.Verify    = 'Y'"
        strSql = strSql & "  AND   g.ItemCd    = I.CodeKy"
        strSql = strSql & "  AND   g.Result1  <> ' '"
        strSql = strSql & "  AND   g.Result4   = '1'"
        strSql = strSql & "  AND   j.Codegu    = '12'"
        strSql = strSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
        strSql = strSql & "  AND   g.JeobsuDt  = h.JeobsuDt(+)"
        strSql = strSql & "  AND   g.SLipno1   = h.SLipno1(+)"
        strSql = strSql & "  AND   g.SLipno2   = h.SLipno2(+)"
        strSql = strSql & "  AND   h.GBdate   >=  TO_DATE('" & sFrJeobsu & "','yyyy-MM-dd hh24:mi')"
        strSql = strSql & "  AND   h.GBDate   <=  TO_DATE('" & sToJeobsu & "','yyyy-MM-dd hh24:mi')"
        strSql = strSql & "  AND   g.Ptno      = p.Ptno(+)"
        strSql = strSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    End If

    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    

    LiReccnt = adoSet.RecordCount
    ssResult.ReDraw = False
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 2:  ssResult.Text = ssResult.Row
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 4:  ssResult.Text = Trim(adoSet.Fields("SLipname").Value & "")
                           LsCodeKy = ssResult.Text
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
                           LsPtNo = ssResult.Text
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("ITemNM").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("Result5").Value & ""
        
        If Trim(adoSet.Fields("Result4").Value & "") = "1" And _
           Trim(adoSet.Fields("Result1").Value & "") = "" Then
                ssResult.BackColor = RGB(235, 245, 235)
                ssResult.Lock = True
                ssResult.Col = 10: ssResult.Text = "의뢰중"
        ElseIf Trim(adoSet.Fields("REsult4").Value & "") = "" And _
               Trim(adoSet.Fields("Result1").Value & "") = "" Then
                ssResult.Col = 10: ssResult.Text = "접수중"
        ElseIf Trim(adoSet.Fields("REsult4").Value & "") = "1" And _
               Trim(adoSet.Fields("Result1").Value & "") <> "" Then
                ssResult.BackColor = RGB(235, 245, 235)
                ssResult.Lock = True
                ssResult.Col = 10: ssResult.Text = "완  료"
        End If
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("RowID").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("GeomchCD").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    ssResult.Col = 1: ssResult.Row = 1
    ssResult.Col2 = 1: ssResult.Row2 = ssResult.MaxRows
    ssResult.BlockMode = True
    ssResult.Lock = False
    ssResult.BlockMode = False


    ssResult.ReDraw = True
    ssResult.SetFocus

End Sub

Private Sub cmdPrint_Click()
    Dim sKey            As String * 7
    Dim sJeobsuDt       As String * 10
    Dim sSLipno1        As String * 2
    Dim sSLipno2        As String * 5
    Dim sPtno           As String * 8
    Dim sName           As String * 10
    Dim sSexAge         As String * 5
    Dim sItemCd         As String * 8
    Dim sItemNM         As String * 28
    Dim sSamplename     As String * 20
    Dim sDrcomment      As String * 80
    Dim sDept           As String * 12
    Dim sDr             As String * 10
    Dim sWard           As String * 5
    Dim sJumin          As String * 14
    Dim sBarLine        As String
    Dim sFrJeobsu       As String
    Dim sToJeobsu       As String
    Dim iLineCount      As Integer
    Dim sBiname         As String * 10
    Dim sOrderDt        As String
    Dim siLLCode        As String
    Dim siLLNameK       As String
    
    
    sFrJeobsu = Format(Me.dtFromJeobsu.Value, "yyyy-MM-dd")
    sToJeobsu = Format(Me.dtToJeobsu.Value, "yyyy-MM-dd")
    
    
    iLineCount = 0
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYYMMDD') JeobsuDt, g.RowID RWID,"
    strSql = strSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd, i.OldCode,"
    strSql = strSql & "        g.Result1, g.Result4, g.Result5, g.Verify, h.CmDoctor, h.GeomchCd, k.Codenm SampleName,"
    strSql = strSql & "        q.Deptnamek, r.Drname, h.RoomCode, h.Bi, h.GBIO,"
    strSql = strSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname,"
    strSql = strSql & "        p.Sname, p.Sex, p.Jumin1, p.Jumin2,"
    strSql = strSql & "        TO_CHAR(h.OrderDt, 'YYYY-MM-DD') OrderDt"
    strSql = strSql & "  FROM  TWEXAM_General_Sub g,"
    strSql = strSql & "        TWEXAM_General     h,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      i,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode     j,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample      k,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT      p,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT         q,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR       r "
    strSql = strSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrJeobsu & "','YYYY-MM-DD') "
    strSql = strSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToJeobsu & "','YYYY-MM-DD') "
    strSql = strSql & "  AND   g.Verify    = 'N'                          "
    strSql = strSql & "  AND   i.GeomsaGb  = 'W'                          "
    strSql = strSql & "  AND   g.ItemCd    = I.CodeKy                     "
    strSql = strSql & "  AND   g.Result4   = '1'"  '등록Button 으로 등록시킨 Data"
    strSql = strSql & "  AND   g.JeobsuDt  = h.JeobsuDt(+)"
    strSql = strSql & "  AND   g.SLipno1   = h.SLipno1(+)"
    strSql = strSql & "  AND   g.SLipno2   = h.SLipno2(+)"
    strSql = strSql & "  AND   j.Codegu    = '12'"
    strSql = strSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
    strSql = strSql & "  AND   h.GeomchCD  = k.Code(+)"
    strSql = strSql & "  AND   g.Ptno      = p.Ptno(+)"
    strSql = strSql & "  AND   h.DeptCode  = q.DeptCode(+)"
    strSql = strSql & "  AND   h.DrCode    = r.Drcode(+)"
    strSql = strSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    
    
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "조회할수 있는 Data 의 건수가 하나도 없습니다...", vbCritical
        Exit Sub
    End If
    
    Printer.Orientation = vbPRORLandscape
    
    sBarLine = ""
    
    For i = 1 To 80
        sBarLine = sBarLine & "━"
    Next
    
    
    GoSub Head_Print_Sub
    
    ReDim LexItemPrRow(adoSet.RecordCount - 1)
    
    
    i = 0: iLineCount = 0
    Do Until adoSet.EOF
        If iLineCount > 12 Then
            Printer.FontName = "굴림체"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.ForeColor = RGB(192, 192, 192)
            Printer.Print sBarLine
            Printer.ForeColor = RGB(0, 0, 0)
            Printer.Print Space(120) & "발행일: " & Dual_Date_Get("yyyy-MM-dd") & "  Page: " & Printer.Page
            Printer.NewPage
            GoSub Head_Print_Sub
            iLineCount = 0
        End If
                
        
        sKey = Format(adoSet.Fields("SLipno1").Value & "", "00") & _
               Format(adoSet.Fields("SLipno2").Value & "", "00000")
        
        LexItemPrRow(i) = adoSet.Fields("RWID").Value & ""
        
        sJeobsuDt = adoSet.Fields("JeobsuDt").Value & ""
        sSLipno1 = adoSet.Fields("Slipno1").Value & ""
        sSLipno2 = adoSet.Fields("Slipno2").Value & ""
        sPtno = adoSet.Fields("Ptno").Value & ""
        sName = adoSet.Fields("Sname").Value & ""
        sSexAge = adoSet.Fields("Sex").Value & "/" & _
                  SetAge_Check(adoSet.Fields("Jumin1").Value & "", adoSet.Fields("Jumin2").Value & "")
        'sItemCd = adoSet.Fields("ItemCD").Value & ""      '건양대병원 ItemCode
        sItemCd = adoSet.Fields("OLdCode").Value & ""      'SCL        ItemCode
        sItemNM = adoSet.Fields("ItemNM").Value & ""
        sSamplename = adoSet.Fields("Samplename").Value & ""
        sDrcomment = Trim(adoSet.Fields("CMDoctor").Value & "")
        sDept = adoSet.Fields("DeptNamek").Value & ""
        sDr = adoSet.Fields("Drname").Value & ""
        sJumin = adoSet.Fields("Jumin1").Value & "-" & adoSet.Fields("Jumin2").Value & ""
        sBiname = Bi_Check(adoSet.Fields("Bi").Value & "")
        sOrderDt = adoSet.Fields("OrderDt").Value & ""
        
        
        Printer.FontName = "굴림체"
        Printer.FontSize = 9
        Printer.FontBold = False
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print sBarLine
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.Print Tab(2); sJeobsuDt; Tab(12); sKey; Tab(21); sPtno; Tab(32); Trim(sName); Tab(42); sSexAge; _
                      Tab(48); sJumin; Tab(63); Trim(sBiname); Tab(73); Trim(sDept); Tab(86); Trim(sDr); Tab(96); sItemCd; _
                      Tab(106); sItemNM; Tab(138); sSamplename
        
        GoSub Get_iLLCodeData
        
        If Trim(sDrcomment) = "" Then
            If Trim(siLLCode) = "" Then
                Printer.Print ""
            Else
                Printer.Print Tab(12); "상병:" & Trim(siLLCode) & "/" & Trim(siLLNameK)
            End If
        Else
            If Trim(siLLCode) = "" Then
                Printer.Print Tab(75); "Rem) " & sDrcomment
            Else
                Printer.Print Tab(12); "상병:" & Trim(siLLCode) & "/" & Trim(siLLNameK); Tab(75); "Rem) " & Trim(sDrcomment)
            End If
        End If
        
        'Printer.FontName = "Code39Two"
        'Printer.FontSize = 16
        'Printer.FontBold = False
        'Printer.Print Space(2) & "*" & sKey & Trim(sItemCd) & "*"
        
        adoSet.MoveNext: i = i + 1: iLineCount = iLineCount + 1
        
    Loop
    
    Printer.FontName = "굴림체"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.ForeColor = RGB(192, 192, 192)
    Printer.Print sBarLine
    Printer.ForeColor = RGB(0, 0, 0)
    Printer.Print Space(120) & "발행일: " & Dual_Date_Get("yyyy-MM-dd") & "  Page: " & Printer.Page
    
    
    Printer.EndDoc
    Call adoSetClose(adoSet)
    
    'If vbYes = MsgBox("외부의뢰 검사가 바르게 출력되었습니까?", vbYesNo + vbQuestion, "FlagSetting") Then
    '    GoSub Set_PrintOK_Flag
    'End If
    
    Exit Sub
    

Get_iLLCodeData:
    Dim adoiLL      As ADODB.Recordset
    
    siLLCode = ""
    siLLNameK = ""
    
    If Trim(adoSet.Fields("GbIO").Value & "") = "I" Then
        'strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TWBAS_iLLs INX_iLLs0) */"
        
        strSql = ""
        strSql = strSql & " SELECT a.iLLCode, b.iLLnamek"
        strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_Master a,"
        strSql = strSql & "        TW_MIS_PMPA.TWBas_iLLs   b"
        strSql = strSql & " WHERE  a.Ptno   =  '" & sPtno & "'"
        strSql = strSql & " AND    RPAD(a.iLLCode, 6)  = b.iLLCode(+)"
    Else
        'strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TWBAS_iLLs INX_iLLs0) */"
        
        strSql = ""
        strSql = strSql & " SELECT a.iLLCode, b.iLLNameK"
        strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_OiLLs a,"
        strSql = strSql & "        TW_MIS_PMPA.TWBAS_iLLs  b "
        strSql = strSql & " WHERE  a.Ptno    =  '" & sPtno & "'"
        strSql = strSql & " AND    a.Bdate   = TO_DATE('" & sOrderDt & "','yyyy-MM-dd')"
        strSql = strSql & " AND    RPAD(a.iLLCode,6) = b.iLLCode(+)"
        strSql = strSql & " ORDER  BY a.Seqno"
    End If
    If False = adoSetOpen(strSql, adoiLL) Then Return
    siLLCode = adoiLL.Fields("iLLCode").Value & ""
    siLLNameK = adoiLL.Fields("iLLNameK").Value & ""
    
    Call adoSetClose(adoiLL)
    Return
    
    
Head_Print_Sub:
    Printer.FontName = "바탕체"
    Printer.FontSize = "20"
    Printer.FontBold = True
    Printer.Print Space(20) & " 외부의뢰 검사List"
    Printer.Print ""
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print " 일자: " & sFrJeobsu & " ~ " & sToJeobsu
    
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.Print ""
    Printer.Print Tab(2); "검체접수"; Tab(12); "Labno"; Tab(21); "등록번호"; Tab(32); "환자명"; Tab(41); "성/age"; _
                  Tab(49); "주민번호"; Tab(63); "구분"; Tab(73); "의뢰과"; Tab(86); "의사명"; Tab(96); "SCLCode"; _
                  Tab(106); "검사명"; Tab(138); "검체명"
    
    
    Return



Set_PrintOK_Flag:
    
    For i = LBound(LexItemPrRow) To UBound(LexItemPrRow)
        strSql = ""
        strSql = strSql & " UPDATE TWEXAM_General_Sub"
        strSql = strSql & " SET    Result4 = '1'"               '의뢰중 Flag 로 Change
        strSql = strSql & " WHERE  ROWID   = '" & LexItemPrRow(i) & "'"
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
    Next
    
    For i = LBound(LexItemPrRow) To UBound(LexItemPrRow)
        LexItemPrRow(i) = ""
    Next
    
    MsgBox "외부검사LIST WORKSHEET 발행 작업이 모두 끝났습니다!..", vbInformation
    
    Return
    
    
End Sub

Private Sub CmdResult_Click()
    
    Dim i               As Integer
    
    Dim sResult4        As String
    Dim sResult5        As String
    Dim sRowID          As String
    
        
    
    For LiRowCnt = 1 To Val(ssResult.DataRowCnt)
        ssResult.Row = LiRowCnt
        ssResult.Col = 10              ' 원래는 9 SCL,녹십자 Check 한것임
        If Trim$(ssResult.Text) <> "" Then
            ssResult.Col = 1
            If ssResult.Value = True Then
                                     sResult4 = "1"
                ssResult.Col = 9:    sResult5 = Trim(ssResult.Text)
                ssResult.Col = 11:   sRowID = ssResult.Text
                
                GoSub CLIENT_SUB_UPDATE
            End If
        End If
    
    Next LiRowCnt

    Call Spread_Set_Clear(ssResult)

Exit Sub


'--------------------------------------------------------------------------------------------
CLIENT_SUB_UPDATE:
        
    strSql = ""
    strSql = strSql & " UPDATE  TWEXAM_GENERAL_SUB "
    strSql = strSql & " SET     Result4    =   '" & sResult4 & "',"
    strSql = strSql & "         Result5    =   '" & sResult5 & "'"
    strSql = strSql & " WHERE   ROWID      =   '" & sRowID & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdSelect_Click()
    
    If cmdSelect.Caption = "▼전체선택" Then
    
        For i = 1 To ssResult.DataRowCnt
            ssResult.Row = i
            ssResult.Col = 1
            ssResult.Value = True
        Next
        cmdSelect.Caption = "▼전체해제"
    Else
        For i = 1 To ssResult.DataRowCnt
            ssResult.Row = i
            ssResult.Col = 1
            ssResult.Value = False
        Next
        cmdSelect.Caption = "▼전체선택"
    End If

End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Activate()
    
    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_Load()
 
    dtFromJeobsu.Value = Dual_Date_Get("yyyy-MM-dd") & " 00:01"
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    
 
    'dtFromJeobsu.Value = Dual_Date_Cal_Get("yyyy-MM-dd", -5)
    'dtToJeobsu.Value = Dual_Date_Cal_Get("yyyy-MM-dd", 0)
    
    optAll = True


    'Select a block of cells
    ssResult.Col = 1
    ssResult.Row = 0
    ssResult.Row2 = -1
    ssResult.Col2 = 7
    ssResult.BlockMode = True
    ssResult.Lock = True
    ssResult.BlockMode = False

    'Select a block of cells
    ssResult.Col = 9
    ssResult.Row = 0
    ssResult.Row2 = -1
    ssResult.Col2 = 9
    ssResult.BlockMode = True
    ssResult.Lock = True
    ssResult.BlockMode = False

    'Define cursor type
    ssResult.CursorType = SS_CURSOR_TYPE_DEFAULT
    ssResult.CursorStyle = SS_CURSOR_STYLE_ARROW



    'Show or hide the current column
    ssResult.Col = 10
    ssResult.ColHidden = False

End Sub


Private Sub mnuExit_Click()

    Unload Me

End Sub



Private Sub ssResult_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Col <> 9 Then
        Exit Sub
    End If

    If Mode = 0 Then
        ssResult.Col = Col
        ssResult.Row = Row
        ssResult.CellType = SS_CELL_TYPE_EDIT
        ssResult.TypeHAlign = SS_CELL_H_ALIGN_LEFT
        ssResult.TypeEditCharCase = SS_CELL_EDIT_CASE_NO_CASE
        ssResult.TypeEditMultiLine = False
        ssResult.TypeEditLen = 10
    Else
        ssResult.Col = Col
        ssResult.Row = Row
        ssResult.CellType = SS_CELL_TYPE_COMBOBOX
        ssResult.TypeComboBoxList = "EWON" & Chr(9) & "SRL" & Chr(9) & "GCRL" & Chr(9) & "SCL" & Chr(9) & " "
        ssResult.TypeComboBoxEditable = False
    End If


End Sub



