VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Begin VB.Form clp040 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "외부의뢰 관리"
   ClientHeight    =   4950
   ClientLeft      =   345
   ClientTop       =   1785
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread ssResult 
      Height          =   6645
      Left            =   75
      TabIndex        =   0
      Top             =   1110
      Width           =   11775
      _Version        =   196608
      _ExtentX        =   20770
      _ExtentY        =   11721
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
      MaxCols         =   10
      MaxRows         =   600
      Position        =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "Clp040.frx":0000
      UserResize      =   0
      VisibleCols     =   10
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1035
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   1826
      _StockProps     =   15
      Caption         =   "외 부 의 뢰 관 리 "
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   15.76
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
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
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36381
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   330
         Left            =   855
         TabIndex        =   13
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36381
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
            Size            =   9.01
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
            Size            =   9.01
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
            Size            =   9.01
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
      Height          =   885
      Left            =   8280
      TabIndex        =   7
      Top             =   180
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1561
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
      RoundedCorners  =   0   'False
      Picture         =   "Clp040.frx":4A67
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   885
      Left            =   10680
      TabIndex        =   6
      Top             =   180
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1561
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
      RoundedCorners  =   0   'False
      Picture         =   "Clp040.frx":4D81
   End
   Begin Threed.SSCommand cmdResult 
      Height          =   885
      Left            =   9480
      TabIndex        =   5
      Top             =   180
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1561
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
      RoundedCorners  =   0   'False
      Picture         =   "Clp040.frx":509B
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
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "clp040"
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
    
    
    Call SSInitialize(ssResult)
    
    ssResult.Col = 8: ssResult.Row = 1
    ssResult.Col2 = 8: ssResult.Row2 = ssResult.MaxRows
    ssResult.BlockMode = True
    ssResult.BackColor = RGB(250, 250, 225)
    ssResult.Lock = False
    ssResult.BlockMode = False
    
    
    sFrJeobsu = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    sToJeobsu = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    '
    ' 전체 검색 문장 (default)
    '
    
    If optAll = True Then
        cmdResult.Enabled = True
        gStrSql = ""
        gStrSql = gStrSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        gStrSql = gStrSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
        gStrSql = gStrSql & "        g.Result1, g.Result4, g.Result5, g.Verify,"
        gStrSql = gStrSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname, p.Sname"
        gStrSql = gStrSql & "  FROM  TWEXAM_General_Sub g,"
        gStrSql = gStrSql & "        TWEXAM_ItemML      i,"
        gStrSql = gStrSql & "        TWEXAM_Specode     j,"
        gStrSql = gStrSql & "        TWBAS_Patient      p"
        gStrSql = gStrSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrJeobsu & "','YYYY-MM-DD') "
        gStrSql = gStrSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToJeobsu & "','YYYY-MM-DD') "
        gStrSql = gStrSql & "  AND   g.Verify    = 'N'                          "
        gStrSql = gStrSql & "  AND   i.GeomsaGb  = 'W'                          "
        gStrSql = gStrSql & "  AND   g.ItemCd    = I.CodeKy                     "
        gStrSql = gStrSql & "  AND  (g.Result4   = ' ' or G.Result4  IS NULL )  "
        gStrSql = gStrSql & "  AND   j.Codegu    = '12'"
        gStrSql = gStrSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
        gStrSql = gStrSql & "  AND   g.Ptno      = p.Ptno(+)"
        gStrSql = gStrSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    ElseIf optJeobsu = True Then
        cmdResult.Enabled = False
        gStrSql = ""
        gStrSql = gStrSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        gStrSql = gStrSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
        gStrSql = gStrSql & "        g.Result1, g.Result4, g.Result5, g.Verify,"
        gStrSql = gStrSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname, p.Sname"
        gStrSql = gStrSql & "  FROM  TWEXAM_General_Sub g,"
        gStrSql = gStrSql & "        TWEXAM_ItemML      i,"
        gStrSql = gStrSql & "        TWEXAM_Specode     j,"
        gStrSql = gStrSql & "        TWBAS_Patient      p"
        gStrSql = gStrSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrJeobsu & "','YYYY-MM-DD') "
        gStrSql = gStrSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToJeobsu & "','YYYY-MM-DD') "
        gStrSql = gStrSql & "  AND   g.Verify    = 'N'                          "
        gStrSql = gStrSql & "  AND   i.GeomsaGb  = 'W'                          "
        gStrSql = gStrSql & "  AND   g.ItemCd    = I.CodeKy                     "
        gStrSql = gStrSql & "  AND   G.Result4   = '1'                          "
        gStrSql = gStrSql & "  AND   j.Codegu    = '12'"
        gStrSql = gStrSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
        gStrSql = gStrSql & "  AND   g.Ptno      = p.Ptno(+)"
        gStrSql = gStrSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    ElseIf optComplete = True Then
        cmdResult.Enabled = False
        gStrSql = ""
        gStrSql = gStrSql & " SELECT TO_CHAR(g.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        gStrSql = gStrSql & "        g.SLipno1, g.SLipno2, g.Codeky1, g.Ptno, g.ItemCd,"
        gStrSql = gStrSql & "        g.Result1, g.Result4, g.Result5, g.Verify,"
        gStrSql = gStrSql & "        i.Codeky,  i.GeomsaGb, i.ItemNM, g.RowID, j.Codenm SLipname, p.Sname"
        gStrSql = gStrSql & "  FROM  TWEXAM_General_Sub g,"
        gStrSql = gStrSql & "        TWEXAM_ItemML      i,"
        gStrSql = gStrSql & "        TWEXAM_Specode     j,"
        gStrSql = gStrSql & "        TWBAS_Patient      p"
        gStrSql = gStrSql & "  WHERE g.JeobsuDt >= TO_DATE('" & sFrJeobsu & "','YYYY-MM-DD') "
        gStrSql = gStrSql & "  AND   g.JeobsuDt <= TO_DATE('" & sToJeobsu & "','YYYY-MM-DD') "
        gStrSql = gStrSql & "  AND   i.GeomsaGb  = 'W'                          "
        gStrSql = gStrSql & "  AND   g.ItemCd    = I.CodeKy                     "
        gStrSql = gStrSql & "  AND   g.Result1  <> ' '                          "
        gStrSql = gStrSql & "  AND   g.Result4   = '1'                          "
        gStrSql = gStrSql & "  AND   j.Codegu    = '12'"
        gStrSql = gStrSql & "  AND   TO_NUMBER(j.Codeky)  =  g.SLipno1"
        gStrSql = gStrSql & "  AND   g.Ptno      = p.Ptno(+)"
        gStrSql = gStrSql & "  ORDER BY JeobsuDt, SlipNo1, SlipNo2, PtNo  ASC   "
    End If

    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    

    LiReccnt = adoSet.RecordCount
    ssResult.ReDraw = False
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:  ssResult.Text = ssResult.Row
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = Trim(adoSet.Fields("SLipname").Value & "")
                           LsCodeKy = ssResult.Text
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
                           LsPtNo = ssResult.Text
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("ITemNM").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Result5").Value & ""
        
        If Trim(adoSet.Fields("Result4").Value & "") = "1" And _
           Trim(adoSet.Fields("Result1").Value & "") = "" Then
                ssResult.BackColor = RGB(235, 245, 235)
                ssResult.Lock = True
                ssResult.Col = 9: ssResult.Text = "의뢰중"
        ElseIf Trim(adoSet.Fields("REsult4").Value & "") = "" And _
               Trim(adoSet.Fields("Result1").Value & "") = "" Then
                ssResult.Col = 9: ssResult.Text = "접수중"
        ElseIf Trim(adoSet.Fields("REsult4").Value & "") = "1" And _
               Trim(adoSet.Fields("Result1").Value & "") <> "" Then
                ssResult.BackColor = RGB(235, 245, 235)
                ssResult.Lock = True
                ssResult.Col = 9: ssResult.Text = "완  료"
        End If
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("RowID").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    ssResult.ReDraw = True
    ssResult.SetFocus
   

End Sub

Private Sub CmdResult_Click()
    
    Dim i               As Integer
    
    Dim sResult4        As String
    Dim sResult5        As String
    Dim sRowID          As String
    
        
    
    For LiRowCnt = 1 To Val(ssResult.DataRowCnt)
        ssResult.Row = LiRowCnt
        ssResult.Col = 8
        If Trim$(ssResult.Text) <> "" Then
                                 sResult4 = "1"
            ssResult.Col = 8:    sResult5 = Trim(ssResult.Text)
            ssResult.Col = 10:   sRowID = ssResult.Text
            
            GoSub CLIENT_SUB_UPDATE
        End If
    
    Next LiRowCnt

    Call SSInitialize(ssResult)

Exit Sub


'--------------------------------------------------------------------------------------------
CLIENT_SUB_UPDATE:
        
    gStrSql = ""
    gStrSql = gStrSql & " UPDATE  TWEXAM_GENERAL_SUB "
    gStrSql = gStrSql & " SET     Result4    =   '" & sResult4 & "',"
    gStrSql = gStrSql & "         Result5    =   '" & sResult5 & "'"
    gStrSql = gStrSql & " WHERE   ROWID      =   '" & sRowID & "'"
    adoConnect.BeginTrans
    If adoExec(gStrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

End Sub

Private Sub Form_Load()
 
    clp040.Left = 0
    clp040.Top = 0
    clp040.Height = MDIMain.ScaleHeight
    clp040.Width = MDIMain.ScaleWidth
    
    dtFromJeobsu.Value = Dual_Date_Cal_Get("yyyy-MM-dd", -5)
    dtToJeobsu.Value = Dual_Date_Cal_Get("yyyy-MM-dd", 0)
    
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

    If Col <> 8 Then
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



