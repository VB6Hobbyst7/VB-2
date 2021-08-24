VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_ItemCode 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ITEM  코드 처리"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   1860
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   12060
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture3 
      Height          =   645
      Left            =   240
      ScaleHeight     =   585
      ScaleWidth      =   6390
      TabIndex        =   6
      Top             =   525
      Width           =   6450
      Begin VB.ComboBox SpeCod 
         BackColor       =   &H00C8FAC8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1710
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   120
         Width           =   4425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "검사종류"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   165
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   6370
      Left            =   240
      ScaleHeight     =   6315
      ScaleWidth      =   11520
      TabIndex        =   5
      Top             =   1365
      Width           =   11580
      Begin FPSpread.vaSpread SPR031 
         Height          =   6312
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11508
         _Version        =   196608
         _ExtentX        =   20299
         _ExtentY        =   11134
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         MaxCols         =   16
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   8388608
         SpreadDesigner  =   "ANATO118.frx":0000
         UserResize      =   1
         VisibleCols     =   16
         VisibleRows     =   100
      End
   End
   Begin Threed.SSCommand SSclose 
      Height          =   705
      Left            =   10260
      TabIndex        =   4
      Top             =   510
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1244
      _StockProps     =   78
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      AutoSize        =   1
      Picture         =   "ANATO118.frx":3F72
   End
   Begin Threed.SSCommand Ssexec 
      Height          =   705
      Left            =   8640
      TabIndex        =   3
      Top             =   510
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1244
      _StockProps     =   78
      Caption         =   "Clear(&E)"
      ForeColor       =   16711680
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
      Picture         =   "ANATO118.frx":428C
   End
   Begin Threed.SSCommand SsInquiry 
      Height          =   705
      Left            =   7020
      TabIndex        =   1
      Top             =   510
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1244
      _StockProps     =   78
      Caption         =   "조회(&I)"
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
      Picture         =   "ANATO118.frx":45A6
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   " J(자체)W(외부)   "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   8010
      Width           =   1890
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "C(CHAR)   N(NUM)   D(합성)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5760
      TabIndex        =   12
      Top             =   8010
      Width           =   2730
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "C (결과형태) :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4230
      TabIndex        =   11
      Top             =   8010
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "1부터시작함   "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1890
      TabIndex        =   10
      Top             =   7785
      Width           =   1470
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "B (검사구분)  :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   8010
      Width           =   11580
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "A (DiffCount) : "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   7785
      Width           =   11580
   End
End
Attribute VB_Name = "Anato_ItemCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim R_codegu            As String
    Dim R_codeky            As String
    Dim R_codenm            As String
    Dim R_specod            As String * 2
    Dim R_chkflg            As String

    Dim R_rowcnt            As Integer

    Dim cCodeky             As String
    Dim cItemnm             As String
    Dim cItemko             As String
    Dim cYageo              As String
    Dim cSugacd             As String
    Dim cGeomsaTm           As String
    Dim cGeomsaW1           As String
    Dim cGeomsaW2           As String
    Dim cChwhYg             As String
    Dim cGeomChc1           As String
    Dim cGeomChc2           As String
    Dim cChunit             As Integer
    Dim cGeomjan1           As String
    Dim cGeomjan2           As String
    Dim cGeomjan3           As String
    Dim cDanwi              As String
    Dim cMincham            As Integer
    Dim cMaxcham            As Integer
    Dim cMinDanger          As Integer
    Dim cMaxDanger          As Integer
    Dim cGbRoutine          As String
    Dim cGbCheck            As String
    Dim cGbinput            As String
    Dim cDiffCount          As Integer
    Dim cMaxdiffc           As Integer
    Dim cChcomment          As String
    Dim cCgcomment          As String
    Dim cGeomsaGb           As String
    Dim cResultW            As String
    Dim cPanicMin           As Integer
    Dim cPanicMax           As Integer
    Dim cDeltaMin           As Integer
    Dim cDeltaMax           As Integer
    Dim cDeltaQC            As String
    Dim cOrderCD            As String



Private Sub Form_Load()

    R_specod = ""
    
'    Ssreport.Enabled = False
 
'Spread sheet 마지막자리수 안나타나게 하는것
    SPR031.Col = 16
    SPR031.ColHidden = True
    SPR031.CursorStyle = SS_CURSOR_STYLE_ARROW
    
    Call ScreeN_Clear
 
    SpeCod.Clear
    SpeCod.AddItem "85 진단병리"   '91
'    SpeCod.AddItem "62    Histology "
'    SpeCod.AddItem "63    CYTOLOGY "

End Sub

Sub ScreeN_Clear()

 Call SSInitialize(SPR031)
 'CLPMDI.mdimess = " "
 
 SPR031.Row = 0
 SPR031.Col = 0
 SPR031.Action = SS_ACTION_ACTIVE_CELL

End Sub

Private Sub SpeCod_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

    KeyAscii = 0
    SendKeys "{tab}"

End Sub


Private Sub SpeCod_LostFocus()
    
     SPR031.Row = 1
     SPR031.Col = 1
     SPR031.Action = SS_ACTION_ACTIVE_CELL
      
     If SpeCod.ListIndex = -1 Then
        SpeCod.SetFocus
        Exit Sub
     End If
     
End Sub


Private Sub SSclose_Click()

   Unload Me
   
End Sub

Private Sub Ssexec_Click()

 Call ScreeN_Clear
' Ssreport.Enabled = False
 SpeCod.SetFocus

End Sub



Private Sub SsInquiry_Click()

    Dim rs                  As ADODB.Recordset
     
     R_specod = Mid(SpeCod, 1, 2)
     If R_specod = "" Then
        SpeCod.SetFocus
        Exit Sub
     End If

''''GoSub Spread_Clear_Set
    SPR031.Row = 1
    SPR031.Row2 = SPR031.DataRowCnt
    SPR031.Col = 1
    SPR031.Col2 = SPR031.DataColCnt
    SPR031.BlockMode = True
    SPR031.Text = ""
    SPR031.BlockMode = False
    
''''Item Code Db Table 검색
    strSQL = ""
    strSQL = strSQL & " SELECT  Codeky,    Itemnm,    ITemko,    Sugacd,    Danwi, "
    strSQL = strSQL & "         Mincham,   Maxcham,   MinDanger, MaxDanger,        "
    strSQL = strSQL & "         GbRoutine, GbCheck, GbInput, DiffCount, GeomsaGb,  "
    strSQL = strSQL & "         ResultW,   ROWID                                   "
    strSQL = strSQL & " FROM    TWEXAM_ITEMML                                      "
    strSQL = strSQL & " WHERE   SubStr(Codeky,1,2) = '" & R_specod & "'            "
    strSQL = strSQL & " ORDER   BY Codeky                                          "
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
    '       CLPMDI.mdimess = "ITem이 없습니다 신규등록하십시요 .."
        SPR031.SetFocus
        Exit Sub
    End If
    
    Do Until rs.EOF
        SPR031.Row = SPR031.DataRowCnt + 1
        SPR031.Col = 1:  SPR031.Text = rs.Fields("Codeky").Value & ""
        SPR031.Col = 2:  SPR031.Text = rs.Fields("Itemnm").Value & ""
        SPR031.Col = 3:  SPR031.Text = rs.Fields("Itemko").Value & ""
        SPR031.Col = 4:  SPR031.Text = rs.Fields("SugaCd").Value & ""
        SPR031.Col = 5:  SPR031.Text = Trim(rs.Fields("DanWi").Value & "")
        SPR031.Col = 6:  SPR031.Text = rs.Fields("MinCham").Value & ""
        SPR031.Col = 7:  SPR031.Text = rs.Fields("MaxCham").Value & ""
        SPR031.Col = 8:  SPR031.Text = rs.Fields("MinDanger").Value & ""
        SPR031.Col = 9:  SPR031.Text = rs.Fields("MaxDanger").Value & ""
        SPR031.Col = 10: SPR031.Text = rs.Fields("Gbroutine").Value & ""
        SPR031.Col = 11: SPR031.Text = rs.Fields("Gbcheck").Value & ""
        SPR031.Col = 12: SPR031.Text = rs.Fields("GbInput").Value & ""
        
'        SPR031.Col = 13: SPR031.Text = Str(rs.Fields("DiffCount").Value & "")
        
        SPR031.Col = 14: SPR031.Text = Trim(rs.Fields("GeomsaGb").Value & "")
        SPR031.Col = 15: SPR031.Text = Trim(rs.Fields("ResultW").Value & "")
        SPR031.Col = 16: SPR031.Text = rs.Fields("ROWID").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
    '    CLPMDI.mdimess = "내용이있으므로 하시고싶은작업을하세요(수정,삭제,출력)"
    
    SPR031.MaxRows = SPR031.DataRowCnt + 1
    
    SPR031.SetFocus
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  미사용


Private Sub spr031_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim rs                  As ADODB.Recordset
    
    Dim strData             As String
    
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''
    If Row = NewRow Then Exit Sub
    
    If SPR031.Row = 0 Then Exit Sub
    
    SPR031.Col = 1: SPR031.Row = Row
    If SPR031.Text = "" Then Exit Sub
    
''''
''''ITEMML_SELECT
    
    SPR031.Col = 1: SPR031.Row = Row
    strSQL = ""
    strSQL = strSQL & " SELECT    Codeky,    Itemnm,    ITemko,    Sugacd,    Danwi, "
    strSQL = strSQL & "       Mincham,   Maxcham,   MinDanger, MaxDanger,        "
    strSQL = strSQL & "       Gbroutine, gbcheck,  GbInput, DiffCount, GeomsaGb, "
    strSQL = strSQL & "       ResultW "
    strSQL = strSQL & " FROM  TWEXAM_ITEMML "
    strSQL = strSQL & " WHERE Codeky = '" & SPR031.Text & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
       R_chkflg = "Insert"
    Else
       R_chkflg = "Update"
    End If

       AdoCloseSet rs

''''
''''insert or update
    If R_chkflg = "Insert" Then
        SPR031.Row = Row
        Call ITEMML_MOVE
        
        SPR031.Col = 12
        'If Mid(spr031.Text, 1, 1) = "J" Or Mid(spr031.Text, 1, 1) = "W" Then
        '   CLPMDI.mdimess = " "
        'Else
        '    Beep
        '   CLPMDI.mdimess = Row & "  " & "검사구분이잘못되었읍니다"
        '   Exit Sub
        'End If
                                                         
        strSQL = ""
        strSQL = strSQL & " INSERT INTO TWEXAM_ITEMML"
        strSQL = strSQL & "       (  Codeky,       Itemnm,      Itemko,      Yageo,        Sugacd,                          "
        strSQL = strSQL & "          Geomsatm,     Geomsaw1,    Geomsaw2,    Chwhyg,       Geomchc1,                        "
        strSQL = strSQL & "          Geomchc2,     Chunit,      Geomjan1,    Geomjan2,     Geomjan3,                        "
        strSQL = strSQL & "          Danwi,        Resultw,     Mincham,     Maxcham,      MinDanger, MaxDanger, Gbroutine, "
        strSQL = strSQL & "          GbCheck,      Diffcount,   Maxdiffc,    Chcomment,    Cgcomment,                       "
        strSQL = strSQL & "          Geomsagb,     GbInput,    Codate,  "
        strSQL = strSQL & "          DeltaMin,    DeltaMax,             "
        strSQL = strSQL & "          DeltaQC,      OrderCd )            "
        strSQL = strSQL & " VALUES ('" & cCodeky & "',"
        strSQL = strSQL & "         '" & cItemnm & "',"
        strSQL = strSQL & "         '" & cItemko & "',"
        strSQL = strSQL & "         '" & cYageo & "',"
        strSQL = strSQL & "         '" & cSugacd & "',"
        strSQL = strSQL & "         '" & cGeomsaTm & "',"
        strSQL = strSQL & "         '" & cGeomsaW1 & "',"
        strSQL = strSQL & "         '" & cGeomsaW2 & "',"
        strSQL = strSQL & "         '" & cChwhYg & "',"
        strSQL = strSQL & "         '" & cGeomChc1 & "',"
        strSQL = strSQL & "         '" & cGeomChc2 & "',"
        strSQL = strSQL & "         '" & cChunit & "',"
        strSQL = strSQL & "         '" & cGeomjan1 & "',"
        strSQL = strSQL & "         '" & cGeomjan2 & "',"
        strSQL = strSQL & "         '" & cGeomjan3 & "',"
        strSQL = strSQL & "         '" & cDanwi & "',"
        strSQL = strSQL & "         '" & cResultW & "',"
        strSQL = strSQL & "          " & cMincham & ","
        strSQL = strSQL & "          " & cMaxcham & ","
        strSQL = strSQL & "          " & cMinDanger & ","
        strSQL = strSQL & "          " & cMaxDanger & ","
        strSQL = strSQL & "         '" & cGbRoutine & "',"
        strSQL = strSQL & "         '" & cGbCheck & "',"
        strSQL = strSQL & "          " & cDiffCount & "',"
        strSQL = strSQL & "          " & cMaxdiffc & ","
        strSQL = strSQL & "         '" & cChcomment & "',"
        strSQL = strSQL & "         '" & cCgcomment & "',"
        strSQL = strSQL & "         '" & cGeomsaGb & "',"
        strSQL = strSQL & "         '" & cGbinput & "',"
        strSQL = strSQL & "         TO_DATE('SYSDATE','YYYY-MM-DD')"
        strSQL = strSQL & "          " & cDeltaMin & ","
        strSQL = strSQL & "          " & cDeltaMax & ","
        strSQL = strSQL & "         '" & cDeltaQC & "',"
        strSQL = strSQL & "         '" & cOrderCD & "')"
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
    '        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    Else
''''''''GoSub SPR031_UPDATE
        
        SPR031.Row = Row
        
        Call ITEMML_MOVE
        
        SPR031.Col = 12
        'If Mid(spr031.Text, 1, 1) = "J" Or Mid(spr031.Text, 1, 1) = "W" Then
        '   CLPMDI.mdimess = " "
        'Else
        '   Beep
        '   CLPMDI.mdimess = Row & "  " & "검사구분이잘못되었읍니다"
        '   Return
        'End If
        strSQL = ""
        strSQL = strSQL & " UPDATE TWEXAM_ITEMML "
        strSQL = strSQL & " SET   itemnm      = '" & cItemnm & "',"
        strSQL = strSQL & "       Itemko      = '" & cItemko & "',"
        strSQL = strSQL & "       Sugacd      = '" & cSugacd & "',"
        strSQL = strSQL & "       Danwi       = '" & cDanwi & "',"
        strSQL = strSQL & "       Mincham     =  " & cMincham & ","
        strSQL = strSQL & "       Maxcham     =  " & cMaxcham & ","
        strSQL = strSQL & "       MinDanger   =  " & cMinDanger & ","
        strSQL = strSQL & "       MaxDanger   =  " & cMaxDanger & ","
        strSQL = strSQL & "       GbRoutine   = '" & cGbRoutine & "',"
        strSQL = strSQL & "       GbCheck     = '" & cGbCheck & "',"
        strSQL = strSQL & "       GbInput     = '" & cGbinput & "',"
        strSQL = strSQL & "       DiffCount   =  " & cDiffCount & ","
        strSQL = strSQL & "       Geomsagb    = '" & cGeomsaGb & "',"
        strSQL = strSQL & "       ResultW     = '" & cResultW & "'"
        strSQL = strSQL & " WHERE Codeky      =  '" & cCodeky & "'"
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
    '        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    End If

End Sub

Sub ITEMML_MOVE()
    

    SPR031.Col = 1: cCodeky = SPR031.Text
    SPR031.Col = 2: cItemnm = SPR031.Text
    SPR031.Col = 3: cItemko = SPR031.Text
    cYageo = ""
    
    SPR031.Col = 4: cSugacd = SPR031.Text
    
    
    cGeomsaTm = ""
    cGeomsaW1 = ""
    cGeomsaW2 = ""
    cChwhYg = ""
    cGeomChc1 = ""
    cGeomChc2 = ""
    cChunit = 0
    cGeomjan1 = ""
    cGeomjan2 = ""
    cGeomjan3 = ""
    SPR031.Col = 5: cDanwi = SPR031.Text
    
   
    SPR031.Col = 6
    If SPR031.Text = "" Then
       cMincham = 0
       cMinDanger = 0
    Else
       cMincham = Val(SPR031.Text)
       cMinDanger = Val(SPR031.Text)
    End If
    
    SPR031.Col = 7
    If SPR031.Text = "" Then
       cMaxcham = 0
       cMaxDanger = 0
    Else
       cMaxcham = Val(SPR031.Text)
       cMaxDanger = Val(SPR031.Text)
    End If
    
    SPR031.Col = 8:
    If SPR031.Text <> "" Then
        cMinDanger = Val(SPR031.Text)
    End If
    
    SPR031.Col = 9:
    If SPR031.Text <> "" Then
        cMaxDanger = Val(SPR031.Text)
    End If
    
    SPR031.Col = 10:  cGbRoutine = Val(SPR031.Text)
    SPR031.Col = 11:  cGbCheck = Val(SPR031.Text)
    SPR031.Col = 12:  cGbinput = Trim(SPR031.Text)
    SPR031.Col = 13:  cDiffCount = Val(SPR031.Text)
    
    cMaxdiffc = 0
    cChcomment = ""
    cCgcomment = ""
    
    SPR031.Col = 14: cGeomsaGb = Trim(SPR031.Text)
    SPR031.Col = 15: cResultW = Trim(SPR031.Text)
    
    cDeltaMin = 0
    cDeltaMax = 0
    cDeltaQC = ""
    cOrderCD = ""
    
End Sub

Private Sub spr031_Click(ByVal Col As Long, ByVal Row As Long)

    SPR031.Row = Row: SPR031.Col = Col
    SPR031.Action = SS_ACTION_ACTIVE_CELL
    
End Sub

Private Sub spr031_DblClick(ByVal Col As Long, ByVal Row As Long)
     
     Dim Messbox            As String
     Dim Style              As String
     Dim Title              As String
     Dim MessRes            As String
     Dim mystring           As String
     
    Exit Sub
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Message Box 사용법
     Messbox = "Item Code를 삭제하시겠읍니까?"
     Style = vbYesNo + vbCritical + vbDefaultButton2
     Title = "Item Code"
     MessRes = MsgBox(Messbox, Style, Title)
''     If MessRes = IDYes Then
     If MessRes = vbYes Then
         mystring = "Yes"
     Else
         mystring = "No"
     End If
    
     If mystring = "No" Then
        Call ScreeN_Clear
      '  Ssreport.Enabled = False
        SpeCod.SetFocus
        'CLPMDI.mdimess = "다시작업하십시요 ..."
        Exit Sub
     End If
    
    
     GoSub DELETE_ROUTINE
     SPR031.Action = SS_ACTION_DELETE_ROW
     SPR031.Action = SS_ACTION_ACTIVE_CELL
     
'    CLPMDI.mdimess = "삭제가되었으므로다른작업을하십시요 .."
       
Exit Sub
    


DELETE_ROUTINE:
    Dim sRowID              As String
    
    SPR031.Col = 16
    strSQL = " DELETE FROM TWEXAM_ITEMML WHERE ROWID = '" & SPR031.Text & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return


End Sub


Private Sub spr031_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

    KeyAscii = 0

End Sub


