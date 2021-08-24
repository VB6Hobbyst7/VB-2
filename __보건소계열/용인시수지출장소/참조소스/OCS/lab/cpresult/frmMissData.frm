VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMissData 
   Caption         =   "미확인 결과Data"
   ClientHeight    =   7635
   ClientLeft      =   330
   ClientTop       =   945
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11625
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel3 
      Height          =   6135
      Left            =   6030
      TabIndex        =   17
      Top             =   1125
      Width           =   4965
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   10821
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
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00800000&
         ForeColor       =   &H80000005&
         Height          =   465
         Left            =   585
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   19
         Top             =   5580
         Width           =   4245
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   285
         Left            =   270
         TabIndex        =   18
         Top             =   225
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "▼ 모두선택"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin FPSpreadADO.fpSpread sprData 
         Height          =   5010
         Left            =   225
         TabIndex        =   20
         Top             =   540
         Width           =   4605
         _Version        =   196608
         _ExtentX        =   8123
         _ExtentY        =   8837
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmMissData.frx":0000
         UserResize      =   1
         Appearance      =   2
      End
      Begin MSForms.CommandButton cmdSet 
         Height          =   420
         Left            =   225
         TabIndex        =   22
         Top             =   5580
         Width           =   330
         Caption         =   "Remark"
         PicturePosition =   327683
         Size            =   "582;741"
         Picture         =   "frmMissData.frx":1948
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsert 
         Height          =   465
         Left            =   2790
         TabIndex        =   21
         Top             =   45
         Width           =   1995
         Caption         =   "결과입력"
         PicturePosition =   327683
         Size            =   "3519;820"
         Picture         =   "frmMissData.frx":1C62
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel panelPr 
      Height          =   2130
      Left            =   3960
      TabIndex        =   12
      Top             =   7245
      Visible         =   0   'False
      Width           =   3840
      _Version        =   65536
      _ExtentX        =   6773
      _ExtentY        =   3757
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
      Begin VB.CommandButton cmdExit 
         Caption         =   "종료"
         Height          =   420
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   1815
      End
      Begin FPSpreadADO.fpSpread sprPr 
         Height          =   5505
         Left            =   90
         TabIndex        =   13
         Top             =   765
         Width           =   11310
         _Version        =   196608
         _ExtentX        =   19950
         _ExtentY        =   9710
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frmMissData.frx":3424
         UserResize      =   1
         Appearance      =   1
      End
      Begin MSForms.CommandButton cmdPrExe 
         Height          =   420
         Left            =   2070
         TabIndex        =   15
         Top             =   270
         Width           =   1680
         Caption         =   "출력확인"
         Size            =   "2963;741"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin VB.TextBox txtSLipno2 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   540
      Width           =   645
   End
   Begin VB.TextBox txtJeobsuDt 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   8685
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   540
      Width           =   1185
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   7290
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   540
      Width           =   1365
   End
   Begin FPSpreadADO.fpSpread sprPtList 
      Height          =   5280
      Left            =   90
      TabIndex        =   7
      Top             =   1845
      Width           =   5820
      _Version        =   196608
      _ExtentX        =   10266
      _ExtentY        =   9313
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
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmMissData.frx":4E0B
      UserResize      =   1
      Appearance      =   2
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   810
      TabIndex        =   5
      Top             =   1440
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   24510467
      CurrentDate     =   36535
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   810
      TabIndex        =   4
      Top             =   1080
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   24510467
      CurrentDate     =   36535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11025
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMissData.frx":678B
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   635
      ButtonWidth     =   1270
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Description     =   "Exit of Form"
            Object.ToolTipText     =   "Exit of Form"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   465
      Left            =   3060
      TabIndex        =   1
      Top             =   405
      Width           =   4200
      _Version        =   65536
      _ExtentX        =   7408
      _ExtentY        =   820
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
      RoundedCorners  =   0   'False
      MouseIcon       =   "frmMissData.frx":6AA7
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   90
         Width           =   2580
      End
      Begin VB.Label Label9 
         Caption         =   "검사종목:"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   3
         Top             =   135
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmMissData.frx":7D49
         Stretch         =   -1  'True
         Top             =   0
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   330
      Left            =   135
      TabIndex        =   16
      Top             =   405
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "미확인 결과 Data 결과확인"
      ForeColor       =   16777215
      BackColor       =   128
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
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   1485
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   1170
      Width           =   465
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   465
      Left            =   4365
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
      Caption         =   "Data출력"
      PicturePosition =   327683
      Size            =   "2566;820"
      Picture         =   "frmMissData.frx":8FDB
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   465
      Left            =   2925
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2566;820"
      Picture         =   "frmMissData.frx":98B5
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmMissData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSLip_Click()
    
    Call SpreadSetClear(sprPtList)
    Call SpreadSetClear(sprData)
    
    For i = 1 To Me.sprData.DataRowCnt
        sprData.Row = i
        sprData.Col = 1
        sprData.Value = False
        cmdSelect.Caption = "▼ 모두선택"
    Next
    txtStatus.Text = ""
    txtJeobsuDt.Text = ""
    txtSLipno2.Text = ""
    
    
End Sub

Private Sub cmdExit_Click()
    
    panelPr.Visible = False
    
End Sub

Private Sub cmdInsert_Click()
    Dim iCheckCount     As Integer
    Dim iResultCount    As Integer
    Dim sStatus         As String
    
    
    If sprData.DataRowCnt = 0 Then Exit Sub
    
    
    For i = 1 To sprData.DataRowCnt
        sprData.Row = i
        sprData.Col = 1
        If sprData.Value = True Then iCheckCount = iCheckCount + 1
        sprData.Col = 5
        If Trim(sprData.Text) <> "" Then iResultCount = iResultCount + 1
    Next
    
    If iCheckCount = 0 Then Exit Sub            'Check 된것이 하나도 없을때.......
    If iResultCount = 0 Then Exit Sub           '결과 입력이 하나도 없을때........
    
    If iCheckCount < sprData.DataRowCnt Then sStatus = "P"          'Check 된것이 모든행보다 작을때
    If sprData.DataRowCnt = iResultCount Then sStatus = "C"         '결과완료
    If sprData.DataRowCnt = sprData.DataRowCnt Then sStatus = "C"   '전부 Check 하였을 경우
        
        
    GoSub Main_Job_Verify_Sub
    Exit Sub
    


Main_Job_Verify_Sub:
    Dim sRowID          As String
    
    For i = 1 To sprData.DataRowCnt
        sprData.Row = i
        sprData.Col = 2: sRowID = sprData.Text
        
        strSql = ""
        strSql = strSql & " UPDATE TWEXAM_GENERAL_SUB"
        strSql = strSql & " SET    Verify  =  'Y'"
        strSql = strSql & " WHERE  ROWID   = '" & sRowID & "'"
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
                
        strSql = ""
        strSql = strSql & " UPDATE TWEXAM_General"
        strSql = strSql & " SET    Status  = '" & sStatus & "'"
        strSql = strSql & " WHERE  JeobsuDt  = TO_DATE('" & txtJeobsuDt.Text & "','yyyy-MM-dd')"
        strSql = strSql & " AND    SLipno1   = " & Val(Left(cmbSLip.Text, 2))
        strSql = strSql & " AND    SLipno2   = " & Val(txtSLipno2.Text)
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
        
    Next
    Return
    
    
    
End Sub

Private Sub cmdPrExe_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    If sprPr.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "결과 미확인 Data 접수내역 LIST"
    
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFrDate.Value, "yyyy-MM-dd") & " / " & _
                                               Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprPr.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                        strFont(1) + "/n" + sBarLine + strFont(1)
    sprPr.PrintFooter = "/f2" & "/l" & sBarLine & _
                        "/n" & Space(80) & "Page : " & "/p" & " of " & sprPr.PrintPageCount
    sprPr.PrintMarginLeft = 0
    sprPr.PrintMarginRight = 0
    sprPr.PrintMarginTop = 0
    sprPr.PrintMarginBottom = 0
    sprPr.PrintColHeaders = True
    sprPr.PrintRowHeaders = False
    sprPr.PrintBorder = False
    sprPr.PrintColor = True
    sprPr.PrintGrid = True
    sprPr.PrintShadows = True
    sprPr.PrintUseDataMax = False
    sprPr.Row = 1
    sprPr.Row2 = sprPr.DataRowCnt
    sprPr.Col = 1
    sprPr.Col2 = sprPr.MaxCols
    sprPr.PrintOrientation = PrintOrientationPortrait
    sprPr.PrintType = PrintTypeCellRange
    sprPr.Action = ActionPrint
    
    panelPr.Visible = False
    

End Sub

Private Sub cmdPrint_Click()
    Dim sFrDate     As String
    Dim sToDate     As String
    Dim sCompare    As String
    
    
    panelPr.Top = 945
    panelPr.Left = 90
    panelPr.Height = 6550
    panelPr.Width = 11445
    panelPr.Visible = True
    panelPr.ZOrder 0
    
    DoEvents:
    sprPr.ReDraw = False
    sprPr.MaxRows = 0
    sprPr.MaxRows = 500
    sprPr.RowHeight(-1) = 11
    sprPr.ReDraw = True
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT TO_CHAR(b.JeobsuDt,'yyyy-MM-dd') JeobsuDt, "
    strSql = strSql & "        b.SLipno2, b.Ptno, c.Sname, c.Sex || '/' || c.AgeYY SA, a.DeptCode, a.Status, b.Result1,"
    strSql = strSql & "        b.ItemCd, d.ItemNM"
    strSql = strSql & " FROM   TWEXAM_GENERAL     a,"
    strSql = strSql & "        TWEXAM_General_Sub b,"
    strSql = strSql & "        TWEXAM_IDNOMST     c,"
    strSql = strSql & "        TWEXAM_ITEMML      d,"
    strSql = strSql & "        TWEXAM_ORDER       e "
    strSql = strSql & " WHERE  a.JeobsuDt  >=  TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt  <=  TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.GBCh      =  'Y'"
    strSql = strSql & " AND    a.JeobsuDt  =   b.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   =   b.SLipno1(+)"
    strSql = strSql & " AND    a.SLipno2   =   b.SLipno2(+)"
    strSql = strSql & " AND    a.JeobsuDt  =   e.CollDate(+)"
    strSql = strSql & " AND    a.Matchno   =   e.Matchno(+)"
    strSql = strSql & " AND    e.JeobsuYN !=   '#'"
    strSql = strSql & " AND    a.SLipno1   =   '" & Left(cmbSLip.Text, 2) & "'"
    strSql = strSql & " AND    b.Verify    =   'N'"
    strSql = strSql & " AND    a.Ptno      =   c.Ptno(+)"
    strSql = strSql & " AND    b.ItemCd    =   d.Codeky(+)"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprPr.Row = sprPr.DataRowCnt + 1
        
        If sCompare <> adoSet.Fields("JeobsuDt").Value & "" & _
                       adoSet.Fields("SLipno2").Value & "" & _
                       adoSet.Fields("Ptno").Value & "" & _
                       adoSet.Fields("Sname").Value & "" & _
                       adoSet.Fields("Sa").Value & "" & _
                       adoSet.Fields("DeptCode").Value & "" & _
                       adoSet.Fields("Status").Value & "" Then
                       
                
            sprPr.Col = 1: sprPr.Text = adoSet.Fields("JeobsuDt").Value & ""
            sprPr.Col = 2: sprPr.Text = adoSet.Fields("SLipno2").Value & ""
            sprPr.Col = 3: sprPr.Text = adoSet.Fields("Ptno").Value & ""
            sprPr.Col = 4: sprPr.Text = adoSet.Fields("Sname").Value & ""
            sprPr.Col = 5: sprPr.Text = adoSet.Fields("Sa").Value & ""
            sprPr.Col = 6: sprPr.Text = adoSet.Fields("DeptCode").Value & ""
            sprPr.Col = 7: sprPr.Text = adoSet.Fields("Status").Value & ""
            
            sprPr.Col = 1:         sprPr.Col2 = sprPr.MaxCols
            sprPr.Row = sprPr.Row: sprPr.Row2 = sprPr.Row
            sprPr.BlockMode = True
            sprPr.CellBorderType = SS_BORDER_TYPE_TOP
            sprPr.CellBorderStyle = CellBorderStyleSolid
            sprPr.Action = ActionSetCellBorder
            sprPr.BlockMode = False
            
        End If
        sprPr.Col = 8: sprPr.Text = adoSet.Fields("ItemCd").Value & ""
        sprPr.Col = 9: sprPr.Text = adoSet.Fields("ItemNM").Value & ""
        sprPr.Col = 10: sprPr.Text = adoSet.Fields("Result1").Value & ""
        
        sCompare = adoSet.Fields("JeobsuDt").Value & "" & _
                      adoSet.Fields("SLipno2").Value & "" & _
                      adoSet.Fields("Ptno").Value & "" & _
                      adoSet.Fields("Sname").Value & "" & _
                      adoSet.Fields("Sa").Value & "" & _
                      adoSet.Fields("DeptCode").Value & "" & _
                      adoSet.Fields("Status").Value & ""
        
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate     As String
    Dim sToDate     As String
    
    Call SpreadSetClear(sprPtList)
    Call SpreadSetClear(sprData)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd HH:mm")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd HH:mm")
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(b.JeobsuDt,'yyyy-MM-dd') JeobsuDt, "
    strSql = strSql & "        b.SLipno2, b.Ptno, c.Sname, c.Sex || '/' || c.AgeYY SA, a.DeptCode, a.Status,"
    strSql = strSql & "        COUNT(b.ItemCd) Count"
    strSql = strSql & " FROM   TWEXAM_GENERAL     a,"
    strSql = strSql & "        TWEXAM_General_Sub b,"
    strSql = strSql & "        TWEXAM_IDNOMST     c,"
    strSql = strSql & "        TWEXAM_ORDER       d "
    strSql = strSql & " WHERE  LTRIM(TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD')) || ' ' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.JeobsuT1, '00')) || ':' || "
    strSql = strSql & "        LTRIM(TO_CHAR(a.JeobsuT2, '00'))   BETWEEN  '" & sFrDate & "'"
    strSql = strSql & "                                           AND      '" & sToDate & "'"
    strSql = strSql & " AND    a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & " AND    a.GBCh     = 'Y'"
    strSql = strSql & " AND    a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & " AND    a.SLipno2  = b.SLipno2(+)"
    strSql = strSql & " AND    a.JeobsuDt = d.CollDate(+)"
    strSql = strSql & " AND    a.Matchno  = d.Matchno(+)"
    strSql = strSql & " AND    d.JeobsuYn != '#'"
    strSql = strSql & " AND    a.SLipno1  = '" & Left(cmbSLip.Text, 2) & "'"
    strSql = strSql & " AND    b.Verify   = 'N'"
    strSql = strSql & " AND    a.Ptno     = c.Ptno(+)"
    strSql = strSql & " GROUP BY  b.JeobsuDt, b.SLipno2, b.Ptno, c.Sname, c.Sex || '/' || c.AgeYY, a.DeptCode, a.Status"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprPtList.Row = sprPtList.DataRowCnt + 1
        sprPtList.Col = 1: sprPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprPtList.Col = 2: sprPtList.Text = adoSet.Fields("SLipno2").Value & ""
        sprPtList.Col = 3: sprPtList.Text = adoSet.Fields("Ptno").Value & ""
        sprPtList.Col = 4: sprPtList.Text = adoSet.Fields("Sname").Value & ""
        sprPtList.Col = 5: sprPtList.Text = adoSet.Fields("Sa").Value & ""
        sprPtList.Col = 6: sprPtList.Text = adoSet.Fields("DeptCode").Value & ""
        sprPtList.Col = 7: sprPtList.Text = adoSet.Fields("Count").Value & ""
        sprPtList.Col = 8: sprPtList.Text = adoSet.Fields("Status").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdSelect_Click()
    
    If cmdSelect.Caption = "▼ 모두선택" Then
        For i = 1 To Me.sprData.DataRowCnt
            sprData.Row = i
            sprData.Col = 1
            sprData.Value = True
            cmdSelect.Caption = "▼ 모두해제"
        Next
    Else
        For i = 1 To Me.sprData.DataRowCnt
            sprData.Row = i
            sprData.Col = 1
            sprData.Value = False
            cmdSelect.Caption = "▼ 모두선택"
        Next
    End If

End Sub

Private Sub cmdSet_Click()

    
    hWndReturn = Me.txtRemark.hwnd
    gSRmkSLipno = Left(Me.cmbSLip.Text, 2)
    clpRemark.Show vbModal
    gSRmkSLipno = ""
    
End Sub

Private Sub Form_Load()
    
    
    'dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    'dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd") & " 00:01"
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd") & " 23:59"
    
    GoSub SLip_Select
    
    GiExamNumb = Val(GetSetting("CP", "CPRESULT", "SLip"))
    
    Call SetComboBox(cmbSLip, GiExamNumb, 2)
    Exit Sub
    
'/--------------------------------------------------------

FORMClear:
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    Return
    
    
SLip_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '90'"
'C    strSql = strSql & " AND    Codeky < '52'"
    strSql = strSql & " ORDER  BY Codeky"
    
    cmbSLip.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return

End Sub

Private Sub sprPtList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sJeobsuDt           As String
    Dim sSLip1              As String
    Dim sSLip2              As String
    Dim sItemCd             As String
    
    Call SpreadSetClear(sprData)
    txtStatus.Text = ""
    
    If Row = 0 Then Exit Sub
    
    sprPtList.Row = Row
    sprPtList.Col = 1: sJeobsuDt = sprPtList.Text
    txtJeobsuDt.Text = sJeobsuDt
    
    sSLip1 = Left(cmbSLip, 2)
    sprPtList.Col = 2: sSLip2 = sprPtList.Text
    txtSLipno2.Text = sSLip2
    
    sprPtList.Col = 8
    Select Case Trim(sprPtList.Text)
        Case "C": txtStatus.Text = "결과완료"
        Case "R": txtStatus.Text = "접수중"
        Case "P": txtStatus.Text = "부분결과"
        Case "U": txtStatus.Text = "미확인"
        Case "X": txtStatus.Text = "ABNormal"
        Case Else: txtStatus.Text = ""
    End Select
    
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.rowid RWID, b.ItemNM,"
    strSql = strSql & "        c.m_min, c.m_max, c.f_min, c.f_max "
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_ITEMML      b,"
    strSql = strSql & "        TWEXAM_REFDATA     c "
    strSql = strSql & " WHERE  a.Jeobsudt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Verify   = 'N'"
    strSql = strSql & " AND    a.SLipno1  = " & Val(sSLip1)
    strSql = strSql & " AND    a.SLipno2  = " & Val(sSLip2)
    strSql = strSql & " AND    a.ItemCd   =  b.Codeky(+)"
    strSql = strSql & " AND    a.ItemCD   =  c.ItemCode(+)"
    strSql = strSql & " AND    a.AgeYY   >=  c.AgeMin(+)"
    strSql = strSql & " AND    a.AgeYY   <=  c.AgeMax(+)"
    strSql = strSql & " AND    NVL(c.appdate,SysDate) = "
    strSql = strSql & "           (Select NVL(MAX(APPDATE), SysDate)"
    strSql = strSql & "            From   TWEXAM_REFDATA d"
    strSql = strSql & "            Where  d.ItemCode = a.ItemCD"
    strSql = strSql & "            And    d.AgeMin  <= a.AgeYY"
    strSql = strSql & "            And    d.AgeMax  >= a.AgeYY)"
    strSql = strSql & "  ORDER  BY  a.ItemCd  "
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        sprData.Row = sprData.DataRowCnt + 1
        sprData.Col = 1: sprData.Value = False
        sprData.Col = 2: sprData.Text = adoSet.Fields("RWID").Value & ""
        sprData.Col = 3: sprData.Text = adoSet.Fields("ItemCD").Value & ""
        sprData.Col = 4: sprData.Text = adoSet.Fields("ItemNm").Value & ""
        sprData.Col = 5: sprData.Text = adoSet.Fields("Result1").Value & ""
        GoSub RET_Setting_Init1 '결과Data Setting
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Exit Sub
    
RET_Setting_Init1:
    Dim sResult     As String
    
    sprData.Col = 5:
    sprData.CellType = SS_CELL_TYPE_EDIT
    sprData.TypeHAlign = SS_CELL_H_ALIGN_LEFT
    sprData.TypeEditCharCase = SS_CELL_EDIT_CASE_NO_CASE
    sprData.TypeEditMultiLine = False
    sprData.TypeEditLen = 50

    sprData.Col = 3: sItemCd = sprData.Text
    sResult = Get_Result_Text(sItemCd)
    If Trim(sResult) <> "" Then
        sprData.Col = 5
        sprData.CellType = CellTypeComboBox
        sprData.TypeComboBoxList = sResult
        sprData.TypeComboBoxEditable = True
    End If
    
    Return
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
