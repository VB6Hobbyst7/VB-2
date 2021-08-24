VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPrItem 
   Caption         =   "Print SLip"
   ClientHeight    =   6990
   ClientLeft      =   210
   ClientTop       =   1320
   ClientWidth     =   11505
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
   ScaleHeight     =   6990
   ScaleWidth      =   11505
   Begin VB.TextBox txtSLipName 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   1035
      TabIndex        =   2
      Top             =   270
      Width           =   2445
   End
   Begin VB.TextBox txtSLipno1 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   225
      TabIndex        =   1
      Top             =   270
      Width           =   780
   End
   Begin FPSpreadADO.fpSpread sprItemPr 
      Height          =   5955
      Left            =   225
      TabIndex        =   0
      Top             =   840
      Width           =   11130
      _Version        =   196608
      _ExtentX        =   19632
      _ExtentY        =   10504
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
      MaxCols         =   15
      ScrollBars      =   2
      SpreadDesigner  =   "frmPrItem.frx":0000
      UserResize      =   1
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   420
      Left            =   3555
      TabIndex        =   3
      Top             =   270
      Width           =   1455
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "2566;741"
      Picture         =   "frmPrItem.frx":1A95
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
Attribute VB_Name = "frmPrItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPr_Click()
    If sprItemPr.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("RoutineCode  Data 의 Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & Trim(txtSLipName.Text)
    strHead(1) = "/f2" & "Page : " & "/p" & " of " & sprItemPr.PrintPageCount & "/r"
    
    sprItemPr.PrintHeader = strFont(0) + strHead(0) + strFont(1) + strHead(1) + "/n" + strFont(1)
    sprItemPr.PrintFooter = "/f2" & "/c" & "Page : " & "/p" & " of " & sprItemPr.PrintPageCount - 1
    sprItemPr.PrintMarginLeft = 0
    sprItemPr.PrintMarginRight = 0
    sprItemPr.PrintMarginTop = 100
    sprItemPr.PrintMarginBottom = 300
    sprItemPr.PrintColHeaders = True
    sprItemPr.PrintRowHeaders = True
    sprItemPr.PrintBorder = True
    sprItemPr.PrintColor = True
    sprItemPr.PrintGrid = False
    sprItemPr.PrintShadows = True
    sprItemPr.PrintUseDataMax = False
    sprItemPr.Row = 1
    sprItemPr.Row2 = sprItemPr.DataRowCnt
    sprItemPr.Col = 1
    sprItemPr.Col2 = sprItemPr.MaxCols
    sprItemPr.PrintType = PrintTypeCellRange
    sprItemPr.PrintOrientation = PrintOrientationLandscape
    sprItemPr.Action = ActionPrint

End Sub

Private Sub Form_Load()
    
    Me.txtSLipno1.Text = frmItemCode.txtSlipno.Text
    Me.txtSLipName.Text = frmItemCode.txtSLipName.Text
    
        
    strSql = ""
    strSql = strSql & " SELECT a.Codeky, a.ItemNM, a.Yageo,"
    strSql = strSql & "        b.Codenm,"
    strSql = strSql & "        a.Danwi,  a.ResultW, a.GeomsaGB, a.DeltaQc, a.Deltamin, a.Deltamax,"
    strSql = strSql & "        a.Panicmin, a.Panicmax, a.BarText, a.BarGb, a.Exgb"
    strSql = strSql & " FROM   TWEXAM_ITEMML a,"
    strSql = strSql & "        TWEXAM_SAMPLE b"
    strSql = strSql & " WHERE  a.Codeky Like '" & txtSLipno1.Text & "%'"
    strSql = strSql & " AND    a.GeomchC1 = b.Code(+)"
    strSql = strSql & " ORDER  BY a.Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprItemPr.Row = sprItemPr.DataRowCnt + 1
        sprItemPr.Col = 1:  sprItemPr.Text = adoSet.Fields("Codeky").Value & ""
        sprItemPr.Col = 2:  sprItemPr.Text = adoSet.Fields("ItemNM").Value & ""
        sprItemPr.Col = 3:  sprItemPr.Text = adoSet.Fields("Yageo").Value & ""
        sprItemPr.Col = 4:  sprItemPr.Text = adoSet.Fields("Codenm").Value & ""
        sprItemPr.Col = 5:  sprItemPr.Text = adoSet.Fields("Danwi").Value & ""
        sprItemPr.Col = 6:  sprItemPr.Text = adoSet.Fields("Resultw").Value & ""
        sprItemPr.Col = 7:  sprItemPr.Text = adoSet.Fields("GeomsaGb").Value & ""
        sprItemPr.Col = 8:  sprItemPr.Text = adoSet.Fields("DeltaQC").Value & ""
        sprItemPr.Col = 9:  sprItemPr.Text = adoSet.Fields("Deltamin").Value & ""
        sprItemPr.Col = 10: sprItemPr.Text = adoSet.Fields("Deltamax").Value & ""
        sprItemPr.Col = 11: sprItemPr.Text = adoSet.Fields("Panicmin").Value & ""
        sprItemPr.Col = 12: sprItemPr.Text = adoSet.Fields("Panicmax").Value & ""
        sprItemPr.Col = 13: sprItemPr.Text = adoSet.Fields("BarText").Value & ""
        sprItemPr.Col = 14: sprItemPr.Text = adoSet.Fields("BarGB").Value & ""
        sprItemPr.Col = 15: sprItemPr.Text = adoSet.Fields("EXgb").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
