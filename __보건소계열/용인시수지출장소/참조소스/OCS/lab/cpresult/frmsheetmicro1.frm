VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSheetMicro1 
   Caption         =   "미생물WorkSheet"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   1095
   ClientWidth     =   11790
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
   ScaleHeight     =   7590
   ScaleWidth      =   11790
   WindowState     =   2  '최대화
   Begin VB.TextBox txtRemark 
      BackColor       =   &H00FFC0C0&
      Height          =   825
      Left            =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   9
      Top             =   6615
      Width           =   10680
   End
   Begin VB.ComboBox cmbSample 
      Height          =   300
      Left            =   1440
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   855
      Width           =   2985
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   635
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Description     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "검체접수"
            Key             =   "micro"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread sprMicro 
      Height          =   5370
      Left            =   450
      TabIndex        =   1
      Top             =   1260
      Width           =   10680
      _Version        =   196608
      _ExtentX        =   18838
      _ExtentY        =   9472
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      BackColorStyle  =   1
      ColsFrozen      =   9
      DisplayColHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   20
      MaxRows         =   502
      RowHeaderDisplay=   0
      RowsFrozen      =   2
      ScrollBars      =   2
      SpreadDesigner  =   "frmSheetMicro1.frx":0000
      UserResize      =   1
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   1440
      TabIndex        =   2
      Top             =   450
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24510467
      CurrentDate     =   36566
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   2925
      TabIndex        =   3
      Top             =   450
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24510467
      CurrentDate     =   36566
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10980
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":1628A
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":165A6
            Key             =   "Diff"
            Object.Tag             =   "Diff"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":168CA
            Key             =   "FuncS"
            Object.Tag             =   "FuncS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":16D1E
            Key             =   "Clear"
            Object.Tag             =   "Clear"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":184B2
            Key             =   "FuncG"
            Object.Tag             =   "FuncG"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":18906
            Key             =   "SLip"
            Object.Tag             =   "SLip"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":19BAA
            Key             =   "Virus"
            Object.Tag             =   "Virus"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":1A486
            Key             =   "UrineCup"
            Object.Tag             =   "UrineCup"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":1A7A2
            Key             =   "Urine"
            Object.Tag             =   "Urine"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":1AABE
            Key             =   "QryPt"
            Object.Tag             =   "QtyPt"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSheetMicro1.frx":1AF12
            Key             =   "micro"
            Object.Tag             =   "micro"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "검체선택"
      Height          =   240
      Left            =   495
      TabIndex        =   7
      Top             =   900
      Width           =   870
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   555
      Left            =   8325
      TabIndex        =   6
      Top             =   540
      Width           =   1500
      Caption         =   "출력"
      PicturePosition =   327683
      Size            =   "2646;979"
      Picture         =   "frmSheetMicro1.frx":1B7EE
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   555
      Left            =   6930
      TabIndex        =   5
      Top             =   540
      Width           =   1410
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2487;979"
      Picture         =   "frmSheetMicro1.frx":1C0C8
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "접수일자:"
      Height          =   195
      Left            =   495
      TabIndex        =   4
      Top             =   495
      Width           =   825
   End
End
Attribute VB_Name = "frmSheetMicro1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    If sprMicro.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "미생물 WORKSHEET"
    
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFrDate.Value, "yyyy-MM-dd") & " / " & _
                                               Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprMicro.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                        strFont(1) + "/n" + sBarLine + strFont(1)
    sprMicro.PrintFooter = "/f2" & "/l" & sBarLine & _
                        "/n" & Space(60) & "출력일시: " & Dual_Date_Get("yyyy-MM-dd hh24:mi") & _
                        "      Page : " & "/p" & " of " & sprMicro.PrintPageCount
                        
    sprMicro.PrintMarginLeft = 0
    sprMicro.PrintMarginRight = 0
    sprMicro.PrintMarginTop = 0
    sprMicro.PrintMarginBottom = 0
    sprMicro.PrintColHeaders = False
    sprMicro.PrintRowHeaders = True
    sprMicro.PrintBorder = False
    sprMicro.PrintColor = True
    sprMicro.PrintGrid = True
    sprMicro.PrintShadows = True
    sprMicro.PrintUseDataMax = False
    sprMicro.Row = 1
    sprMicro.Row2 = sprMicro.DataRowCnt
    sprMicro.Col = 1
    sprMicro.Col2 = sprMicro.MaxCols
    sprMicro.PrintOrientation = PrintOrientationPortrait
    sprMicro.PrintType = PrintTypeCellRange
    sprMicro.Action = ActionPrint
    
    
    

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate             As String
    Dim sToDate             As String
    Dim sWhere              As String
    

    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprMicro.Row = 3
    sprMicro.Row2 = sprMicro.DataRowCnt
    sprMicro.Col = 0
    sprMicro.Col2 = sprMicro.DataColCnt
    sprMicro.BlockMode = True
    sprMicro.Action = ActionClearText
    sprMicro.BlockMode = False
    
    GoSub Get_Micro_TotalData
    
    Exit Sub
    
    
    
Get_Micro_TotalData:
    strSql = ""
    strSql = strSql & "  SELECT JeobsuDt1, SLipno, Labno, Samplename, Ptno,  Sname, Sx, Age, DeptCode, MSeq,"
    strSql = strSql & "         MAX(Decode(ItemCode, '420201', '*', "
    strSql = strSql & "                              '420202', '*',"
    strSql = strSql & "                              '420203', '*',"
    strSql = strSql & "                              '420204', '*',"
    strSql = strSql & "                              '420205', '*', '')) RoutineCulture,"
    strSql = strSql & "         MAX(Decode(ItemCode, '440201', '*', "
    strSql = strSql & "                              '440202', '*', "
    strSql = strSql & "                              '440203', '*', '')) AFBCulture,"
    strSql = strSql & "         MAX(Decode(ItemCode, '410103', '*', '')) FungusCulture,"
    strSql = strSql & "         MAX(Decode(ItemCode, '420101', '*', '')) GramStain,"
    strSql = strSql & "         MAX(Decode(ItemCode, '440101', '*', '')) AFBStain,"
    strSql = strSql & "         ''                                       FungusStain,"
    strSql = strSql & "         MAX(Decode(ItemCode, '420501', '*', '')) WetSmear,"
    strSql = strSql & "         MAX(Decode(ItemCode, '410101', '*', '')) IndiaInk,"
    strSql = strSql & "         MAX(Decode(ItemCode, '420503', '*',"
    strSql = strSql & "                              '420703', '*',"
    strSql = strSql & "                              '420704', '*',"
    strSql = strSql & "                              '420801', '*',"
    strSql = strSql & "                              '420802', '*',"
    strSql = strSql & "                              '420803', '*',"
    strSql = strSql & "                              '420804', '*', '')) Special,"
    strSql = strSql & "         MAX(Decode(ItemCode, '420502', 'alb',"
    strSql = strSql & "                              '410101', 'KOH',"
    strSql = strSql & "                              '410104', 'ant', '' )) Other"
    strSql = strSql & "  FROM(  SELECT a.*, TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt1,c.DeptCode,"
    strSql = strSql & "                a.SLipno1 SLipno, a.SLipno2 Labno, RTRIM(a.ItemCd) ItemCode,"
    strSql = strSql & "                b.Sname, b.Sex sx, b.AgeYY age, d.Codenm SampleName,"
    strSql = strSql & "                NVL(LTRIM(RTRIM(a.Result1)), '..') RESULT11"
    strSql = strSql & "         FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "                TWEXAM_IDNOMST     b,"
    strSql = strSql & "                TWEXAM_General     c,"
    strSql = strSql & "                TWEXAM_Sample      d "
    'strSql = strSql & "         WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    'strSql = strSql & "         AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    
    strSql = strSql & "         WHERE  a.MDate    >= TO_DATE('" & sFrDate & " 00:01','yyyy-MM-dd hh24:mi')"
    strSql = strSql & "         AND    a.MDate    <= TO_DATE('" & sToDate & " 23:59','yyyy-MM-dd hh24:mi')"
    strSql = strSql & "         AND    a.ITemCD   IN ('420201','420202','420203','420204','420205',"
    strSql = strSql & "                               '440201','440202','440203','410103','420101',"
    strSql = strSql & "                               '440101','420501','410101','420503',"
    strSql = strSql & "                               '420703','420704','420801','420802','420803',"
    strSql = strSql & "                               '420804','420502','410101','410104')"
    strSql = strSql & "         AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & "         AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "         AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "         AND    a.SLipno2   = c.SLipno2(+)"
    If Trim(cmbSample.Text) <> "" Then
        strSql = strSql & "     AND    a.GeomchCd  = '" & Trim(Left(cmbSample.Text, 8)) & "'"
    End If
    strSql = strSql & "         AND    a.GeomchCD  = d.Code(+)"
    
    strSql = strSql & "         AND    c.GBCh      = 'Y'"
    strSql = strSql & "   )"
    strSql = strSql & "  GROUP BY JeobsuDt1, SLipno, Labno, Samplename, Ptno,  Sname, Sx, Age, DeptCode, MSeq"
    strSql = strSql & "  Order By MSeq"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 1
    Do Until adoSet.EOF
        sprMicro.Row = sprMicro.DataRowCnt + 1
        sprMicro.Col = 0:  sprMicro.Text = i
    
        sprMicro.Col = 1:  sprMicro.Text = adoSet.Fields("JeobsuDt1").Value & ""
        sprMicro.Col = 2:  sprMicro.Text = adoSet.Fields("Labno").Value & ""
        sprMicro.Col = 3:  sprMicro.Text = adoSet.Fields("Ptno").Value & ""
        sprMicro.Col = 4:  sprMicro.Text = Trim(adoSet.Fields("Sname").Value & "")
        sprMicro.Col = 5:  sprMicro.Text = adoSet.Fields("Sx").Value & ""
        sprMicro.Col = 6:  sprMicro.Text = adoSet.Fields("Age").Value & ""
        sprMicro.Col = 7:  sprMicro.Text = adoSet.Fields("DeptCode").Value & ""
        sprMicro.Col = 8:  sprMicro.Text = adoSet.Fields("Samplename").Value & ""
        sprMicro.Col = 9:  sprMicro.Text = adoSet.Fields("Mseq").Value & ""
        
        sprMicro.Col = 10:  sprMicro.Text = adoSet.Fields("RoutineCulture").Value & ""
        sprMicro.Col = 11: sprMicro.Text = adoSet.Fields("AFBCulture").Value & ""
        sprMicro.Col = 12: sprMicro.Text = adoSet.Fields("FungusCulture").Value & ""
        
        sprMicro.Col = 13: sprMicro.Text = adoSet.Fields("GramStain").Value & ""
        sprMicro.Col = 14: sprMicro.Text = adoSet.Fields("AFBStain").Value & ""
        sprMicro.Col = 15: sprMicro.Text = adoSet.Fields("FungusStain").Value & ""
        
        sprMicro.Col = 16: sprMicro.Text = adoSet.Fields("WetSmear").Value & ""
        sprMicro.Col = 17: sprMicro.Text = adoSet.Fields("IndiaInk").Value & ""
        sprMicro.Col = 18: sprMicro.Text = adoSet.Fields("Special").Value & ""
        sprMicro.Col = 19: sprMicro.Text = adoSet.Fields("Other").Value & ""
        sprMicro.Col = 20: sprMicro.Text = adoSet.Fields("SLipno").Value & ""
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
    
End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    sprMicro.Row = 3
    sprMicro.Row2 = sprMicro.DataRowCnt
    sprMicro.Col = 0
    sprMicro.Col2 = sprMicro.DataColCnt
    sprMicro.BlockMode = True
    sprMicro.Action = ActionClearText
    sprMicro.BlockMode = False
        
    GoSub Get_MicroSampleCode
    Exit Sub
    
    
Get_MicroSampleCode:
    Dim strSampleCode As String * 8
    
    cmbSample.Clear
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_SAMPLE"
    strSql = strSql & " WHERE  Class1 = 'm'"
    strSql = strSql & " ORDER  BY Seqno, Code"
        
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        strSampleCode = adoSet.Fields("Code").Value & ""
        cmbSample.AddItem strSampleCode & adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    cmbSample.AddItem " "
    Call adoSetClose(adoSet)
    
    Return
    
    

End Sub

Private Sub sprMicro_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sJDate          As String
    Dim sLabno          As String
    Dim sPtno           As String
    Dim sDeptCode       As String
    Dim iMseq           As String
    Dim sSLipno1        As String
    
    
    txtRemark.Text = ""
    
    'If Col = 0 Then
        sprMicro.Row = Row
        sprMicro.Col = 1: sJDate = sprMicro.Text
        sprMicro.Col = 2: sLabno = sprMicro.Text
        sprMicro.Col = 3: sPtno = sprMicro.Text
        sprMicro.Col = 7: sDeptCode = sprMicro.Text
        sprMicro.Col = 9: iMseq = sprMicro.Text
        sprMicro.Col = 20: sSLipno1 = sprMicro.Text
        GoSub Get_Remark_Text
    'End If
    
    Exit Sub
    
Get_Remark_Text:
    Dim sItemName       As String * 30
    
    strSql = ""
    strSql = strSql & " SELECT a.ItemCd ItemCode, a.CmDoctor, b.ItemNM ItemName,"
    strSql = strSql & "        a.ORderno"
    strSql = strSql & " FROM   TWEXAM_Order  a,"
    strSql = strSql & "        TWEXAM_ItemML b "
    strSql = strSql & " WHERE  a.CollDate = TO_DATE('" & sJDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.PTNO     = '" & sPtno & "'"
    strSql = strSql & " AND    a.SLipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.ItemCd   = b.Codeky"
    strSql = strSql & " UNION ALL "
    strSql = strSql & " SELECT Distinct a.ItemCd ItemCode, a.CmDoctor, b.RoutinNM ItemName,"
    strSql = strSql & "        a.ORderno"
    strSql = strSql & " FROM   TWEXAM_Order   a,"
    strSql = strSql & "        TWEXAM_Routine b "
    strSql = strSql & " WHERE  a.CollDate = TO_DATE('" & sJDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.PTNO     = '" & sPtno & "'"
    strSql = strSql & " AND    a.SLipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.ItemCd   = b.RoutinCd"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    
    Do Until adoSet.EOF
        sItemName = adoSet.Fields("ItemName").Value & ""
        
        txtRemark.Text = txtRemark.Text & sItemName & " : " & _
                         Trim(adoSet.Fields("CmDoctor").Value & "") & "  :  " & _
                         "(" & adoSet.Fields("Orderno").Value & ")" & vbCrLf
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub sprMicro_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim nMicroSeq       As Long
    Dim sJeobsuDt       As String
    Dim GnSLipno1       As Integer
    Dim GnSLipno2       As Long
    
    
    If Row < 3 Then Exit Sub
    
    sprMicro.Row = Row
    sprMicro.Col = 1: sJeobsuDt = sprMicro.Text
    sprMicro.Col = 9: nMicroSeq = Val(sprMicro.Text)
    
    frmResult.Show
    frmResult.ZOrder 0
    
    GoSub GET_SLipno
    
    frmResult.dtJeobsu.Value = sJeobsuDt
    Call SetComboBox(frmResult.cmbSLip, GnSLipno1, 2)
    frmResult.txtSLipno2.Text = GnSLipno2
    Call frmResult.txtSLipno2_KeyDown(vbKeyReturn, 1)
    
    Exit Sub
    
    
GET_SLipno:
    GnSLipno1 = 0
    GnSLipno2 = 0
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_General_sub"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    MSEQ     = " & nMicroSeq
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    GnSLipno1 = Val(adoSet.Fields("SLipno1").Value & "")
    GnSLipno2 = Val(adoSet.Fields("SLipno2").Value & "")
    
    Call adoSetClose(adoSet)
    Return
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    
    Select Case Button.Index
        Case 1: Unload Me
        'Case 3: frmMicroEnrol.Show
        '        frmMicroEnrol.ZOrder 0
        
    End Select
    
End Sub
