VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmItemResult 
   Caption         =   "Item 별 결과관리 화면"
   ClientHeight    =   7725
   ClientLeft      =   1125
   ClientTop       =   3120
   ClientWidth     =   11820
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
   ScaleHeight     =   7725
   ScaleWidth      =   11820
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5985
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemResult.frx":0000
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemResult.frx":0324
            Key             =   "ABO"
            Object.Tag             =   "ABO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   635
      ButtonWidth     =   1376
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
            Object.ToolTipText     =   "Exit of ItemResult"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ABO"
            Key             =   "ABO"
            Object.ToolTipText     =   "ABO & Rh Type 입력"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox CboPart 
      Height          =   300
      Left            =   1125
      Style           =   2  '드롭다운 목록
      TabIndex        =   7
      Top             =   1260
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Caption         =   "결과확인"
      Height          =   240
      Left            =   945
      TabIndex        =   6
      Top             =   1620
      Width           =   1050
   End
   Begin VB.OptionButton Option2 
      Caption         =   "결과미확인"
      Height          =   240
      Left            =   2070
      TabIndex        =   5
      Top             =   1620
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.TextBox txtDeltaMin 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   10305
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   765
      Width           =   600
   End
   Begin VB.TextBox txtDeltaMax 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   10890
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   765
      Width           =   600
   End
   Begin VB.TextBox txtDeltaQc 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   765
      Width           =   240
   End
   Begin VB.TextBox txtPanicMax 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   10755
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox txtPanicMin 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   450
      Width           =   690
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   780
      Left            =   225
      TabIndex        =   8
      Top             =   450
      Width           =   3300
      _Version        =   65536
      _ExtentX        =   5821
      _ExtentY        =   1376
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
      Alignment       =   0
      Begin MSComCtl2.DTPicker dtTdate 
         Height          =   315
         Left            =   1935
         TabIndex        =   9
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36306
      End
      Begin MSComCtl2.DTPicker dtFdate 
         Height          =   315
         Left            =   630
         TabIndex        =   10
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36306
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자:From/To"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   135
         Width           =   1455
      End
   End
   Begin FPSpreadADO.fpSpread SprCodeNm 
      Height          =   5415
      Left            =   135
      TabIndex        =   12
      Top             =   1935
      Width           =   3360
      _Version        =   196608
      _ExtentX        =   5927
      _ExtentY        =   9551
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
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   8
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmItemResult.frx":0C00
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   100
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread ssResult 
      Height          =   5415
      Left            =   3690
      TabIndex        =   13
      Top             =   1935
      Width           =   7935
      _Version        =   196608
      _ExtentX        =   13996
      _ExtentY        =   9551
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
      MaxCols         =   14
      ScrollBars      =   2
      SpreadDesigner  =   "frmItemResult.frx":1B9C
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread ssRefdata 
      Height          =   825
      Left            =   3735
      TabIndex        =   14
      Top             =   450
      Width           =   5685
      _Version        =   196608
      _ExtentX        =   10028
      _ExtentY        =   1455
      _StockProps     =   64
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
      MaxCols         =   7
      MaxRows         =   10
      OperationMode   =   1
      ScrollBars      =   0
      SpreadDesigner  =   "frmItemResult.frx":59F5
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdRetInsert 
      Height          =   465
      Left            =   9765
      TabIndex        =   21
      Top             =   1440
      Width           =   1455
      Caption         =   "결과입력"
      PicturePosition =   327683
      Size            =   "2566;820"
      Picture         =   "frmItemResult.frx":5F40
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdChoise 
      Height          =   465
      Left            =   3690
      TabIndex        =   20
      Top             =   1485
      Width           =   1230
      Caption         =   "전체선택"
      PicturePosition =   327683
      Size            =   "2170;820"
      Picture         =   "frmItemResult.frx":7702
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   465
      Left            =   8280
      TabIndex        =   19
      Top             =   1440
      Width           =   1500
      Caption         =   "Sheet출력"
      PicturePosition =   327683
      Size            =   "2646;820"
      Picture         =   "frmItemResult.frx":7FDC
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      Caption         =   "검사종류:"
      Height          =   195
      Left            =   225
      TabIndex        =   17
      Top             =   1305
      Width           =   825
   End
   Begin VB.Label Label4 
      Caption         =   "Delta"
      Height          =   195
      Left            =   9495
      TabIndex        =   16
      Top             =   810
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "Panic"
      Height          =   195
      Left            =   9495
      TabIndex        =   15
      Top             =   495
      Width           =   510
   End
End
Attribute VB_Name = "frmItemResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FsExamGubun         As String
Dim FiRow               As Integer
Dim FsitemCD            As String


Private Sub CboPart_Click()
    Dim ii              As Integer


    Call SpreadSetClear(SprCodeNm)
    
    FsExamGubun = Left$(CboPart.Text, 2)
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT CodeKy,itemNm, PanicMin, PanicMax, DeltaQc, DeltaMin, DeltaMax"
    gStrSql = gStrSql & " FROM   TWEXAM_ITEMML "
    gStrSql = gStrSql & " WHERE  CodeKy   Like  '" & FsExamGubun & "%'   "
    'gStrSql = gStrSql & " AND    GbRoutine   =  'I'  "
    gStrSql = gStrSql & " ORDER  BY  CodeKy "
    
    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        SprCodeNm.Row = SprCodeNm.DataRowCnt + 1
        SprCodeNm.Col = 2: SprCodeNm.Text = adoSet.Fields("itemNm").Value & ""
        SprCodeNm.Col = 3: SprCodeNm.Text = adoSet.Fields("CodeKy").Value & ""
        SprCodeNm.Col = 4: SprCodeNm.Text = adoSet.Fields("PanicMin").Value & ""
        SprCodeNm.Col = 5: SprCodeNm.Text = adoSet.Fields("PanicMax").Value & ""
        
        SprCodeNm.Col = 6: SprCodeNm.Text = adoSet.Fields("DeltaQC").Value & ""
        SprCodeNm.Col = 7: SprCodeNm.Text = adoSet.Fields("DeltaMin").Value & ""
        SprCodeNm.Col = 8: SprCodeNm.Text = adoSet.Fields("DeltaMax").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    

End Sub

Private Sub cmdChoise_Click()
    
    If cmdChoise.Caption = "전체선택" Then
        For i = 1 To ssResult.DataRowCnt
            ssResult.Row = i
            ssResult.Col = 1
            ssResult.Value = True
        Next
        cmdChoise.Caption = "전체해제"
    Else
        For i = 1 To ssResult.DataRowCnt
            ssResult.Row = i
            ssResult.Col = 1
            ssResult.Value = False
        Next
        cmdChoise.Caption = "전체선택"
    End If
    
End Sub


Private Sub cmdPr_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    SprCodeNm.Row = SprCodeNm.ActiveRow
    SprCodeNm.Col = 2: sItemName = SprCodeNm.Text
    
    
    If ssResult.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "Item별 접수내역 LIST" & " - " & sItemName
    
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFdate.Value, "yyyy-MM-dd hh:mm ampm") & " / " & _
                                               Format(dtTdate.Value, "yyyy-MM-dd hh:mm ampm")
    
    ssResult.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                            strFont(1) + "/n" + sBarLine + strFont(1)
    ssResult.PrintFooter = "/f2" & "/l" & sBarLine & _
                            "/n" & Space(80) & "Page : " & "/p" & " of " & ssResult.PrintPageCount
    ssResult.PrintMarginLeft = 0
    ssResult.PrintMarginRight = 0
    ssResult.PrintMarginTop = 0
    ssResult.PrintMarginBottom = 0
    ssResult.PrintColHeaders = True
    ssResult.PrintRowHeaders = True
    ssResult.PrintBorder = False
    ssResult.PrintColor = False
    ssResult.PrintGrid = True
    ssResult.PrintShadows = True
    ssResult.PrintUseDataMax = False
    ssResult.Row = 1
    ssResult.Row2 = ssResult.DataRowCnt
    ssResult.Col = 2
    ssResult.Col2 = ssResult.MaxCols
    ssResult.PrintOrientation = 1
    ssResult.PrintOrientation = PrintOrientationPortrait
    ssResult.PrintType = PrintTypeCellRange
    ssResult.Action = ActionPrint

End Sub

Private Sub cmdRetInsert_Click()
    Dim sJeobsuDt       As String
    Dim sSLno1          As String
    Dim sSLno2          As String
    Dim sItemCd         As String
    Dim sResult1        As String
    Dim sRowID          As String
    
    
    If ssResult.DataRowCnt = 0 Then Exit Sub
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        If ssResult.Value = True Then
            ssResult.Col = 2:  sRowID = ssResult.Text
            ssResult.Col = 3:  sJeobsuDt = ssResult.Text
            ssResult.Col = 11: sResult1 = ssResult.Text
            ssResult.Col = 4:  sSLno1 = ssResult.Text
            ssResult.Col = 5:  sSLno2 = ssResult.Text
            
            If Trim(sSLno1) = "" Then Exit Sub
            
            strSql = ""
            strSql = strSql & " UPDATE TWEXAM_General_Sub"
            strSql = strSql & " SET   Result1  = '" & Quot_Conv(sResult1) & "',"
            strSql = strSql & "       Verify   = 'Y'"
            strSql = strSql & " WHERE ROWID    = '" & sRowID & "'"
            adoConnect.BeginTrans
            If adoExec(strSql) Then
                adoConnect.CommitTrans
                GoSub Set_Status_General
            Else
                adoConnect.RollbackTrans
            End If
            
            
        End If
    Next
    
    MsgBox "결과를 입력하였습니다!.........", vbInformation
    Exit Sub
    
Set_Status_General:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    Status   = 'P'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & Val(sSLno1)
    strSql = strSql & " AND    SLipno2  = " & Val(sSLno2)
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
        
End Sub

Private Sub Form_Load()
    Dim ii              As Integer
    Dim LsExamGubun     As String


    gStrSql = ""
    gStrSql = gStrSql & " SELECT CODEKY, CODENM  FROM TWEXAM_SPECODE "
    gStrSql = gStrSql & " WHERE  CODEGU = '12'    "
    gStrSql = gStrSql & " AND    CODEKY < '90'    "
    gStrSql = gStrSql & " ORDER  BY CODEKY    ASC "
    
    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        LsExamGubun = Trim(adoSet.Fields("CODEKY").Value & "")
        LsExamGubun = LsExamGubun & " " & Trim(adoSet.Fields("CODENM").Value & "")
        CboPart.AddItem LsExamGubun
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    dtFdate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtTdate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    FiRow = 1
    
    SprCodeNm.Col = 3
    SprCodeNm.ColHidden = True

End Sub


Private Sub SprCodeNm_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim sItemCd     As String
    Dim sFromDate   As String
    Dim sToDate     As String
    
    
    ssResult.MaxRows = 0
    ssResult.MaxRows = 200
    ssResult.RowHeight(-1) = 9.55
    
    If Row > SprCodeNm.DataRowCnt Then Exit Sub
    
    If Col = 1 Then
        For i = 1 To SprCodeNm.DataRowCnt
            SprCodeNm.Row = i
            SprCodeNm.Col = 1
            SprCodeNm.TypeButtonText = ""
        Next
        SprCodeNm.Row = Row
        SprCodeNm.Col = 1
        SprCodeNm.TypeButtonText = "☞"
    End If
    
        
    sFromDate = Format(dtFdate.Value, "yyyy-MM-dd")
    sToDate = Format(dtTdate.Value, "yyyy-MM-dd")
    
    SprCodeNm.Row = Row
    SprCodeNm.Col = 3:  sItemCd = Trim(SprCodeNm.Text)
    If Row = 0 Then Exit Sub
    If Col > 1 Then Exit Sub
    
    If Trim(sItemCd) = "" Then Exit Sub
        
    txtPanicMin.Text = "": txtPanicMax.Text = ""
    SprCodeNm.Row = Row
    SprCodeNm.Col = 4: txtPanicMin.Text = SprCodeNm.Text
    SprCodeNm.Col = 5: txtPanicMax.Text = SprCodeNm.Text
    
    
    SprCodeNm.Col = 6: txtDeltaQc.Text = SprCodeNm.Text
    SprCodeNm.Col = 7: txtDeltaMin.Text = SprCodeNm.Text
    SprCodeNm.Col = 8: txtDeltaMax.Text = SprCodeNm.Text
    
    
    GoSub Get_RefData
    GoSub Get_GeneralSub
    GoSub Default_Result_Set
    
    Exit Sub
    

    
Get_RefData:
    Call SpreadSetClear(ssRefdata)
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT a.*, TO_CHAR(a.appDate, 'YYYY-MM-DD') APPdate"
    gStrSql = gStrSql & " FROM   TWEXAM_RefData a"
    gStrSql = gStrSql & " WHERE  a.iTemcode  =  '" & sItemCd & "'"
    gStrSql = gStrSql & " ORDER  BY a.appDate, a.AgeMin"
    If False = adoSetOpen(gStrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        ssRefdata.Row = ssRefdata.DataRowCnt + 1
        ssRefdata.Col = 1: ssRefdata.Text = adoSet.Fields("appDate").Value & ""
        ssRefdata.Col = 2: ssRefdata.Text = Trim(adoSet.Fields("AgeMin").Value & "")
        ssRefdata.Col = 3: ssRefdata.Text = Trim(adoSet.Fields("AgeMax").Value & "")
        ssRefdata.Col = 4: ssRefdata.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefdata.Col = 5: ssRefdata.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefdata.Col = 6: ssRefdata.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefdata.Col = 7: ssRefdata.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

Get_GeneralSub:
    Call SpreadSetClear(ssResult)
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT a.RowiD RWID, a.*, TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    gStrSql = gStrSql & "        b.Sname, b.Sex, b.AgeYY"
    gStrSql = gStrSql & " FROM   TWEXAM_GENERAL_SUB a,"
    gStrSql = gStrSql & "        TWEXAM_IDnomst     b,"
    gStrSql = gStrSql & "        TWEXAM_General     c"
    gStrSql = gStrSql & " WHERE  a.Jeobsudt >= TO_DATE('" & sFromDate & "','yyyy-MM-dd')"
    gStrSql = gStrSql & " AND    a.Jeobsudt <= TO_DATE('" & sToDate & "',  'yyyy-MM-dd')"
    gStrSql = gStrSql & " AND    a.itemCD    =  '" & sItemCd & "'"
    gStrSql = gStrSql & " AND    a.JeobsuDt  = c.JeobsuDt(+)"
    gStrSql = gStrSql & " AND    a.SLipno1   = c.SLipno1(+)"
    gStrSql = gStrSql & " AND    a.SLipno2   = c.SLipno2(+)"
    
    If Option1.Value = True Then
        gStrSql = gStrSql & " AND a.Verify = 'Y'"
    Else
        gStrSql = gStrSql & " AND a.Verify <> 'Y'"
    End If
    gStrSql = gStrSql & " AND    c.GbCh      = 'Y'"
    gStrSql = gStrSql & " AND    a.Ptno      = b.Ptno(+)"
    
    
    If False = adoSetOpen(gStrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("RWID").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("Slipno1").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("Slipno2").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("ItemCD").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 10:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 11: ssResult.Text = Trim$(adoSet.Fields("Result1").Value & "")
        ssResult.Col = 13: GoSub Prev_Data_Select
        Call ssResult_LeaveCell(11, ssResult.Row, 11, ssResult.Row, True)
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

Prev_Data_Select:
    Dim sPrePtno        As String
    Dim sPreJdate       As String
    Dim sPreItemCd      As String
    Dim sPreLabno       As String
    
    Dim adoPre          As ADODB.Recordset
    
    ssResult.Row = ssResult.Row
    ssResult.Col = 3: sPreJdate = ssResult.Text
    ssResult.Col = 5: sPreLabno = ssResult.Text
    ssResult.Col = 6: sPreItemCd = ssResult.Text
    ssResult.Col = 7: sPrePtno = ssResult.Text
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt,"
    gStrSql = gStrSql & "        a.Result1, a.Slipno2, a.ItemCD"
    gStrSql = gStrSql & " FROM   TWEXAM_GENERAL_SUB a"
    gStrSql = gStrSql & " WHERE  a.Ptno      =  '" & sPrePtno & "'"
    gStrSql = gStrSql & " AND    a.ItemCD    =  '" & sPreItemCd & "'"
    gStrSql = gStrSql & " AND    a.Verify    =  'Y'"
    gStrSql = gStrSql & " AND    a.JeobsuDt || a.SLipno2  = ( SELECT Max(b.JeobsuDt || b.SLipno2)"
    gStrSql = gStrSql & "                        FROM   TWEXAM_GENERAL_SUB b"
    gStrSql = gStrSql & "                        WHERE  TO_CHAR(b.JeobsuDt,'YYYY-MM-DD') || Ltrim(b.SLipno2) < "
    gStrSql = gStrSql & "                               '" & sPreJdate & "' || '" & sPreLabno & "'"
    gStrSql = gStrSql & "                        AND    b.Ptno     =  '" & sPrePtno & "'"
    gStrSql = gStrSql & "                        AND    b.ItemCD   = '" & sPreItemCd & "'"
    gStrSql = gStrSql & "                        AND    b.Verify   = 'Y')"
    gStrSql = gStrSql & " ORDER  BY a.JeobsuDt DESC, a.SLipno2 DESC"
    
    If False = adoSetOpen(gStrSql, adoPre) Then Return
    
    ssResult.Col = 13
    ssResult.Text = adoPre.Fields("Result1").Value & ""
    Call adoSetClose(adoPre)
    
    Return
        
        
Default_Result_Set:
    Dim sResult     As String
    Dim sOldRet     As String
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 6
        sResult = Get_Result_Text(Trim(ssResult.Text))
        If Trim(sResult) <> "" Then
            ssResult.Col = 11
            ssResult.CellType = CellTypeComboBox
            ssResult.TypeComboBoxList = sResult
            ssResult.TypeComboBoxEditable = True
            
        End If
    Next
    Return


End Sub

Private Sub ssResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim sngResult   As Single
    
    GoSub Panic_Data_Check
    GoSub Delta_Data_Check
    Exit Sub
    


Panic_Data_Check:
    
    If Trim(txtPanicMin.Text) = "" And _
       Trim(txtPanicMax.Text) = "" Then Return
    
        
    ssResult.Row = Row
    ssResult.Col = 11
    
    If Trim(ssResult.Text) = "" Then Return
    
    If False = IsNumeric(ssResult.Text) Then Return
    
    sngResult = CSng(ssResult.Text)
    
    If sngResult < CSng(txtPanicMin.Text) Or sngResult > CSng(txtPanicMax.Text) Then
        ssResult.Col = 12: ssResult.BackColor = RGB(255, 255, 230)
                           ssResult.Text = "p"
    Else
        ssResult.Col = 12: ssResult.BackColor = RGB(255, 255, 255)
                           ssResult.Text = ""
    End If
    
    Return
    
    
Delta_Data_Check:
    Dim LiPreVal
    Dim LiCurVal
    Dim LiDeltaMin
    Dim LiDeltaMax
    Dim LsQC        As String
    


    ssResult.Col = 13
    If ssResult.Text = "" Then
        ssResult.Text = ""
        ssResult.BackColor = RGB(255, 255, 255)
        Return
    End If
    
    If False = IsNumeric(ssResult.Text) Then Return
    
    ssResult.Row = Row
    ssResult.Col = 11:   LiCurVal = Val(ssResult.Text)
    ssResult.Col = 13:   LiPreVal = Val(ssResult.Text)
                         LsQC = txtDeltaQc.Text
                         LiDeltaMin = Val(txtDeltaMin.Text)
                         LiDeltaMax = Val(txtDeltaMax.Text)
    
    
    If LiPreVal <> 0 And LiCurVal <> 0 Then       '양쪽(계산할 2개의 Data가 모두 있을때...)
        If LsQC = "1" Or LsQC = "2" Or LsQC = "3" Or LsQC = "4" Then
            If DeltaCheck(LiCurVal, LiPreVal, LsQC) < LiDeltaMin Or _
               DeltaCheck(LiCurVal, LiPreVal, LsQC) > LiDeltaMax Then
                ssResult.Col = 14
                If Trim(ssResult.Text) = "" Then
                    ssResult.Text = "d"
                    ssResult.BackColor = RGB(250, 0, 0)
                End If
            Else
                ssResult.Col = 14
                ssResult.Text = ""
                ssResult.BackColor = RGB(255, 255, 255)
            End If
        End If
    End If
    
    Return
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
        Case 3: frmABOType.Show vbModal
    End Select
    
End Sub
