VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSheetStool 
   Caption         =   "기생충검사 결과관리"
   ClientHeight    =   7830
   ClientLeft      =   225
   ClientTop       =   930
   ClientWidth     =   11670
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
   ScaleHeight     =   7830
   ScaleWidth      =   11670
   WindowState     =   2  '최대화
   Begin VB.TextBox txtRowh 
      Height          =   285
      Left            =   7425
      TabIndex        =   11
      Text            =   "11"
      Top             =   675
      Width           =   645
   End
   Begin FPSpreadADO.fpSpread sprStool 
      Height          =   6000
      Left            =   135
      TabIndex        =   2
      Top             =   1125
      Width           =   11535
      _Version        =   196608
      _ExtentX        =   20346
      _ExtentY        =   10583
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
      MaxCols         =   12
      ScrollBars      =   2
      SpreadDesigner  =   "frmSheetStool.frx":0000
      UserResize      =   1
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   3375
      TabIndex        =   1
      Top             =   630
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36566
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11295
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
            Picture         =   "frmSheetStool.frx":1BFE
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
      Width           =   11670
      _ExtentX        =   20585
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
            Description     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.OptionButton Option3 
         Caption         =   "ALL"
         Height          =   225
         Left            =   5490
         TabIndex        =   10
         Top             =   45
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.OptionButton Option2 
         Caption         =   "NoVerify"
         Height          =   225
         Left            =   4365
         TabIndex        =   9
         Top             =   45
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Verify"
         Height          =   225
         Left            =   3420
         TabIndex        =   8
         Top             =   45
         Width           =   870
      End
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   4860
      TabIndex        =   7
      Top             =   630
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24444931
      CurrentDate     =   36566
   End
   Begin VB.Label Label2 
      Caption         =   "RowHeight"
      Height          =   195
      Left            =   6525
      TabIndex        =   12
      Top             =   720
      Width           =   870
   End
   Begin MSForms.CommandButton cmdChoise 
      Height          =   465
      Left            =   405
      TabIndex        =   6
      Top             =   630
      Width           =   1815
      Caption         =   "전체선택"
      PicturePosition =   327683
      Size            =   "3201;820"
      Picture         =   "frmSheetStool.frx":1F1A
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   465
      Left            =   9585
      TabIndex        =   5
      Top             =   630
      Width           =   1500
      Caption         =   "출력"
      PicturePosition =   327683
      Size            =   "2646;820"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "접수일자:"
      Height          =   195
      Left            =   2430
      TabIndex        =   4
      Top             =   675
      Width           =   825
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   465
      Left            =   8190
      TabIndex        =   3
      Top             =   630
      Width           =   1410
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2487;820"
      Picture         =   "frmSheetStool.frx":27F4
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmSheetStool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChoise_Click()
    If cmdChoise.Caption = "전체선택" Then
        For i = 1 To sprStool.DataRowCnt
            sprStool.Row = i
            sprStool.Col = 1
            sprStool.Value = True
        Next
        cmdChoise.Caption = "전체해제"
    Else
        For i = 1 To sprStool.DataRowCnt
            sprStool.Row = i
            sprStool.Col = 1
            sprStool.Value = False
        Next
        cmdChoise.Caption = "전체선택"
    End If


End Sub

Private Sub cmdPrint_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    If sprStool.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "기생충검사  접수내역 LIST"
    
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFrDate.Value, "yyyy-MM-dd") & " / " & _
                                               Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprStool.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                            strFont(1) + "/n" + sBarLine + strFont(1)
    sprStool.PrintFooter = "/f2" & "/l" & sBarLine & _
                            "/n" & Space(80) & "Page : " & "/p" & " of " & sprStool.PrintPageCount
    sprStool.PrintMarginLeft = 0
    sprStool.PrintMarginRight = 0
    sprStool.PrintMarginTop = 0
    sprStool.PrintMarginBottom = 0
    sprStool.PrintColHeaders = True
    sprStool.PrintRowHeaders = True
    sprStool.PrintBorder = False
    sprStool.PrintColor = False
    sprStool.PrintGrid = True
    sprStool.PrintShadows = True
    sprStool.PrintUseDataMax = False
    sprStool.Row = 1
    sprStool.Row2 = sprStool.DataRowCnt
    sprStool.Col = 2
    sprStool.Col2 = sprStool.MaxCols
    sprStool.PrintOrientation = 1
    sprStool.PrintOrientation = PrintOrientationPortrait
    sprStool.PrintType = PrintTypeCellRange
    sprStool.Action = ActionPrint

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate             As String
    Dim sToDate             As String
    

    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprStool.ReDraw = False
    sprStool.MaxRows = 0
    sprStool.MaxRows = 300
    If Trim(Me.txtRowh.Text) = "" Then
        sprStool.RowHeight(-1) = 11
    Else
        sprStool.RowHeight(-1) = Val(Me.txtRowh.Text)
    End If
    sprStool.ReDraw = True
    
    
    
    For i = 1 To sprStool.DataRowCnt
        sprStool.Row = i
        sprStool.Col = 1
        sprStool.Value = False
    Next
    cmdChoise.Caption = "전체선택"
    
    
    GoSub Get_Stool_Data
    
    Exit Sub
    
    

Get_Stool_Data:
    strSql = ""
    strSql = strSql & " SELECT JeobsuDt1, Labno, Ptno,  Sname, Sx, Age, DeptCode, "
    strSql = strSql & "        MAX(Decode(ItemCode, '430101', Result11, '--')) S1,"
    strSql = strSql & "        MAX(Decode(ItemCode, '430102', Result11, '--')) S2,"
    strSql = strSql & "        MAX(Decode(ItemCode, '430103', Result11, '--')) S3,"
    strSql = strSql & "        MAX(Decode(ItemCode, '430104', Result11, '--')) S4 "
    strSql = strSql & " FROM(  SELECT a.*, TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt1,c.DeptCode,"
    strSql = strSql & "               a.SLipno2 Labno, RTRIM(a.ItemCd) ItemCode,b.Sname, b.Sex sx, b.AgeYY age,"
    strSql = strSql & "               NVL(LTRIM(RTRIM(a.Result1)), '..') RESULT11"
    strSql = strSql & "        FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "               TWEXAM_IDNOMST     b,"
    strSql = strSql & "               TWEXAM_General     c "
    strSql = strSql & "        WHERE  a.MDate >= TO_DATE('" & sFrDate & " 00:01','yyyy-MM-dd hh24:mi')"
    strSql = strSql & "        AND    a.MDate <= TO_DATE('" & sToDate & " 23:59','yyyy-MM-dd hh24:mi')"
    strSql = strSql & "        AND    a.ITEMCD   IN ('430101','430102','430103','430104')"
    
    If Option1.Value = True Then
        strSql = strSql & "    AND    a.Verify    = 'Y'": End If
    If Option2.Value = True Then
        strSql = strSql & "    AND    a.Verify    = 'N'": End If
        
    strSql = strSql & "        AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & "        AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "        AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "        AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "        AND    c.GBCH      = 'Y')"
    strSql = strSql & " GROUP BY JeobsuDt1, Labno, Ptno,  Sname, Sx, Age, DeptCode"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprStool.Row = sprStool.DataRowCnt + 1
        sprStool.Col = 1:  sprStool.Value = False
        sprStool.Col = 2:  sprStool.Text = adoSet.Fields("JeobsuDt1").Value & ""
        sprStool.Col = 3:  sprStool.Text = adoSet.Fields("Labno").Value & ""
        sprStool.Col = 4:  sprStool.Text = adoSet.Fields("Ptno").Value & ""
        sprStool.Col = 5:  sprStool.Text = Trim(adoSet.Fields("Sname").Value & "")
        sprStool.Col = 6:  sprStool.Text = adoSet.Fields("Sx").Value & ""
        sprStool.Col = 7:  sprStool.Text = adoSet.Fields("Age").Value & ""
        sprStool.Col = 8:  sprStool.Text = adoSet.Fields("DeptCode").Value & ""
        
        For i = 1 To 4
            sprStool.Col = 8 + i
            Select Case Trim(adoSet.Fields("S" & i).Value & "")
                Case "..":    sprStool.Text = ""
                Case "--":  sprStool.Text = "---------------"
                              sprStool.Lock = True
                Case Else:    sprStool.Text = Trim(adoSet.Fields("S" & i).Value & "")
            End Select
        Next
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
    
    
End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    Call SpreadSetClear(Me.sprStool)
    

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
