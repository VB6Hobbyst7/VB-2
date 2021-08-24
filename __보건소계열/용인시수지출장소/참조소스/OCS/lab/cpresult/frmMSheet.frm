VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMEnrol 
   Caption         =   "Item 별 결과관리 화면"
   ClientHeight    =   7725
   ClientLeft      =   90
   ClientTop       =   1230
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
   Begin FPSpreadADO.fpSpread sprSheet 
      Height          =   465
      Left            =   7065
      TabIndex        =   17
      Top             =   405
      Visible         =   0   'False
      Width           =   1410
      _Version        =   196608
      _ExtentX        =   2487
      _ExtentY        =   820
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
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
      MaxRows         =   14
      ScrollBars      =   0
      SpreadDesigner  =   "frmMSheet.frx":0000
      UserResize      =   0
      Appearance      =   1
   End
   Begin VB.TextBox txtRowh 
      Height          =   285
      Left            =   6975
      TabIndex        =   15
      Text            =   "11"
      Top             =   1260
      Width           =   645
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   420
      Left            =   3240
      TabIndex        =   13
      Top             =   1215
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2725
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   " >"
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
      Alignment       =   1
      Begin VB.TextBox txtMseq 
         Appearance      =   0  '평면
         Height          =   330
         Left            =   360
         TabIndex        =   14
         Top             =   45
         Width           =   1005
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11340
      Top             =   495
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
            Picture         =   "frmMSheet.frx":13BD
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMSheet.frx":16E1
            Key             =   "ABO"
            Object.Tag             =   "ABO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
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
            Object.ToolTipText     =   "Exit of ItemResult"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "결과확인"
      Height          =   240
      Left            =   225
      TabIndex        =   1
      Top             =   1080
      Width           =   1230
   End
   Begin VB.OptionButton Option2 
      Caption         =   "결과미확인"
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   1350
      Value           =   -1  'True
      Width           =   1230
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   225
      TabIndex        =   2
      Top             =   450
      Width           =   6675
      _Version        =   65536
      _ExtentX        =   11774
      _ExtentY        =   979
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
      Begin VB.ComboBox CboPart 
         Height          =   300
         Left            =   3690
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   90
         Width           =   2640
      End
      Begin MSComCtl2.DTPicker dtTdate 
         Height          =   315
         Left            =   2295
         TabIndex        =   3
         Top             =   90
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
         Format          =   24772611
         CurrentDate     =   36306
      End
      Begin MSComCtl2.DTPicker dtFdate 
         Height          =   315
         Left            =   990
         TabIndex        =   4
         Top             =   90
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
         Format          =   24772611
         CurrentDate     =   36306
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자:"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   135
         Width           =   825
      End
   End
   Begin FPSpreadADO.fpSpread SprCodeNm 
      Height          =   5415
      Left            =   45
      TabIndex        =   6
      Top             =   1710
      Width           =   3090
      _Version        =   196608
      _ExtentX        =   5450
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
      MaxCols         =   9
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmMSheet.frx":1FBD
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   100
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread ssResult 
      Height          =   5415
      Left            =   3150
      TabIndex        =   7
      Top             =   1710
      Width           =   8610
      _Version        =   196608
      _ExtentX        =   15187
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
      MaxCols         =   18
      ScrollBars      =   2
      SpreadDesigner  =   "frmMSheet.frx":2F77
      Appearance      =   1
      TextTip         =   1
   End
   Begin MSForms.CommandButton cmdPr2 
      Height          =   465
      Left            =   10395
      TabIndex        =   18
      Top             =   1125
      Width           =   1230
      Caption         =   "장부"
      PicturePosition =   327683
      Size            =   "2170;820"
      Picture         =   "frmMSheet.frx":6EAF
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      Caption         =   "RowHeight"
      Height          =   195
      Left            =   6075
      TabIndex        =   16
      Top             =   1305
      Width           =   870
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   600
      Left            =   1530
      TabIndex        =   11
      Top             =   1035
      Width           =   1635
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2884;1058"
      Picture         =   "frmMSheet.frx":7791
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdChoise 
      Height          =   465
      Left            =   7740
      TabIndex        =   10
      Top             =   1125
      Width           =   1230
      Caption         =   "전체선택"
      PicturePosition =   327683
      Size            =   "2170;820"
      Picture         =   "frmMSheet.frx":8073
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPr 
      Height          =   465
      Left            =   9000
      TabIndex        =   9
      Top             =   1125
      Width           =   1410
      Caption         =   "Sheet출력"
      PicturePosition =   327683
      Size            =   "2487;820"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmMEnrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FsExamGubun         As String
Dim FiRow               As Integer
Dim FsitemCD            As String
Dim gStrSql             As String

Public Function isSensItem(ByVal CheckItem As String) As Integer
    Dim adoSi       As ADODB.Recordset
    
    isSensItem = False
    
    strSql = ""
    strSql = strSql & " SELECT GeomsaAb"
    strSql = strSql & " FROM   TWEXAM_ItemML"
    strSql = strSql & " WHERE  Codeky = '" & CheckItem & "'"
    
    If False = adoSetOpen(strSql, adoSi) Then Exit Function
    
    If Trim(adoSi.Fields("GeomsaAb").Value & "") = "S" Then
        isSensItem = True
    End If
    
    Call adoSetClose(adoSi)

End Function



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
    
    
    For i = 1 To 80
        sBarLine = sBarLine & "━"
    Next
    
    
    SprCodeNm.Row = SprCodeNm.ActiveRow
    SprCodeNm.Col = 2: sItemName = SprCodeNm.Text
    
    
    If ssResult.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    
    ssResult.ColWidth(13) = 25
    
    strFont(0) = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "Item별 접수내역 LIST" & " - " & sItemName
    
    strHead(1) = "/f2" & "/l" & "접수일자(Fr/To): " & Format(dtFdate.Value, "yyyy-MM-dd hh:mm ampm") & " / " & _
                                                      Format(dtTdate.Value, "yyyy-MM-dd hh:mm ampm")
    
    ssResult.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                           strFont(1) + "/n" + strFont(1) + "/l" + sBarLine + strFont(1)
    ssResult.PrintFooter = "/f2" & "/l" & strFont(1) + sBarLine & _
                           "/n" & Space(80) & "Page : " & "/p" & " of " & ssResult.PrintPageCount
    ssResult.PrintMarginLeft = 0
    ssResult.PrintMarginRight = 0
    ssResult.PrintMarginTop = 0
    ssResult.PrintMarginBottom = 0
    ssResult.PrintColHeaders = True
    ssResult.PrintRowHeaders = True
    ssResult.PrintBorder = True
    ssResult.PrintColor = True
    ssResult.PrintGrid = False
    ssResult.PrintShadows = True
    ssResult.PrintUseDataMax = False
    ssResult.Row = 1
    ssResult.Row2 = ssResult.DataRowCnt
    ssResult.Col = 2
    ssResult.Col2 = ssResult.MaxCols
    ssResult.PrintOrientation = 1
    ssResult.PrintOrientation = PrintOrientationLandscape
    ssResult.PrintType = PrintTypeCellRange
    ssResult.Action = ActionPrint
    
    ssResult.ColWidth(12) = 10.13

End Sub

Private Sub cmdPr2_Click()
    Dim sJeobsuDt       As String
    Dim sSLipno1        As String * 2
    Dim sSLipno2        As String * 5
    Dim sMicrono        As String * 8
    Dim sLabno          As String * 15
    Dim sPtno           As String
    Dim sName           As String
    Dim sDept           As String
    Dim sWard           As String
    Dim sSample         As String
    Dim iRowCheck       As Integer
    Dim iMaxRow         As Integer
    Dim sRemark         As String
    Dim sMDate          As String
    
    
    
    iRowCheck = 0
    iMaxRow = 0
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        If ssResult.Value = True Then
            iMaxRow = i
        End If
    Next
    
    
    GoSub Clear_Sheet
    GoSub DataCheck
    'GoSub Print_Spread
    
    Exit Sub
    


Clear_Sheet:
    Dim iCls        As Integer
    
    sprSheet.Row = 2
    For iCls = 1 To 5
        sprSheet.Col = iCls: sprSheet.Text = ""
    Next
    
    sprSheet.Row = 7
    For iCls = 1 To 5
        sprSheet.Col = iCls: sprSheet.Text = ""
    Next
    
    sprSheet.Row = 9
    For iCls = 1 To 5
        sprSheet.Col = iCls: sprSheet.Text = ""
    Next
    sprSheet.Row = 14
    For iCls = 1 To 5
        sprSheet.Col = iCls: sprSheet.Text = ""
    Next
    
    Return


DataCheck:
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        If ssResult.Value = True Then
            
            ssResult.Col = 7: sMicrono = ssResult.Text
            
            ssResult.Col = 3: sJeobsuDt = ssResult.Text
            ssResult.Col = 4: sSLipno1 = ssResult.Text
            ssResult.Col = 5: sSLipno2 = ssResult.Text
            ssResult.Col = 8: sPtno = ssResult.Text
            ssResult.Col = 9: sName = ssResult.Text
            
            ssResult.Col = 12: sWard = ssResult.Text
            ssResult.Col = 15: sRemark = ssResult.Text
            ssResult.Col = 16: sSample = ssResult.Text
            ssResult.Col = 17: sDept = ssResult.Text
            ssResult.Col = 18: sMDate = ssResult.Text
            
            If iRowCheck = 0 Then GoSub Clear_Sheet
                
            If iRowCheck = 0 Then
                sprSheet.Col = 2
                sprSheet.Row = 7: sprSheet.Text = sRemark
                sprSheet.Row = 2: iRowCheck = sprSheet.Row
            ElseIf iRowCheck = 2 Then
                sprSheet.Col = 2
                sprSheet.Row = 14: sprSheet.Text = sRemark
                sprSheet.Row = 9: iRowCheck = sprSheet.Row
                
            End If
            
            
            sprSheet.Col = 1:  sprSheet.Text = sMDate & " : " & sMicrono & " : " & _
                                               sJeobsuDt & "," & sSLipno1 & "," & sSLipno2
            sprSheet.Col = 2: sprSheet.Text = sPtno
            sprSheet.Col = 3: sprSheet.Text = sName
            sprSheet.Col = 4: sprSheet.Text = Trim(sDept) & "/" & Trim(sWard)
            sprSheet.Col = 5: sprSheet.Text = sSample
            
            
            If iRowCheck = 9 Then GoSub Print_Spread: iRowCheck = 0
            
            If iRowCheck = 2 Then
                If i = iMaxRow Then
                    GoSub Print_Spread
                    Exit Sub
                End If
            End If
                
        End If
        
    Next
     Return
     
     
     
Print_Spread:
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    Dim iLine             As Integer
    
    For iLine = 1 To 80
        sBarLine = sBarLine & "━"
    Next
    
    
    SprCodeNm.Row = SprCodeNm.ActiveRow
    SprCodeNm.Col = 2: sItemName = SprCodeNm.Text
    
    If sprSheet.DataRowCnt = 0 Then Exit Sub
    
    
    strFont(0) = "/fn""바탕체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "배 양 검 사 대 장"
    'strHead(0) = "/f1" & "/c" & "Item별 접수내역 LIST" & " - " & sItemName
    strHead(1) = "/f2" & "/l" & "출력일 : " & Dual_Date_Get("yyyy-MM-dd")
    
    sprSheet.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                           strFont(1) + "/n" + strFont(1) + "/l" + sBarLine + strFont(1)
    sprSheet.PrintFooter = "/f2" & "/l" & strFont(1) + "/l" + sBarLine
    
    sprSheet.PrintMarginLeft = 900
    sprSheet.PrintMarginRight = 0
    sprSheet.PrintMarginTop = 1000
    sprSheet.PrintMarginBottom = 0
    sprSheet.PrintColHeaders = False
    sprSheet.PrintRowHeaders = False
    sprSheet.PrintBorder = True
    sprSheet.PrintColor = True
    sprSheet.PrintGrid = True
    sprSheet.PrintShadows = True
    sprSheet.PrintUseDataMax = False
    sprSheet.Row = 1
    sprSheet.Row2 = sprSheet.MaxRows
    sprSheet.Col = 1
    sprSheet.Col2 = sprSheet.MaxCols
    sprSheet.PrintOrientation = 1
    sprSheet.PrintOrientation = PrintOrientationPortrait
    sprSheet.PrintType = PrintTypeCellRange
    sprSheet.Action = ActionPrint
    
    Return
    
    
End Sub

Private Sub cmdQuery_Click()
    Dim ii              As Integer


    Call SpreadSetClear(SprCodeNm)
    
    txtMSeq.Text = ""
    
    If CboPart.ListIndex = -1 Then Exit Sub
    
    FsExamGubun = Left$(CboPart.Text, 2)
    
    
    gStrSql = ""
    Select Case FsExamGubun
        Case "42":     '세균검사(+BLoodculture)
            gStrSql = gStrSql & "  SELECT CodeKy ItemCode,"
            gStrSql = gStrSql & "         Decode(itemNm, 'Blood Culture(1st)', 'BloodCulture', itemNm)  ItemName, "
            gStrSql = gStrSql & "         'I' ri"
            gStrSql = gStrSql & "  FROM   TWEXAM_ITEMML "
            gStrSql = gStrSql & "  WHERE  CodeKy   Like  '42%'"
            gStrSql = gStrSql & "  AND    SUBSTR(ItemNm, 1,2) != '  '"
            gStrSql = gStrSql & "  AND    Codeky NOT IN ('420403','420404')"         '420401=Blood Culture1
            gStrSql = gStrSql & "  ORDER  BY  CodeKy"
        Case "43":     '기생충검사
            gStrSql = gStrSql & "  SELECT distinct ROUTINCD ItemCode , ROUTINNM ItemName, 'R' ri "
            gStrSql = gStrSql & "  FROM   TWEXAM_ROUTINE"
            gStrSql = gStrSql & "  WHERE  ROUTINCD LIKE '43%'"
            gStrSql = gStrSql & "  UNION ALL "
            gStrSql = gStrSql & "  SELECT CODEKY ItemCode, ITEMNM ItemName, 'I' ri"
            gStrSql = gStrSql & "  FROM   TWEXAM_ITEMML"
            gStrSql = gStrSql & "  WHERE  CODEKY LIKE '43%'"
            gStrSql = gStrSql & "  Order By ItemCode"
        Case Else:
            gStrSql = gStrSql & " SELECT CodeKy ItemCode,itemNm ItemName, 'I' ri"
            gStrSql = gStrSql & " FROM   TWEXAM_ITEMML "
            gStrSql = gStrSql & " WHERE  CodeKy   Like  '" & FsExamGubun & "%'"
            gStrSql = gStrSql & " AND    SUBSTR(ItemNm, 1,2) != '  '"
            gStrSql = gStrSql & " ORDER  BY  CodeKy "
    End Select
    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        SprCodeNm.Row = SprCodeNm.DataRowCnt + 1
        
        SprCodeNm.Col = 2: SprCodeNm.Text = adoSet.Fields("itemName").Value & ""
        SprCodeNm.Col = 3: SprCodeNm.Text = adoSet.Fields("ItemCode").Value & ""
        SprCodeNm.Col = 9: SprCodeNm.Text = adoSet.Fields("ri").Value & ""

        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    

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
'    gStrSql = gStrSql & " AND    CODEKY < '52'    "
    gStrSql = gStrSql & " AND    SUBSTR(Codeky,1,1) = '4'"
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
    
    
    GiExamNumb = Val(GetSetting("CP", "CPRESULT", "SLip"))
    Call SetComboBox(CboPart, GiExamNumb, 2)


End Sub


Private Sub SprCodeNm_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim sItemCd     As String
    Dim sFromDate   As String
    Dim sToDate     As String
    Dim sGeomsaAB   As String
    Dim sRoutine    As String
    Dim sGeomchCD   As String
    
    
    ssResult.MaxRows = 0
    ssResult.MaxRows = 200
    If Trim(Me.txtRowh.Text) = "" Then
        ssResult.RowHeight(-1) = 11
    Else
        ssResult.RowHeight(-1) = Val(txtRowh.Text)
    End If
    
    
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
    
        
    sFromDate = Format(dtFdate.Value, "yyyy-MM-dd") & " 00:01"
    sToDate = Format(dtTdate.Value, "yyyy-MM-dd") & " 23:59"
    
    SprCodeNm.Row = Row
    SprCodeNm.Col = 3:  sItemCd = Trim(SprCodeNm.Text)
    SprCodeNm.Col = 9:
    If Trim(SprCodeNm.Text) = "R" Then
        sRoutine = "R"
    Else
        sRoutine = ""
    End If
    
    If Row = 0 Then Exit Sub
    If Col > 1 Then Exit Sub
    
    If Trim(sItemCd) = "" Then Exit Sub
        
    GoSub Get_GeneralSub
    'GoSub Default_Result_Set
    
    Exit Sub
    


Get_GeneralSub:
    Dim sBLoodculture       As String
    
    Call SpreadSetClear(ssResult)
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT DISTINCT a.RowiD RWID, a.*, TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    gStrSql = gStrSql & "        TO_CHAR(a.MDate, 'yyyy-MM-dd') MDate,"
    gStrSql = gStrSql & "        b.Sname, b.Sex, b.AgeYY, d.CmDoctor, e.Codenm SampleName, f.WardCode,"
    gStrSql = gStrSql & "        d.DeptCode"
    gStrSql = gStrSql & " FROM   TWEXAM_GENERAL_SUB a,"
    gStrSql = gStrSql & "        TWEXAM_IDnomst     b,"
    gStrSql = gStrSql & "        TWEXAM_General     c,"
    gStrSql = gStrSql & "        TWEXAM_Order       d,"
    gStrSql = gStrSql & "        TWEXAM_Sample      e,"
    gStrSql = gStrSql & "        TW_MIS_PMPA.TWBAS_Room         f "
    gStrSql = gStrSql & " WHERE  a.MDate >= TO_DATE('" & sFromDate & "','yyyy-MM-dd hh24:mi')"
    gStrSql = gStrSql & " AND    a.MDate <= TO_DATE('" & sToDate & "',  'yyyy-MM-dd hh24:mi')"
    
    
    If Trim(sItemCd) = "420402" Then   'BloodCulture
        gStrSql = gStrSql & " AND   a.ItemCd  IN ('420402','420403','420404')" 'BloodCulture 1,2,3
    Else
        If sRoutine = "R" Then
            gStrSql = gStrSql & " AND    a.RoutinCd = '" & sItemCd & "'"
        Else
            gStrSql = gStrSql & " AND    a.itemCD    =  '" & sItemCd & "'"
        End If
    End If
    
    If Trim(txtMSeq.Text) <> "" Then
        gStrSql = gStrSql & " AND    a.MSeq      > " & Val(txtMSeq.Text)
    End If
    
    gStrSql = gStrSql & " AND    a.JeobsuDt  = c.JeobsuDt(+)"
    gStrSql = gStrSql & " AND    a.SLipno1   = c.SLipno1(+)"
    gStrSql = gStrSql & " AND    a.SLipno2   = c.SLipno2(+)"
    gStrSql = gStrSql & " AND    a.GeomChCD  = e.Code(+)"
    
    If Option1.Value = True Then
        gStrSql = gStrSql & " AND a.Verify = 'Y'"
    Else
        gStrSql = gStrSql & " AND a.Verify <> 'Y'"
    End If
    'gStrSql = gStrSql & " AND    c.GbCh      = 'Y'"
    gStrSql = gStrSql & " AND    a.Ptno      = b.Ptno(+)"
    gStrSql = gStrSql & " AND    a.Ptno      = d.Ptno(+)"
    gStrSql = gStrSql & " AND    a.Orderno   = d.Orderno(+)"
    gStrSql = gStrSql & " AND    d.RoomCode  = f.RoomCode(+)"
    gStrSql = gStrSql & " ORDER  BY a.MSeq, a.MDate,a.SLipno2"
    
    If False = adoSetOpen(gStrSql, adoSet) Then Return
    
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("RWID").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("Slipno1").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("Slipno2").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("ItemCD").Value & ""
                                           
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("MSeq").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("AgeYY").Value & ""
    
        Select Case Trim(adoSet.Fields("ItemCd").Value & "")  'BLoodCulture Check
            Case "420402": sBLoodculture = "1) "      'BLoodCuture 1st
            Case "420403": sBLoodculture = "2) "      'BLoodCuture 2nd
            Case "420404": sBLoodculture = "3) "      'BLoodCuture 3rd
            Case Else:     sBLoodculture = ""
        End Select
        
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("WardCode").Value & ""
        
        ssResult.Col = 13: ssResult.Text = sBLoodculture & Trim$(adoSet.Fields("Result1").Value & "")
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("Verify").Value & ""
        ssResult.Col = 15: ssResult.Text = Trim$(adoSet.Fields("CmDoctor").Value & "")
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("SampleName").Value & ""
        ssResult.Col = 17: ssResult.Text = Trim(adoSet.Fields("DeptCode").Value & "")
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("MDate").Value & ""
        
        If isSensItem(Trim(adoSet.Fields("ItemCd").Value & "")) Then
            sGeomchCD = Trim(adoSet.Fields("GeomchCD").Value & "")
            GoSub Get_GramStain
        End If
        

        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

Get_GramStain:
    Dim adoGs       As ADODB.Recordset
    Dim cStrSql     As String
    
    
    cStrSql = ""
    cStrSql = cStrSql & " SELECT a.*, TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt, b.CmDoctor"
    cStrSql = cStrSql & " FROM   TWEXAM_GENERAL_SUB a,"
    cStrSql = cStrSql & "        TWEXAM_Order       b"
    cStrSql = cStrSql & " WHERE  a.JeobsuDt  = TO_DATE('" & adoSet.Fields("JeobsuDt").Value & "','yyyy-MM-dd')"
    cStrSql = cStrSql & " AND    a.SLipno1   =  " & Val(adoSet.Fields("Slipno1").Value & "")
    cStrSql = cStrSql & " AND    a.SLipno2   =  " & Val(adoSet.Fields("Slipno2").Value & "")
    cStrSql = cStrSql & " AND    a.itemCD    =  '420101'"
    cStrSql = cStrSql & " AND    a.GeomchCd  =  '" & sGeomchCD & "'"
    cStrSql = cStrSql & " AND    a.Ptno      = b.Ptno(+)"
    cStrSql = cStrSql & " AND    a.Orderno   = b.Orderno(+)"
    
    If False = adoSetOpen(cStrSql, adoGs) Then Return
    ssResult.Row = ssResult.DataRowCnt + 1
    ssResult.Col = 7:  ssResult.Text = adoGs.Fields("MSeq").Value & ""
    ssResult.Col = 8:  ssResult.Text = "GramStain"
    ssResult.Col = 14: ssResult.Text = adoSet.Fields("Verify").Value & ""
    ssResult.Col = 15: ssResult.Text = Trim$(adoSet.Fields("CmDoctor").Value & "")
    Call adoSetClose(adoGs)
    
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
        Case 2: 'Separator
    End Select
    
End Sub
