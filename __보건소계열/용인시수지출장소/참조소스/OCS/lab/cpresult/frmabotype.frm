VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmABOType 
   Caption         =   "ABOType & RhType 입력화면"
   ClientHeight    =   7695
   ClientLeft      =   5160
   ClientTop       =   975
   ClientWidth     =   6570
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6570
   Begin Threed.SSPanel SSPanel1 
      Height          =   510
      Left            =   1485
      TabIndex        =   7
      Top             =   450
      Width           =   4830
      _Version        =   65536
      _ExtentX        =   8520
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSForms.CommandButton cmdPrint 
         Height          =   420
         Left            =   1530
         TabIndex        =   10
         Top             =   45
         Width           =   1410
         Caption         =   "출력"
         PicturePosition =   327683
         Size            =   "2487;741"
         Picture         =   "frmABOType.frx":0000
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   420
         Left            =   135
         TabIndex        =   9
         Top             =   45
         Width           =   1410
         Caption         =   "조회"
         PicturePosition =   327683
         Size            =   "2487;741"
         Picture         =   "frmABOType.frx":08DA
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdVerify 
         Height          =   420
         Left            =   3195
         TabIndex        =   8
         Top             =   45
         Width           =   1410
         Caption         =   "입력"
         PicturePosition =   327683
         Size            =   "2487;741"
         Picture         =   "frmABOType.frx":11BC
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread sprABO 
      Height          =   6450
      Left            =   225
      TabIndex        =   1
      Top             =   990
      Width           =   6135
      _Version        =   196608
      _ExtentX        =   10821
      _ExtentY        =   11377
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      ScrollBars      =   2
      SpreadDesigner  =   "frmABOType.frx":297E
      UserResize      =   1
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6570
      _ExtentX        =   11589
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
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.OptionButton Option1 
         Caption         =   "Vr"
         Height          =   225
         Left            =   4770
         TabIndex        =   6
         Top             =   45
         Width           =   600
      End
      Begin VB.OptionButton Option2 
         Caption         =   "NoVr"
         Height          =   225
         Left            =   5445
         TabIndex        =   5
         Top             =   45
         Value           =   -1  'True
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36544
      End
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   315
         Left            =   2970
         TabIndex        =   3
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36544
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   1170
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
            Picture         =   "frmABOType.frx":433E
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdSelect 
      Height          =   375
      Left            =   225
      TabIndex        =   2
      Top             =   585
      Width           =   1140
      Caption         =   "▼전체선택"
      Size            =   "2011;661"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmABOType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCheckRH            As String

Private Sub cmdPrint_Click()
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    Dim sBarLine          As String
    Dim sItemName         As String
    
    
    For i = 1 To 60
        sBarLine = sBarLine & "━"
    Next
    
    If sprABO.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread의 Data Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & "Item별 접수내역 LIST" & " - " & "(ABO & Rh Type)"
    
    strHead(1) = "/f2" & "접수일자(Fr/To): " & Format(dtFrDate.Value, "yyyy-MM-dd") & " / " & _
                                               Format(dtToDate.Value, "yyyy-MM-dd")
    
    sprABO.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + _
                            strFont(1) + "/n" + sBarLine + strFont(1)
    sprABO.PrintFooter = "/f2" & "/l" & sBarLine & _
                            "/n" & Space(80) & "Page : " & "/p" & " of " & sprABO.PrintPageCount
    sprABO.PrintMarginLeft = 0
    sprABO.PrintMarginRight = 0
    sprABO.PrintMarginTop = 0
    sprABO.PrintMarginBottom = 0
    sprABO.PrintColHeaders = True
    sprABO.PrintRowHeaders = True
    sprABO.PrintBorder = False
    sprABO.PrintColor = False
    sprABO.PrintGrid = True
    sprABO.PrintShadows = True
    sprABO.PrintUseDataMax = False
    sprABO.Row = 1
    sprABO.Row2 = sprABO.DataRowCnt
    sprABO.Col = 2
    sprABO.Col2 = sprABO.MaxCols
    sprABO.PrintOrientation = 1
    sprABO.PrintOrientation = PrintOrientationPortrait
    sprABO.PrintType = PrintTypeCellRange
    sprABO.Action = ActionPrint


End Sub

Private Sub cmdQuery_Click()
    Dim sVr         As String
    Dim sFrDate     As String
    Dim sToDate     As String
    
    
    
    Call SpreadSetClear(sprABO)
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If Option1.Value = True Then sVr = "Y"
    If Option2.Value = True Then sVr = "N"
    
    strSql = ""
    strSql = strSql & " SELECT JeobsuDt1, Ptno, Sname, Sx, age, sLipno1, slipno2,"
    strSql = strSql & "        MAX(DECODE(LTRIM(rtrim(ITEMCD)), '510101', RESULT1, '')) abo,"
    strSql = strSql & "        MAX(DECODE(LTRIM(rtrim(ITEMCD)), '510102', RESULT1, '')) rh"
    strSql = strSql & " FROM(  SELECT a.*, b.Sname, b.Sex Sx, b.AgeYY age,"
    strSql = strSql & "               TO_CHAR(a.Jeobsudt, 'yyyy-MM-dd') JeobsuDt1"
    strSql = strSql & "        FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "               TWEXAM_IDNOMST     b,"
    strSql = strSql & "               TWEXAM_GENERAL     c "
    strSql = strSql & "        WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & "        AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & "        AND    a.ITEMCD   IN ('510101','510102')"
    strSql = strSql & "        AND    a.VERIFY    = '" & sVr & "'"
    strSql = strSql & "        AND    a.SLIPNO1   = 51"
    strSql = strSql & "        AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & "        AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "        AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "        AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "        AND    c.GBCh      = 'Y')"
    strSql = strSql & " GROUP BY JeobsuDt1, Ptno, Sname, Sx, age, sLipno1, slipno2"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        sprABO.Row = sprABO.DataRowCnt + 1
        sprABO.Col = 2: sprABO.Text = adoSet.Fields("JeobsuDt1").Value & ""
        sprABO.Col = 3: sprABO.Text = adoSet.Fields("Ptno").Value & ""
        sprABO.Col = 4: sprABO.Text = adoSet.Fields("Sname").Value & ""
        sprABO.Col = 5: sprABO.Text = adoSet.Fields("Sx").Value & ""
        sprABO.Col = 6: sprABO.Text = adoSet.Fields("age").Value & ""
        sprABO.Col = 7: sprABO.Text = adoSet.Fields("slipno2").Value & ""
        sprABO.Col = 8: sprABO.Text = adoSet.Fields("abo").Value & ""
        sprABO.Col = 9: sprABO.Text = adoSet.Fields("rh").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub cmdSelect_Click()
    
    If cmdSelect.Caption = "▼전체선택" Then
        For i = 1 To sprABO.DataRowCnt
            sprABO.Row = i
            sprABO.Col = 1
            sprABO.Value = True
        Next
        cmdSelect.Caption = "▼전체해제"
    Else
        For i = 1 To sprABO.DataRowCnt
            sprABO.Row = i
            sprABO.Col = 1
            sprABO.Value = False
        Next
        cmdSelect.Caption = "▼전체선택"
    End If

End Sub

Private Sub cmdVerify_Click()
    Dim sJeobsuDt       As String
    Dim sSLno1          As String
    Dim sSLno2          As String
    Dim sItemCd         As String
    Dim sResult1        As String
    Dim sRowID          As String
    Dim sAboRet         As String
    Dim sRhRet          As String
    Dim sPtno           As String
    
    
    If sprABO.DataRowCnt = 0 Then Exit Sub
    
    GoSub Set_General_Sub
    If vbYes = MsgBox("결과를 Verify 했습니다. 상태를 결과완료로 Setting 시키겠습니까? ", _
                       vbYesNo + vbQuestion, _
                      "Status Set") Then GoSub Set_General
    
    Exit Sub
    
    
    
Set_General_Sub:
    For i = 1 To sprABO.DataRowCnt
        sprABO.Row = i
        sprABO.Col = 1
        If sprABO.Value = True Then
            sprABO.Col = 2: sJeobsuDt = sprABO.Text
            sprABO.Col = 3: sPtno = sprABO.Text
                            sSLno1 = "51"
            sprABO.Col = 7: sSLno2 = sprABO.Text
            sprABO.Col = 8: sAboRet = sprABO.Text
            sprABO.Col = 9: sRhRet = sprABO.Text
            
            If Trim(sSLno2) <> "" Then
                strSql = ""
                strSql = strSql & " UPDATE TWEXAM_General_Sub"
                strSql = strSql & " SET   Result1  = '" & Quot_Conv(sAboRet) & "',"
                strSql = strSql & "       Verify   = 'Y'"
                strSql = strSql & " WHERE JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
                strSql = strSql & " AND   SLipno1  = 51"
                strSql = strSql & " AND   SLipno2  = " & Val(sSLno2)
                strSql = strSql & " AND   ItemCD   = '510101'"
                adoConnect.BeginTrans
                If adoExec(strSql) Then
                    adoConnect.CommitTrans
                Else
                    adoConnect.RollbackTrans
                End If
                
                
                strSql = ""
                strSql = strSql & " UPDATE TWEXAM_General_Sub"
                strSql = strSql & " SET   Result1  = '" & Quot_Conv(sRhRet) & "',"
                strSql = strSql & "       Verify   = 'Y'"
                strSql = strSql & " WHERE JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
                strSql = strSql & " AND   SLipno1  = 51"
                strSql = strSql & " AND   SLipno2  = " & Val(sSLno2)
                strSql = strSql & " AND   ItemCD   = '510102'"
                
                adoConnect.BeginTrans
                If adoExec(strSql) Then
                    adoConnect.CommitTrans
                Else
                    adoConnect.RollbackTrans
                End If
            End If
            
        End If
    Next
    Return
    
Set_General:
    
    For i = 1 To sprABO.DataRowCnt
        sprABO.Row = i
        sprABO.Col = 1
        If sprABO.Value = True Then
            sprABO.Col = 2: sJeobsuDt = sprABO.Text
            sprABO.Col = 3: sPtno = sprABO.Text
                            sSLno1 = "51"
            sprABO.Col = 7: sSLno2 = sprABO.Text
            sprABO.Col = 8: sAboRet = sprABO.Text
            sprABO.Col = 9: sRhRet = sprABO.Text
            If Trim(sSLno2) <> "" Then
                strSql = ""
                strSql = strSql & " UPDATE TWEXAM_General"
                strSql = strSql & " SET    Status   = 'C'"
                strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
                strSql = strSql & " AND    SLipno1  = 51"
                strSql = strSql & " AND    SLipno2  = " & Val(sSLno2)
                adoConnect.BeginTrans
                If adoExec(strSql) Then
                    adoConnect.CommitTrans
                Else
                    adoConnect.RollbackTrans
                End If
            End If
        End If
    Next
    Return
    

End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
End Sub

Private Sub sprABO_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    
    
    If Row = 0 Then
        If Col = 9 Then
            GoSub SetRHType
        End If
    End If
    
    If Col = 9 And Row > 0 Then
        sprABO.Row = Row
        sprABO.Col = 9
        Select Case Trim(sprABO.Text)
            Case "+":  sprABO.Text = "-"
            Case "-":  sprABO.Text = "+"
            Case Else: sprABO.Text = "+"
        End Select
    End If
    
    Exit Sub
    
SetRHType:
    For i = 1 To sprABO.DataRowCnt
        sprABO.Row = i
        sprABO.Col = 9
        
        Select Case sCheckRH
            Case "+":  sprABO.Text = "-"
            Case "-":  sprABO.Text = "+"
            Case Else: sprABO.Text = "+"
        End Select
    Next
    
    Select Case sCheckRH
        Case "+":  sCheckRH = "-"
        Case "-":  sCheckRH = "+"
        Case Else: sCheckRH = "+"
    End Select
    
    Return
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub
