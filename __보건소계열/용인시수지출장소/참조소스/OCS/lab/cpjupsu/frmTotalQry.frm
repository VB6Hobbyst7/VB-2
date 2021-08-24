VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTotalQry 
   Caption         =   "접수내역 조회"
   ClientHeight    =   8145
   ClientLeft      =   300
   ClientTop       =   540
   ClientWidth     =   11355
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
   ScaleHeight     =   8145
   ScaleWidth      =   11355
   WindowState     =   2  '최대화
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   2250
      TabIndex        =   7
      Top             =   630
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36584
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   540
      TabIndex        =   6
      Top             =   630
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36584
   End
   Begin FPSpreadADO.fpSpread sprGeneral 
      Height          =   1095
      Left            =   225
      TabIndex        =   5
      Top             =   4050
      Width           =   6270
      _Version        =   196608
      _ExtentX        =   11060
      _ExtentY        =   1931
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
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmTotalQry.frx":0000
      UserResize      =   0
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread sprGeneralSub 
      Height          =   3300
      Left            =   6525
      TabIndex        =   4
      Top             =   4050
      Width           =   3525
      _Version        =   196608
      _ExtentX        =   6218
      _ExtentY        =   5821
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
      MaxCols         =   2
      ScrollBars      =   2
      SpreadDesigner  =   "frmTotalQry.frx":190E
      UserResize      =   1
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread sprOrder 
      Height          =   2850
      Left            =   225
      TabIndex        =   3
      Top             =   1170
      Width           =   11580
      _Version        =   196608
      _ExtentX        =   20426
      _ExtentY        =   5027
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
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
      MaxCols         =   15
      MaxRows         =   300
      ScrollBars      =   2
      SpreadDesigner  =   "frmTotalQry.frx":3110
      UserResize      =   0
      Appearance      =   1
   End
   Begin VB.TextBox txtPtno 
      Height          =   285
      Left            =   4860
      TabIndex        =   2
      Top             =   675
      Width           =   1365
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10260
      Top             =   450
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
            Picture         =   "frmTotalQry.frx":432C
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
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   635
      ButtonWidth     =   1270
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
            Object.ToolTipText     =   "Unload-me Screen"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   465
      Left            =   6255
      TabIndex        =   8
      Top             =   675
      Width           =   1410
      Caption         =   "조회확인"
      Size            =   "2487;820"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "등록번호"
      Height          =   195
      Left            =   4005
      TabIndex        =   1
      Top             =   720
      Width           =   780
   End
End
Attribute VB_Name = "frmTotalQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuery_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call Spread_Set_Clear(sprOrder)
    Call Spread_Set_Clear(sprGeneral)
    Call Spread_Set_Clear(sprGeneralSub)
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT "
    strSql = strSql & "        a.*, a.ROWID OdrRow,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        TO_CHAR(a.EntTime, 'yyyy-MM-dd hh24:mi') EntTime,"
    strSql = strSql & "        TO_CHAR(a.CollDate,'yyyy-MM-dd') COLLDate,"
    strSql = strSql & "        TO_CHAR(a.GBDate,  'yyyy-MM-dd hh24:mi') GBDate,"
    strSql = strSql & "        b.RoutinNM ItemName"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine b "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
'    strSql = strSql & " AND    a.SLipno1   BETWEEN 0 AND  52"
    strSql = strSql & " AND    a.SLipno1   BETWEEN 0 AND  90    "
    strSql = strSql & " AND    a.Ptno      = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.ItemCd    = b.RoutinCd"
    strSql = strSql & " UNION  ALL"
    strSql = strSql & " SELECT a.*, a.RowID OdrRow,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        TO_CHAR(a.EntTime, 'yyyy-MM-dd hh24:mi') EntTime,"
    strSql = strSql & "        TO_CHAR(a.CollDate,'yyyy-MM-dd') COLLDate,"
    strSql = strSql & "        TO_CHAR(a.GBDate,  'yyyy-MM-dd hh24:mi') GBDate,"
    strSql = strSql & "        b.iTemNM ItemName"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML  b "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
'C    strSql = strSql & " AND    a.SLipno1   BETWEEN 0 AND 52"
    strSql = strSql & " AND    a.SLipno1   BETWEEN 0 AND 90 "
    strSql = strSql & " AND    a.Ptno      = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.ItemCd    = b.Codeky"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
        
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1:  sprOrder.Text = adoSet.Fields("OdrRow").Value & ""
        sprOrder.Col = 2:  sprOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprOrder.Col = 3:  sprOrder.Text = adoSet.Fields("SLipno1").Value & ""
        sprOrder.Col = 4:  sprOrder.Text = adoSet.Fields("Orderno").Value & ""
        sprOrder.Col = 5:  sprOrder.Text = adoSet.Fields("ItemName").Value & ""
        sprOrder.Col = 6:  sprOrder.Text = adoSet.Fields("OrderGb").Value & ""
        sprOrder.Col = 7:  sprOrder.Text = adoSet.Fields("RoomCode").Value & ""
        sprOrder.Col = 8:  sprOrder.Text = adoSet.Fields("DeptCode").Value & ""
        sprOrder.Col = 9:  sprOrder.Text = adoSet.Fields("DrCode").Value & ""
        sprOrder.Col = 10: sprOrder.Text = adoSet.Fields("CollDate").Value & ""
        sprOrder.Col = 11: sprOrder.Text = adoSet.Fields("CollID").Value & ""
        sprOrder.Col = 12: sprOrder.Text = adoSet.Fields("JeobsuYN").Value & ""
        sprOrder.Col = 13: sprOrder.Text = adoSet.Fields("GBCh").Value & ""
        sprOrder.Col = 14: sprOrder.Text = adoSet.Fields("GBDate").Value & ""
        sprOrder.Col = 15: sprOrder.Text = adoSet.Fields("Matchno").Value & ""

        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    

    
End Sub

Private Sub Form_Load()
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub sprOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    Dim iSLipno2        As Integer
    Dim iMatchno        As Integer
    Dim iOrderno        As Long
    Dim sPtno           As String
    Dim sCollDate       As String

    
    If sprOrder.Row > sprOrder.DataRowCnt Then Exit Sub
    If Row = 0 Then Exit Sub
    
    
    sprOrder.Row = Row
    sprOrder.Col = 2: sJeobsuDt = sprOrder.Text
    sprOrder.Col = 3: iSLipno1 = Val(sprOrder.Text)
    sprOrder.Col = 4: iOrderno = Val(sprOrder.Text)
    sprOrder.Col = 10: sCollDate = sprOrder.Text
    sprOrder.Col = 15: iMatchno = Val(sprOrder.Text)
        
    GoSub Get_General
    GoSub Get_General_Sub
    Exit Sub
    


Get_General:
    Call Spread_Set_Clear(sprGeneral)
    
    strSql = ""
    strSql = strSql & " SELECT a.*, "
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        TO_CHAR(a.GbDate, 'yyyy-MM-dd hh24:mi') GBDate"
    strSql = strSql & " FROM   TWEXAM_General a"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sCollDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.Matchno  = " & iMatchno
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprGeneral.Row = sprGeneral.DataRowCnt + 1
        sprGeneral.Col = 1: sprGeneral.Text = adoSet.Fields("Jeobsudt").Value & ""
        sprGeneral.Col = 2: sprGeneral.Text = adoSet.Fields("SLipno1").Value & ""
        sprGeneral.Col = 3: sprGeneral.Text = adoSet.Fields("SLipno2").Value & ""
        sprGeneral.Col = 4: sprGeneral.Text = adoSet.Fields("GBCH").Value & ""
        sprGeneral.Col = 5: sprGeneral.Text = adoSet.Fields("GBDate").Value & ""
        sprGeneral.Col = 6: sprGeneral.Text = adoSet.Fields("Status").Value & ""
        sprGeneral.Col = 7: sprGeneral.Text = adoSet.Fields("Matchno").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
Get_General_Sub:
    Call Spread_Set_Clear(sprGeneralSub)
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.ItemNm ItemName"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sCollDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.Matchno  = " & iMatchno
    strSql = strSql & " AND    a.Orderno  = " & iOrderno
    strSql = strSql & " AND    a.ItemCd   = b.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprGeneralSub.Row = sprGeneralSub.DataRowCnt + 1
        sprGeneralSub.Col = 1: sprGeneralSub.Text = adoSet.Fields("ItemName").Value & ""
        sprGeneralSub.Col = 2: sprGeneralSub.Text = adoSet.Fields("Verify").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
    
End Sub

Private Sub txtPtno_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        GoSub Yes_PatientNo
        cmdQuery.SetFocus
    End If
    Exit Sub
    
    

Yes_PatientNo:
    If Trim(txtPtno.Text) = "" Then Exit Sub
    txtPtno.Text = Format(txtPtno.Text, "00000000")
    
    Return
    
End Sub
