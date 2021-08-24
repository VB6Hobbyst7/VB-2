VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQryOrder 
   Caption         =   "Order 조회"
   ClientHeight    =   7710
   ClientLeft      =   330
   ClientTop       =   915
   ClientWidth     =   11535
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
   ScaleHeight     =   7710
   ScaleWidth      =   11535
   WindowState     =   2  '최대화
   Begin VB.Frame Frame3 
      Caption         =   "Order내역"
      Height          =   4785
      Left            =   4455
      TabIndex        =   10
      Top             =   2835
      Width           =   6945
      Begin FPSpreadADO.fpSpread sprOrder 
         Height          =   4425
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   6630
         _Version        =   196608
         _ExtentX        =   11695
         _ExtentY        =   7805
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
         MaxCols         =   4
         MaxRows         =   200
         ScrollBars      =   2
         SpreadDesigner  =   "frmQryOrder.frx":0000
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "외래진료 접수내역"
      Height          =   4785
      Left            =   135
      TabIndex        =   9
      Top             =   2835
      Width           =   4110
      Begin FPSpreadADO.fpSpread sprJupMst 
         Height          =   4425
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   3885
         _Version        =   196608
         _ExtentX        =   6853
         _ExtentY        =   7805
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
         GrayAreaBackColor=   16761024
         MaxCols         =   5
         MaxRows         =   20
         ScrollBars      =   2
         SpreadDesigner  =   "frmQryOrder.frx":1979
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "조회종류"
      Height          =   960
      Left            =   3915
      TabIndex        =   3
      Top             =   135
      Width           =   7485
      Begin VB.OptionButton Option2 
         Caption         =   "등록번호"
         Height          =   180
         Left            =   1350
         TabIndex        =   5
         Top             =   405
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "환자명"
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   405
         Width           =   915
      End
      Begin Threed.SSPanel panelPtno 
         Height          =   510
         Left            =   2745
         TabIndex        =   7
         Top             =   270
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   900
         _StockProps     =   15
         Caption         =   " 병록번호:"
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
         Alignment       =   1
         Begin VB.TextBox txtPtno 
            Height          =   330
            Left            =   990
            TabIndex        =   0
            Top             =   90
            Width           =   1230
         End
      End
      Begin Threed.SSPanel panelName 
         Height          =   510
         Left            =   2745
         TabIndex        =   6
         Top             =   270
         Visible         =   0   'False
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   900
         _StockProps     =   15
         Caption         =   " 환자명:"
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
         Alignment       =   1
         Begin VB.TextBox txtSname 
            Height          =   330
            Left            =   765
            TabIndex        =   1
            Top             =   90
            Width           =   1230
         End
      End
      Begin MSForms.CommandButton cmdQryOk 
         Height          =   510
         Left            =   5985
         TabIndex        =   8
         Top             =   240
         Width           =   1410
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2487;900"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread sprQryPt 
      Height          =   1545
      Left            =   135
      TabIndex        =   2
      Top             =   1170
      Width           =   11265
      _Version        =   196608
      _ExtentX        =   19870
      _ExtentY        =   2725
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
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryOrder.frx":1EEE
      Appearance      =   1
      ScrollBarTrack  =   1
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   465
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   820
      _StockProps     =   15
      Caption         =   "Order 조회"
      ForeColor       =   65535
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "궁서체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   2430
      TabIndex        =   14
      Top             =   720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36431
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   1080
      TabIndex        =   15
      Top             =   720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36431
   End
   Begin VB.Label Label1 
      Caption         =   "OrderDate"
      Height          =   195
      Left            =   135
      TabIndex        =   16
      Top             =   810
      Width           =   915
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQryOK_Click()
    
    Dim iCheckGubun As Integer
    
        
    If Option2.Value = True Then
        If Trim(txtPtno.Text) = "" Then Exit Sub
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        iCheckGubun = 1
    End If
    
    
    If Option1.Value = True Then
        If Trim(txtSname.Text) = "" Then Exit Sub
        iCheckGubun = 2
    End If
    
    Call Spread_Set_Clear(sprQryPt)
    
    Select Case iCheckGubun
        Case 1               '병록번호조회
            'strSql = ""
            'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */  "
            
            strSql = ""
            strSql = strSql & " SELECT a.*, b.DeptNameK, c.Drname,"
            strSql = strSql & "        d.PostName1, d.PostName2, d.PostName3"
            strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
            strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT    b,"
            strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  c,"
            strSql = strSql & "        TW_MIS_PMPA.TWBas_POST    d "
            strSql = strSql & " WHERE  a.Ptno      = '" & txtPtno.Text & "'"
            strSql = strSql & " AND    a.DeptCode  = b.DeptCode(+)"
            strSql = strSql & " AND    a.DrCode    = c.DrCode(+)"
            strSql = strSql & " AND    a.PostCode1 = d.PostCode1(+)"
            strSql = strSql & " AND    a.PostCode2 = d.PostCode2(+)"
        Case 2               '수진자명 조회
            'strSql = ""
            'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT1) */ "
            
            strSql = ""
            strSql = strSql & " SELECT a.*, b.DeptNameK, c.Drname,"
            strSql = strSql & "        d.PostName1, d.PostName2, d.PostName3"
            strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
            strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT    b,"
            strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  c,"
            strSql = strSql & "        TW_MIS_PMPA.TWBas_POST    d "
            strSql = strSql & " WHERE  a.Sname     Like  '" & txtSname.Text & "%'"
            strSql = strSql & " AND    a.DeptCode  =     b.DeptCode(+)"
            strSql = strSql & " AND    a.DrCode    =     c.DrCode(+)"
            strSql = strSql & " AND    a.PostCode1 =     d.PostCode1(+)"
            strSql = strSql & " AND    a.PostCode2 =     d.PostCode2(+)"
    End Select
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprQryPt.Row = sprQryPt.DataRowCnt + 1
        sprQryPt.Col = 1:  sprQryPt.CellType = CellTypeButton
                           sprQryPt.TypeButtonText = ""
        sprQryPt.Col = 2:  sprQryPt.Text = adoSet.Fields("Ptno").Value & ""
        sprQryPt.Col = 3:  sprQryPt.Text = adoSet.Fields("Sname").Value & ""
        sprQryPt.Col = 4:  sprQryPt.Text = adoSet.Fields("Sex").Value & "/" & _
                                           SetAge_Check(adoSet.Fields("Jumin1").Value & "", adoSet.Fields("Jumin2").Value & "")
        sprQryPt.Col = 5:  sprQryPt.Text = adoSet.Fields("Jumin1").Value & "-" & adoSet.Fields("Jumin2").Value & ""
        
        sprQryPt.Col = 6:  sprQryPt.Text = adoSet.Fields("Tel").Value & ""
        sprQryPt.Col = 7:  sprQryPt.Text = adoSet.Fields("LastDate").Value & ""
        sprQryPt.Col = 8:  sprQryPt.Text = adoSet.Fields("DeptNameK").Value & ""
        sprQryPt.Col = 9:  sprQryPt.Text = adoSet.Fields("Drname").Value & ""
        sprQryPt.Col = 10: sprQryPt.Text = Trim(adoSet.Fields("Postname1").Value & "") & " " & _
                                           Trim(adoSet.Fields("Postname2").Value & "") & " " & _
                                           Trim(adoSet.Fields("Postname3").Value & "") & " " & _
                                           Trim(adoSet.Fields("Juso").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    If sprQryPt.DataRowCnt = 1 Then
        Call sprQryPt_ButtonClicked(1, 1, 1)
    End If
    
End Sub

Private Sub Form_Activate()
    
    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_Load()
    
    Call Spread_Set_Clear(Me.sprQryPt)
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtFrDate.Value = Dual_Date_Cal_Get("yyyy-MM-dd", -30)
    
        
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub Option1_Click()
    
    If Option1.Value = True Then
        panelName.Visible = True
        panelPtno.Visible = False
        panelName.ZOrder 0
        txtSname.SetFocus
    End If
        
End Sub

Private Sub Option2_Click()
    
    If Option2.Value = True Then
        panelName.Visible = False
        panelPtno.Visible = True
        panelPtno.ZOrder 0
        txtPtno.SetFocus
    End If

End Sub

Private Sub sprJupMst_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Row > 0 Then
        If Col = 4 Then
            Call sprJupMst_DblClick(1, Row)
        End If
    End If
    
End Sub

Private Sub sprJupMst_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sQryPtno        As String
    Dim sBDate          As String
    Dim sDeptC          As String
    
    
    If Row = 0 Then Exit Sub
    
    
    sprQryPt.Row = sprQryPt.ActiveRow
    sprQryPt.Col = 2
    sQryPtno = sprQryPt.Text

    sprJupMst.Row = Row
    sprJupMst.Col = 1
    sBDate = sprJupMst.Text
    
    sprJupMst.Col = 5: sDeptC = sprJupMst.Text   '과코드
    
    If Trim(sQryPtno) = "" Or Trim(sBDate) = "" Then
        Exit Sub
    End If
    
    
    GoSub Get_OOrder:
    Exit Sub
    
    
    
Get_OOrder:
    strSql = ""
    strSql = strSql & "   SELECT TO_CHAR(a.Bdate, 'YYYY-MM-DD') Bdate,                  " & vbLf
    strSql = strSql & "          b.OrdernameS, a.Gbsunap, a.Seqno, c.Deptnamek           " & vbLf
    strSql = strSql & "   FROM   TW_MIS_OCS.TWOCS_oorder    a,                          " & vbLf
    strSql = strSql & "          TW_MIS_OCS.TWOCS_Ordercode b,                          " & vbLf
    strSql = strSql & "          TW_MIS_PMPA.TWBAS_DEPT      c                          " & vbLf
    strSql = strSql & "   WHERE  a.Bdate     =  TO_DATE('" & sBDate & "','YYYY-MM-DD')  " & vbLf
    strSql = strSql & "   AND    a.Ptno      = '" & sQryPtno & "'                       " & vbLf
    strSql = strSql & "   AND    a.DeptCode  = '" & sDeptC & "'                         " & vbLf
    strSql = strSql & "   AND    a.Bun      >=  '56'                                    " & vbLf
    strSql = strSql & "   AND    a.Bun      <   '70'                                    " & vbLf
    strSql = strSql & "   AND    a.SLipno    = b.SLipno(+)                              " & vbLf
    strSql = strSql & "   AND    a.DeptCode  = c.DeptCode(+)                            " & vbLf
    strSql = strSql & "   AND    a.OrderCode = b.OrderCode(+)                           "
    strSql = strSql & "   ORDER  BY a.Bdate DESC , a.Seqno                              "
    
    Call Spread_Set_Clear(sprOrder)
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        If i > 50 Then Exit Do
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1: sprOrder.Text = adoSet.Fields("Bdate").Value & ""
        sprOrder.Col = 2: sprOrder.Text = adoSet.Fields("DeptnameK").Value & ""
        sprOrder.Col = 3: sprOrder.Text = adoSet.Fields("OrderNameS").Value & ""
        sprOrder.Col = 4:
        Select Case adoSet.Fields("GbSunap").Value & ""
            Case "1": sprOrder.Text = "O"       '수납
            Case "0": sprOrder.Text = "X"       '미수납
        End Select
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
End Sub

Private Sub sprQryPt_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sQryPtno        As String
    Dim sFrDate         As String
    Dim sToDate         As String
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If Row = 0 Then Exit Sub
    
    sprQryPt.Row = Row
    sprQryPt.Col = 2
    sQryPtno = sprQryPt.Text
    
    
    GoSub Get_JupMaster
    Exit Sub
    
    
Get_JupMaster:
    Dim sActDate        As String * 10
    Dim sDeptNamek      As String * 20
    Dim sDrName         As String * 12
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWOPD_JUPMST INDEX_JUPMST1) */"
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.actDate,'YYYY-MM-DD') ActDate, "
    strSql = strSql & "        b.DeptNamek, c.Drname, a.DeptCode"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWOPD_JUPMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT   b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR c"
    strSql = strSql & " WHERE  a.Ptno      = '" & sQryPtno & "'"
    strSql = strSql & " AND    a.ActDate  >= To_Date('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.ActDate  <= To_Date('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.DelMark  <> '*'"
    strSql = strSql & " AND    a.DeptCode  = b.DeptCode(+)"
    strSql = strSql & " AND    a.Drcode    = c.Drcode(+)"
    strSql = strSql & " Order By a.ACTDate DESC, a.DeptCode"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Call Spread_Set_Clear(sprJupMst)
    
    Do Until adoSet.EOF
        sprJupMst.Row = sprJupMst.DataRowCnt + 1
        
        sprJupMst.Col = 1: sprJupMst.Text = adoSet.Fields("ActDate").Value & ""
        sprJupMst.Col = 2: sprJupMst.Text = adoSet.Fields("DeptNameK").Value & ""
        sprJupMst.Col = 3: sprJupMst.Text = adoSet.Fields("Drname").Value & ""
        sprJupMst.Col = 4: sprJupMst.CellType = CellTypeButton
                           sprJupMst.TypeButtonText = "▶"
        sprJupMst.Col = 5: sprJupMst.Text = adoSet.Fields("DeptCode").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
End Sub

Private Sub txtPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Trim(txtPtno.Text) <> "" Then
            txtPtno.Text = Format(txtPtno.Text, "00000000")
            cmdQryOK.SetFocus
        End If
    End If

End Sub

Private Sub txtSname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQryOK.SetFocus
    End If
End Sub
